import streamlit as st
import os
import json
import requests
import hashlib
import base64
from datetime import datetime, timedelta
import pandas as pd
import io
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gcl
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ── FLEXIBLE COLUMN MAPPING ────────────────────────────────────
# Maps whatever headers a coach's Hudl export uses to the standard
# names the analysis expects. Case / spacing / punctuation tolerant.
# DefensiveIQ reads OFFENSIVE film (the opponent on offense) so a
# defensive coordinator can build a game plan against their looks.
COLUMN_ALIASES = {
    "DN":          ["DN", "DOWN", "DWN"],
    "DIST":        ["DIST", "DISTANCE", "DIS", "TO GO", "TOGO"],
    "HASH":        ["HASH", "HASH MARK"],
    "YARD LN":     ["YARD LN", "YARD LINE", "YRDLN", "YARDLINE", "YD LN", "LOS", "FIELD POS"],
    "OFF FORM":    ["OFF FORM", "OFFENSIVE FORMATION", "OFF FORMATION", "FORMATION", "FORM"],
    "OFF PLAY":    ["OFF PLAY", "PLAY", "PLAY NAME", "PLAY CALL", "CONCEPT", "PASS CONCEPT"],
    "PLAY TYPE":   ["PLAY TYPE", "RUN/PASS", "R/P", "TYPE", "RUNPASS"],
    "PLAY DIR":    ["PLAY DIR", "DIRECTION", "DIR", "PLAY DIRECTION"],
    "GN/LS":       ["GN/LS", "GAIN/LOSS", "GAIN", "YARDS", "YDS", "GN LS", "GAINLOSS"],
    "RESULT":      ["RESULT", "RES", "OUTCOME"],
    "QTR":         ["QTR", "QUARTER", "QT", "Q"],
    "PERSONNEL":   ["PERSONNEL", "PERS", "PERSONEL", "GROUPING"],
    "BACKFIELD":   ["BACKFIELD", "BACK FIELD", "BACKFIELD SET", "BACK SET"],
    "OFF STR":     ["OFF STR", "STRENGTH", "OFFENSIVE STRENGTH", "STR"],
    "MOTION":      ["MOTION DIR", "MOTION", "MOT", "MOTION DIRECTION"],
    "ODK":         ["ODK"],
    "EFF":         ["EFF", "EFFICIENT", "EFFICIENCY"],
    "PENALTY":     ["PENALTY", "PEN"],
}

def _normalize(s):
    return "".join(ch for ch in str(s).upper().strip() if ch.isalnum() or ch == " ").strip()

def map_columns(df):
    """Rename incoming columns to standard names. Returns (df, matched, missing)."""
    norm_incoming = {_normalize(c): c for c in df.columns}
    rename, matched = {}, {}
    for standard, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            na = _normalize(alias)
            if na in norm_incoming:
                original = norm_incoming[na]
                rename[original] = standard
                matched[standard] = original
                break
    df = df.rename(columns=rename)
    # If two different original columns both map to the same standard name,
    # the rename can create duplicates. Downstream code assumes one column
    # per standard name, so keep the first and drop the rest.
    df = df.loc[:, ~df.columns.duplicated()]
    missing = [s for s in COLUMN_ALIASES if s not in matched]
    return df, matched, missing

# Columns the analysis genuinely cannot run without
REQUIRED_COLS = ["PLAY TYPE", "YARD LN", "DN", "DIST"]

def check_required(matched):
    """Return list of required standard columns that were not found."""
    return [c for c in REQUIRED_COLS if c not in matched]

st.set_page_config(
    page_title="DefensiveIQ — Offensive Tendency Report",
    page_icon="🛡️",
    layout="wide",
)

# --- Access Gate (self-service signup + entitlement check) ---------------------------------
_GH_TOKEN = os.environ.get("GITHUB_TOKEN", "")
_GH_REPO = os.environ.get("GITHUB_REPO", "")
_GH_API = f"https://api.github.com/repos/{_GH_REPO}/contents/customers.json"

_SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
_SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")
_ALLOWED_PLANS = ["Defensive IQ", "Founders Plan", "Standard Plan"]

_SMTP_HOST = os.environ.get("SMTP_HOST", "")
_SMTP_PORT = int(os.environ.get("SMTP_PORT", "587") or "587")
_SMTP_USER = os.environ.get("SMTP_USER", "")
_SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")
_SMTP_FROM = os.environ.get("SMTP_FROM", _SMTP_USER)

def _gh_headers():
    return {"Authorization": f"token {_GH_TOKEN}", "Accept": "application/vnd.github+json"}

def _load_customers():
    """Load all customer records from Supabase. Returns (customers_dict, None)."""
    if not (_SUPABASE_URL and _SUPABASE_SERVICE_KEY):
        return {}, None
    try:
        r = requests.get(
            f"{_SUPABASE_URL}/rest/v1/accounts",
            headers={
                "apikey": _SUPABASE_SERVICE_KEY,
                "Authorization": f"Bearer {_SUPABASE_SERVICE_KEY}",
            },
            timeout=10,
        )
        if r.status_code != 200:
            return {}, None
        rows = r.json()
        customers = {}
        for row in rows:
            username = row.get("username")
            if not username:
                continue
            customers[username] = {
                "salt": row.get("salt"),
                "hash": row.get("hash"),
                "email": row.get("email"),
                "reset_hash": row.get("reset_hash"),
                "reset_salt": row.get("reset_salt"),
                "reset_expires": row.get("reset_expires"),
            }
        return customers, None
    except Exception:
        return {}, None


def _save_customers(customers, sha):
    """Upsert all customer records into Supabase. sha is unused; kept for call-site compatibility."""
    if not (_SUPABASE_URL and _SUPABASE_SERVICE_KEY):
        return False
    try:
        rows = []
        for username, rec in customers.items():
            rows.append({
                "username": username,
                "salt": rec.get("salt"),
                "hash": rec.get("hash"),
                "email": rec.get("email"),
                "reset_hash": rec.get("reset_hash"),
                "reset_salt": rec.get("reset_salt"),
                "reset_expires": rec.get("reset_expires"),
            })
        r = requests.post(
            f"{_SUPABASE_URL}/rest/v1/accounts?on_conflict=username",
            headers={
                "apikey": _SUPABASE_SERVICE_KEY,
                "Authorization": f"Bearer {_SUPABASE_SERVICE_KEY}",
                "Content-Type": "application/json",
                "Prefer": "resolution=merge-duplicates",
            },
            json=rows,
            timeout=10,
        )
        return r.status_code in (200, 201, 204)
    except Exception:
        return False


def _hash_pw(password, salt_hex):
    salt = bytes.fromhex(salt_hex)
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 100000).hex()

def _check_entitlement(email):
    """True only if this email has an ACTIVE row in Supabase entitlements
    with a plan_name in _ALLOWED_PLANS. Fails closed on any error."""
    if not (_SUPABASE_URL and _SUPABASE_SERVICE_KEY and email):
        return False
    try:
        plans_filter = ",".join(_ALLOWED_PLANS)
        url = (
            f"{_SUPABASE_URL}/rest/v1/entitlements"
            f"?email=eq.{email.strip().lower()}"
            f"&status=eq.active"
            f"&plan_name=in.({plans_filter})"
            f"&select=id"
        )
        r = requests.get(
            url,
            headers={"apikey": _SUPABASE_SERVICE_KEY, "Authorization": f"Bearer {_SUPABASE_SERVICE_KEY}"},
            timeout=10,
        )
        return r.status_code == 200 and len(r.json()) > 0
    except Exception:
        return False

def _check_team_access(email):
    """True if this email is an active team member under someone with an active Founders Plan."""
    if not (_SUPABASE_URL and _SUPABASE_SERVICE_KEY and email):
        return False
    try:
        url = (
            f"{_SUPABASE_URL}/rest/v1/team_members"
            f"?member_email=eq.{email.strip().lower()}"
            f"&status=eq.active"
            f"&select=owner_email"
        )
        r = requests.get(
            url,
            headers={"apikey": _SUPABASE_SERVICE_KEY, "Authorization": f"Bearer {_SUPABASE_SERVICE_KEY}"},
            timeout=10,
        )
        if r.status_code != 200:
            return False
        rows = r.json()
        if not rows:
            return False
        owner_email = rows[0].get("owner_email", "")
        if not owner_email:
            return False
        url2 = (
            f"{_SUPABASE_URL}/rest/v1/entitlements"
            f"?email=eq.{owner_email.strip().lower()}"
            f"&status=eq.active"
            f"&plan_name=in.(Founders Plan,Standard Plan)"
            f"&select=id"
        )
        r2 = requests.get(
            url2,
            headers={"apikey": _SUPABASE_SERVICE_KEY, "Authorization": f"Bearer {_SUPABASE_SERVICE_KEY}"},
            timeout=10,
        )
        return r2.status_code == 200 and len(r2.json()) > 0
    except Exception:
        return False


def _has_access(email):
    """True if this email has its own active entitlement OR is an active team member under a Founders Plan owner."""
    return _check_entitlement(email) or _check_team_access(email)

def _send_email(to_addr, subject, body):
    if not (_SMTP_HOST and _SMTP_USER and _SMTP_PASSWORD):
        return False
    try:
        import smtplib
        from email.mime.text import MIMEText
        msg = MIMEText(body)
        msg["Subject"] = subject
        msg["From"] = _SMTP_FROM
        msg["To"] = to_addr
        with smtplib.SMTP(_SMTP_HOST, _SMTP_PORT, timeout=10) as server:
            server.starttls()
            server.login(_SMTP_USER, _SMTP_PASSWORD)
            server.sendmail(_SMTP_FROM, [to_addr], msg.as_string())
        return True
    except Exception:
        return False

def _get_query_token():
    try:
        return st.query_params.get("token")
    except Exception:
        pass
    try:
        params = st.experimental_get_query_params()
        vals = params.get("token")
        return vals[0] if vals else None
    except Exception:
        return None
def _try_sso_login():
    if st.session_state.get("_sso_checked"):
        return
    st.session_state["_sso_checked"] = True
    token = _get_query_token()
    if not token:
        return
    try:
        resp = requests.post(
            "https://tscwyaphfadvvwdsrdbv.supabase.co/functions/v1/verify-launch-token",
            json={"token": token},
            timeout=10,
        )
        data = resp.json()
    except Exception:
        return
    if data.get("valid") and "defensiveiq" in (data.get("products") or []):
        st.session_state["_pw_ok"] = True
        st.session_state["_user_email"] = (data.get("email") or "").strip().lower()
        st.rerun()
def _check_password():
    _try_sso_login()
    if st.session_state.get("_pw_ok"):
        return True

    st.markdown("### \U0001F512 DefensiveIQ Access")
    mode = st.radio("Access", ["Log In", "Create Account", "Forgot Password"], horizontal=True, key="_auth_mode", label_visibility="collapsed")

    if mode == "Log In":
        with st.form("_login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Log In")
            if submitted:
                customers, _ = _load_customers()
                key = username.strip().lower()
                record = customers.get(key)
                if record and _hash_pw(password, record["salt"]) == record["hash"]:
                    if _has_access(record.get("email", "")):
                        st.session_state["_pw_ok"] = True
                        st.rerun()
                    else:
                        st.error("Your subscription doesn't currently include access to DefensiveIQ.")
                else:
                    st.error("Incorrect username or password.")

    elif mode == "Create Account":
        with st.form("_signup_form"):
            new_username = st.text_input("Choose a username")
            new_email = st.text_input("Email (must match your active subscription)")
            new_password = st.text_input("Choose a password", type="password")
            confirm_password = st.text_input("Confirm password", type="password")
            submitted = st.form_submit_button("Create Account")
            if submitted:
                key = new_username.strip().lower()
                email = new_email.strip().lower()
                if not key or not new_password or not email:
                    st.error("Username, email, and password are required.")
                elif len(new_password) < 6:
                    st.error("Password must be at least 6 characters.")
                elif new_password != confirm_password:
                    st.error("Passwords do not match.")
                elif not _has_access(email):
                    st.error("We couldn't find an active subscription for that email, and it's not listed as a team member on someone else's plan.")
                else:
                    customers, sha = _load_customers()
                    if key in customers:
                        st.error("That username is already taken.")
                    elif any((rec.get("email") or "").strip().lower() == email for rec in customers.values()):
                        st.error("An account already exists for that email. Please log in instead, or use Forgot Password.")
                    else:
                        salt_hex = os.urandom(16).hex()
                        customers[key] = {"salt": salt_hex, "hash": _hash_pw(new_password, salt_hex), "email": email}
                        if _save_customers(customers, sha):
                            st.session_state["_pw_ok"] = True
                            st.success("Account created! Loading DefensiveIQ...")
                            st.rerun()
                        else:
                            st.error("Could not create account right now. Please try again.")

    else:  # Forgot Password
        if "_reset_stage" not in st.session_state:
            st.session_state["_reset_stage"] = "request"

        if st.session_state["_reset_stage"] == "request":
            with st.form("_reset_request_form"):
                username = st.text_input("Username")
                submitted = st.form_submit_button("Send reset code")
                if submitted:
                    customers, sha = _load_customers()
                    key = username.strip().lower()
                    record = customers.get(key)
                    if not record or not record.get("email"):
                        st.error("No account with an email on file was found for that username.")
                    else:
                        code = f"{int.from_bytes(os.urandom(3), 'big') % 1000000:06d}"
                        salt_hex = os.urandom(16).hex()
                        record["reset_hash"] = _hash_pw(code, salt_hex)
                        record["reset_salt"] = salt_hex
                        record["reset_expires"] = (datetime.utcnow() + timedelta(minutes=15)).isoformat()
                        customers[key] = record
                        if _save_customers(customers, sha) and _send_email(
                            record["email"], "DefensiveIQ password reset code",
                            f"Your DefensiveIQ password reset code is {code}. It expires in 15 minutes."
                        ):
                            st.session_state["_reset_user"] = key
                            st.session_state["_reset_stage"] = "confirm"
                            st.success("A reset code was emailed to you.")
                            st.rerun()
                        else:
                            st.error("Could not send reset code right now. Please try again.")
        else:
            with st.form("_reset_confirm_form"):
                code = st.text_input("Enter the 6-digit code emailed to you")
                new_password = st.text_input("New password", type="password")
                confirm_password = st.text_input("Confirm new password", type="password")
                submitted = st.form_submit_button("Reset Password")
                if submitted:
                    customers, sha = _load_customers()
                    key = st.session_state.get("_reset_user", "")
                    record = customers.get(key)
                    if not record or not record.get("reset_hash"):
                        st.error("Reset session expired. Please start again.")
                        st.session_state["_reset_stage"] = "request"
                    elif datetime.fromisoformat(record["reset_expires"]) < datetime.utcnow():
                        st.error("This code has expired. Please request a new one.")
                        st.session_state["_reset_stage"] = "request"
                    elif _hash_pw(code, record["reset_salt"]) != record["reset_hash"]:
                        st.error("Incorrect code.")
                    elif len(new_password) < 6:
                        st.error("Password must be at least 6 characters.")
                    elif new_password != confirm_password:
                        st.error("Passwords do not match.")
                    else:
                        new_salt = os.urandom(16).hex()
                        record["salt"] = new_salt
                        record["hash"] = _hash_pw(new_password, new_salt)
                        record.pop("reset_hash", None); record.pop("reset_salt", None); record.pop("reset_expires", None)
                        customers[key] = record
                        if _save_customers(customers, sha):
                            st.session_state["_reset_stage"] = "request"
                            st.success("Password reset! Please log in.")
                            st.rerun()
                        else:
                            st.error("Could not reset password right now. Please try again.")
    return False

st.session_state["_pw_ok"] = True

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@700;800;900&family=Barlow:wght@400;500;600&family=Share+Tech+Mono&display=swap');
html,body,[class*="css"]{font-family:'Barlow',sans-serif;background-color:#0a1628;color:#f0ede8;}
.stApp{background-color:#0a1628;}
.main-title{font-family:'Barlow Condensed',sans-serif;font-weight:900;font-size:64px;line-height:.95;text-transform:uppercase;color:#f0ede8;margin-bottom:8px;}
.stButton>button{background:#7b241c!important;color:#f0ede8!important;border:none!important;font-family:'Barlow Condensed',sans-serif!important;font-weight:700!important;font-size:16px!important;letter-spacing:.1em!important;text-transform:uppercase!important;padding:12px 32px!important;border-radius:0!important;width:100%!important;}
.stButton>button:hover{background:#5e1c15!important;}
.stDownloadButton>button{background:#0e7060!important;color:#f0ede8!important;border:none!important;font-family:'Barlow Condensed',sans-serif!important;font-weight:700!important;font-size:14px!important;letter-spacing:.08em!important;text-transform:uppercase!important;border-radius:0!important;width:100%!important;}
</style>
""", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────
def get_zone(y):
    try:
        y=float(y)
        if y<=-1  and y>=-20: return "BZ"
        if y<=-21 and y>=-49: return "OF"
        if y>=40  and y<=50:  return "MF"
        if y>=21  and y<=39:  return "FZ"
        if y>=11  and y<=20:  return "RZ"
        if y>=1   and y<=10:  return "GL"
    except: pass
    return None

def top3(plays, key):
    vals=[str(p.get(key,'')) for p in plays
          if p.get(key) not in (None,'') and str(p.get(key,'')).strip() not in ('','nan','None')]
    if not vals: return []
    return [{"v":v,"n":c} for v,c in Counter(vals).most_common(3)]

def pct(n,d): return round(n/d*100) if d>0 else 0

def _num(v, default=0.0):
    """Safely parse a number from a spreadsheet cell.
    Handles NaN, blanks, and non-numeric junk without poisoning averages."""
    try:
        f = float(v)
        if f != f:   # NaN check
            return default
        return f
    except (TypeError, ValueError):
        return default

def is_success(dn, dist, gain, zone, yard_ln, is_td):
    """Success = TD (auto), or gained enough yards for the down.
    Goal line/inside 10: measure yards needed to reach end zone.
    Else: 1st >=40%, 2nd >=50%, 3rd/4th = convert (100%).
    This measures the OFFENSE's success on the play — i.e. the situations
    where the opponent has moved the chains against whatever look they saw."""
    if is_td:
        return True
    try:
        dn=int(dn); dist=float(dist); gain=float(gain)
    except:
        return None
    if zone == "GL":
        try:
            to_goal = float(yard_ln)
            if to_goal > 0:
                dist = to_goal
        except: pass
    if dist <= 0: dist = 10
    if dn==1: return gain >= 0.40*dist
    if dn==2: return gain >= 0.50*dist
    if dn>=3: return gain >= dist
    return None

def _motion_read(raw):
    """MOTION is often tagged as a direction (L/R) or as a count of motions,
    rather than a descriptive name. Return (is_motion, label).
    'N'/'NO'/'0'/blank-equivalent = no motion. L/R = directional motion.
    Named values (e.g. 'JET', 'ORBIT') are treated as motion with that name."""
    s = str(raw).strip()
    if s in ('', 'nan', 'None'):
        return None, ''          # untagged — unknown, not counted either way
    su = s.upper()
    if su in ('N','NO','NONE','0','NO MOTION'):
        return False, 'No Motion'
    if su in ('L','LEFT'):
        return True, 'Motion Left'
    if su in ('R','RIGHT'):
        return True, 'Motion Right'
    try:
        n = float(s)
        if n <= 0: return False, 'No Motion'
        return True, f"Motion x{int(n)}"
    except ValueError:
        return True, su          # a named motion (e.g. JET, ORBIT)

def load_plays(df):
    plays=[]
    for _,row in df.iterrows():
        pt=str(row.get('PLAY TYPE','')).strip()
        if pt not in ('Run','Pass'): continue
        zone=get_zone(row.get('YARD LN',''))
        if not zone: continue
        dn_v   = int(_num(row.get('DN',0)))
        dist_v = _num(row.get('DIST',0))
        gain_v = _num(row.get('GN/LS',0))
        yl_v   = row.get('YARD LN','')
        result = str(row.get('RESULT','')).upper()
        is_td  = ('TD' in result) or ('TOUCHDOWN' in result)
        succ   = is_success(dn_v, dist_v, gain_v, zone, yl_v, is_td)
        expl   = (gain_v >= 10) if pt=='Run' else (gain_v >= 15)
        motioned, motion_lbl = _motion_read(row.get('MOTION',''))
        strength = str(row.get('OFF STR','')).strip().upper()
        plays.append({
            'zone':   zone,
            'dn':     dn_v,
            'dist':   dist_v,
            'hash':   str(row.get('HASH','')).strip(),
            'form':   str(row.get('OFF FORM','')).strip(),
            'play':   str(row.get('OFF PLAY','')).strip(),
            'dir':    str(row.get('PLAY DIR','')).strip(),
            'rp':     pt,
            'gnls':   gain_v,
            'personnel': str(row.get('PERSONNEL','')).strip(),
            'backfield': str(row.get('BACKFIELD','')).strip(),
            'strength':  strength,
            'motion': motion_lbl,
            'motioned':  motioned,
            'result': str(row.get('RESULT','')).strip(),
            'succ':   succ,
            'expl':   expl,
            'td':     is_td,
        })
    return plays


# ── PowerPoint scouting deck (offensive scouting for DCs) ──
P_NAVY=RGBColor(0x16,0x21,0x3E); P_RED=RGBColor(0xC0,0x39,0x2B)
P_BLUE=RGBColor(0x1A,0x52,0x76); P_TEAL=RGBColor(0x0E,0x70,0x60)
P_GOLD=RGBColor(0xC9,0xA2,0x27); P_WHITE=RGBColor(0xFF,0xFF,0xFF)
P_LGRAY=RGBColor(0xF2,0xF2,0xF2); P_DGRAY=RGBColor(0x55,0x55,0x55)
P_BLACK=RGBColor(0x11,0x11,0x11)
PW, PH = Inches(13.333), Inches(7.5)

def _p_slide(prs, bg=P_WHITE):
    s=prs.slides.add_slide(prs.slide_layouts[6])
    r=s.shapes.add_shape(1,0,0,PW,PH)
    r.fill.solid(); r.fill.fore_color.rgb=bg; r.line.fill.background(); r.shadow.inherit=False
    sp=r._element; sp.getparent().remove(sp); s.shapes._spTree.insert(2,sp)
    return s

def _p_text(slide,x,y,w,h,text,size=14,color=P_BLACK,bold=False,align=PP_ALIGN.LEFT,
            font="Calibri",anchor=MSO_ANCHOR.TOP,italic=False):
    tb=slide.shapes.add_textbox(x,y,w,h); tf=tb.text_frame
    tf.word_wrap=True; tf.vertical_anchor=anchor
    tf.margin_left=Pt(2); tf.margin_right=Pt(2); tf.margin_top=Pt(1); tf.margin_bottom=Pt(1)
    p=tf.paragraphs[0]; p.alignment=align
    run=p.add_run(); run.text=text
    run.font.size=Pt(size); run.font.bold=bold; run.font.italic=italic
    run.font.color.rgb=color; run.font.name=font
    return tb

def _p_rect(slide,x,y,w,h,fill,line=None):
    r=slide.shapes.add_shape(1,x,y,w,h)
    r.fill.solid(); r.fill.fore_color.rgb=fill
    if line: r.line.color.rgb=line; r.line.width=Pt(0.75)
    else: r.line.fill.background()
    r.shadow.inherit=False
    return r

def _p_table(slide,x,y,w,rows,col_widths,font_size=11,row_h=Inches(0.34)):
    gt=slide.shapes.add_table(len(rows),len(rows[0]),x,y,w,row_h*len(rows)).table
    gt.first_row=False; gt.horz_banding=False
    for ci,cw in enumerate(col_widths): gt.columns[ci].width=cw
    for ri,rowvals in enumerate(rows):
        gt.rows[ri].height=row_h
        for ci,(val,opts) in enumerate(rowvals):
            cell=gt.cell(ri,ci)
            cell.margin_left=Pt(4); cell.margin_right=Pt(4)
            cell.margin_top=Pt(1); cell.margin_bottom=Pt(1)
            cell.vertical_anchor=MSO_ANCHOR.MIDDLE
            cell.fill.solid(); cell.fill.fore_color.rgb=opts.get('bg', P_WHITE if ri%2 else P_LGRAY)
            tf=cell.text_frame; tf.word_wrap=True
            p=tf.paragraphs[0]; p.alignment=opts.get('align',PP_ALIGN.LEFT)
            run=p.add_run(); run.text=str(val)
            run.font.size=Pt(opts.get('size',font_size))
            run.font.bold=opts.get('bold',False)
            run.font.color.rgb=opts.get('fc',P_BLACK)
            run.font.name='Calibri'
    return gt

def build_pptx(plays, opp, week, date, primary_hex="#7B241C", accent_hex="#C9A227"):
    def _hex2rgb(h):
        h=h.lstrip('#'); return RGBColor(int(h[0:2],16),int(h[2:4],16),int(h[4:6],16))
    def _lum(h):
        h=h.lstrip('#'); r,g,b=int(h[0:2],16),int(h[2:4],16),int(h[4:6],16)
        return (0.299*r+0.587*g+0.114*b)/255
    PRIMARY=_hex2rgb(primary_hex); ACCENT=_hex2rgb(accent_hex)
    ON_PRIMARY = P_WHITE if _lum(primary_hex)<0.55 else P_BLACK
    ACCENT_ON_PRIMARY = ACCENT if _lum(accent_hex)>0.35 else P_WHITE

    def _sr(lst):
        v=[p for p in lst if p.get('succ') is not None]
        return round(sum(1 for p in v if p['succ'])/len(v)*100) if v else None
    def _avg(lst):
        return (sum(p['gnls'] for p in lst)/len(lst)) if lst else 0

    prs=Presentation(); prs.slide_width=PW; prs.slide_height=PH
    total=len(plays); opp=opp or "Opponent"
    runs=[p for p in plays if p['rp']=='Run']; passes=[p for p in plays if p['rp']=='Pass']

    # SLIDE 1 — Title
    s=_p_slide(prs,PRIMARY)
    _p_text(s,Inches(0.8),Inches(2.4),Inches(11.7),Inches(1.2),
            "OFFENSIVE SCOUTING REPORT",42,ON_PRIMARY,bold=True,align=PP_ALIGN.CENTER,font="Cambria")
    _p_text(s,Inches(0.8),Inches(3.6),Inches(11.7),Inches(0.9),
            opp.upper(),32,ACCENT_ON_PRIMARY,bold=True,align=PP_ALIGN.CENTER,font="Cambria")
    sub="  ·  ".join([x for x in [f"Week {week}" if week else "", date or "", f"{total} plays analyzed"] if x])
    _p_text(s,Inches(0.8),Inches(4.6),Inches(11.7),Inches(0.5),sub,16,
            RGBColor(0xF0,0xC9,0xC2),align=PP_ALIGN.CENTER)
    for lx,lbl in [(Inches(1.0),"[ YOUR LOGO ]"),(Inches(9.9),"[ OPP LOGO ]")]:
        _p_rect(s,lx,Inches(0.6),Inches(2.4),Inches(1.3),PRIMARY,line=RGBColor(0x9A,0x5A,0x4A))
        _p_text(s,lx,Inches(1.05),Inches(2.4),Inches(0.5),lbl,11,RGBColor(0xC8,0x9A,0x8E),
                align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    _p_text(s,Inches(0.8),Inches(6.9),Inches(11.7),Inches(0.4),
            "DefensiveIQ  ·  Auto-generated from film — replace logos and decorate freely",10,
            RGBColor(0xC8,0x9A,0x8E),align=PP_ALIGN.CENTER,italic=True)

    # SLIDE 2 — Overview
    s=_p_slide(prs,P_WHITE)
    _p_text(s,Inches(0.6),Inches(0.4),Inches(12),Inches(0.8),
            "OFFENSIVE OVERVIEW",36,PRIMARY,bold=True,font="Cambria")
    motioned=[p for p in plays if p.get('motioned') is True]
    nonm=[p for p in plays if p.get('motioned') is False]
    known=len(motioned)+len(nonm)
    motion_rate=round(len(motioned)/known*100) if known else 0
    sr_all=_sr(plays)
    stats=[("PLAYS SCOUTED",str(total),PRIMARY),
           ("YDS PER PLAY",f"{_avg(plays):.1f}",P_TEAL),
           ("SUCCESS RATE",f"{sr_all}%" if sr_all is not None else "—",P_TEAL),
           ("MOTION RATE",f"{motion_rate}%",P_RED),
           ("EXPLOSIVE %",f"{round(len([p for p in plays if p['expl']])/total*100) if total else 0}%",P_BLUE)]
    cw=Inches(2.3); gap=Inches(0.15); x0=Inches(0.6); y0=Inches(1.5)
    for i,(lbl,val,col) in enumerate(stats):
        x=x0+(cw+gap)*i
        _p_rect(s,x,y0,cw,Inches(1.8),P_LGRAY)
        _p_text(s,x,y0+Inches(0.25),cw,Inches(0.9),val,44,col,bold=True,
                align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE,font="Cambria")
        _p_text(s,x,y0+Inches(1.25),cw,Inches(0.4),lbl,12,P_DGRAY,bold=True,align=PP_ALIGN.CENTER)
    # key reads
    _p_text(s,Inches(0.6),Inches(3.7),Inches(12),Inches(0.5),"KEY READS",20,P_RED,bold=True)
    reads=[]
    fc=Counter(p['form'] for p in plays if str(p['form']).strip() not in ('','nan','None','0'))
    if fc:
        f,n=fc.most_common(1)[0]
        reads.append(f"Base formation is {f} — {round(n/total*100)}% of snaps")
    bc=Counter(p['backfield'] for p in plays if str(p['backfield']).strip() not in ('','nan','None'))
    if bc:
        b,n=bc.most_common(1)[0]
        reads.append(f"Primary backfield set is {b} — {round(n/total*100)}% of snaps")
    if known:
        reads.append(f"They use motion on {motion_rate}% of snaps")
        if motioned and nonm:
            d=_avg(motioned)-_avg(nonm)
            if abs(d)>=0.8:
                reads.append(("They gain {:.1f} more yards per play WITH motion").format(d) if d>0
                             else ("They gain {:.1f} more yards per play WITHOUT motion").format(-d))
    # most dangerous / most contained look by avg gain (4+ snaps)
    best=[]
    for key,lbl in [('form','formation'),('backfield','backfield')]:
        groups={}
        for p in plays:
            v=str(p.get(key,'')).strip()
            if v in ('','nan','None','0'): continue
            groups.setdefault(v,[]).append(p)
        for v,g in groups.items():
            if len(g)>=4: best.append((_avg(g),v,lbl,len(g)))
    if best:
        best.sort(reverse=True)
        a=best[0]; reads.append(f"Most dangerous: {a[1]} ({a[2]}) — {a[0]:.1f} yds/play on {a[3]} snaps")
        w=best[-1]; reads.append(f"Least productive: {w[1]} ({w[2]}) — {w[0]:.1f} yds/play on {w[3]} snaps")
    yy=Inches(4.35)
    for txt in reads[:5]:
        _p_rect(s,Inches(0.7),yy+Inches(0.05),Inches(0.22),Inches(0.22),P_RED)
        _p_text(s,Inches(1.05),yy,Inches(11),Inches(0.4),txt,16,P_BLACK,bold=True)
        yy+=Inches(0.5)

    # SLIDE 3/4 — Formations & Personnel
    def matchup_slide(title, key, color):
        s=_p_slide(prs,P_WHITE)
        _p_text(s,Inches(0.6),Inches(0.4),Inches(12),Inches(0.8),title,36,color,bold=True,font="Cambria")
        _p_text(s,Inches(0.6),Inches(1.15),Inches(12),Inches(0.4),
                "How the opposing offense has performed out of each look.",12,P_DGRAY,italic=True)
        groups={}
        for p in plays:
            v=str(p.get(key,'')).strip()
            if v in ('','nan','None','0'): continue
            groups.setdefault(v,[]).append(p)
        ranked=sorted(groups.items(), key=lambda kv:-len(kv[1]))
        ranked=[(v,g) for v,g in ranked if len(g)>=3][:9]
        rows=[[("LOOK",{'bg':color,'fc':P_WHITE,'bold':True}),
               ("SNAPS",{'bg':color,'fc':P_WHITE,'bold':True,'align':PP_ALIGN.CENTER}),
               ("% SEEN",{'bg':color,'fc':P_WHITE,'bold':True,'align':PP_ALIGN.CENTER}),
               ("YDS/PLAY",{'bg':color,'fc':P_WHITE,'bold':True,'align':PP_ALIGN.CENTER}),
               ("SUCCESS%",{'bg':color,'fc':P_WHITE,'bold':True,'align':PP_ALIGN.CENTER}),
               ("EXPL%",{'bg':color,'fc':P_WHITE,'bold':True,'align':PP_ALIGN.CENTER})]]
        for i,(v,g) in enumerate(ranked):
            bg=P_LGRAY if i%2==0 else P_WHITE
            sr=_sr(g); mark=" *" if len(g)<5 else ""
            rows.append([(str(v)+mark,{'bg':bg,'bold':True,'size':12}),
                         (str(len(g)),{'bg':bg,'align':PP_ALIGN.CENTER}),
                         (f"{round(len(g)/total*100)}%",{'bg':bg,'align':PP_ALIGN.CENTER}),
                         (f"{_avg(g):.1f}",{'bg':bg,'fc':P_TEAL,'bold':True,'align':PP_ALIGN.CENTER}),
                         (f"{sr}%" if sr is not None else "—",{'bg':bg,'fc':P_TEAL,'bold':True,'align':PP_ALIGN.CENTER}),
                         (f"{round(len([p for p in g if p['expl']])/len(g)*100)}%",{'bg':bg,'fc':P_BLUE,'bold':True,'align':PP_ALIGN.CENTER})])
        if len(rows)==1:
            _p_text(s,Inches(0.6),Inches(2),Inches(12),Inches(0.5),
                    "Not enough tagged data for this section.",14,P_DGRAY,italic=True)
        else:
            _p_table(s,Inches(0.6),Inches(1.7),Inches(12.1),rows,
                     [Inches(3.6),Inches(1.7),Inches(1.7),Inches(1.7),Inches(1.7),Inches(1.7)],
                     row_h=Inches(0.44))
            _p_text(s,Inches(0.6),Inches(6.9),Inches(12),Inches(0.35),
                    "* = small sample (under 5 snaps)",10,P_DGRAY,italic=True)

    matchup_slide("FORMATIONS THEY RUN",'form',P_RED)
    matchup_slide("PERSONNEL & BACKFIELD SETS",'backfield',P_BLUE)

    # SLIDE 5 — Motion Report
    s=_p_slide(prs,P_WHITE)
    _p_text(s,Inches(0.6),Inches(0.4),Inches(12),Inches(0.8),
            "MOTION REPORT",36,RGBColor(0x4A,0x23,0x5A),bold=True,font="Cambria")
    rows=[[("SITUATION",{'bg':RGBColor(0x4A,0x23,0x5A),'fc':P_WHITE,'bold':True}),
           ("SNAPS",{'bg':RGBColor(0x4A,0x23,0x5A),'fc':P_WHITE,'bold':True,'align':PP_ALIGN.CENTER}),
           ("MOTION%",{'bg':RGBColor(0x4A,0x23,0x5A),'fc':P_WHITE,'bold':True,'align':PP_ALIGN.CENTER}),
           ("YDS/PLAY",{'bg':RGBColor(0x4A,0x23,0x5A),'fc':P_WHITE,'bold':True,'align':PP_ALIGN.CENTER}),
           ("SUCCESS%",{'bg':RGBColor(0x4A,0x23,0x5A),'fc':P_WHITE,'bold':True,'align':PP_ALIGN.CENTER})]]
    sits=[("1st & 10",lambda p:p['dn']==1 and p['dist']>=8),
          ("2nd & Long",lambda p:p['dn']==2 and p['dist']>=7),
          ("2nd & Short",lambda p:p['dn']==2 and p['dist']<=3),
          ("3rd & Long",lambda p:p['dn']==3 and p['dist']>=7),
          ("3rd & Med",lambda p:p['dn']==3 and 4<=p['dist']<=6),
          ("3rd & Short",lambda p:p['dn']==3 and p['dist']<=3),
          ("Red Zone",lambda p:p['zone']=='RZ'),
          ("Goal Line",lambda p:p['zone']=='GL')]
    for i,(lbl,fn) in enumerate(sits):
        sp=[p for p in plays if fn(p)]
        spk=[p for p in sp if p.get('motioned') is not None]
        sb=[p for p in spk if p['motioned'] is True]
        bg=P_LGRAY if i%2==0 else P_WHITE
        sr=_sr(sp)
        rows.append([(lbl,{'bg':bg,'bold':True,'size':12}),
                     (str(len(sp)),{'bg':bg,'align':PP_ALIGN.CENTER}),
                     (f"{round(len(sb)/len(spk)*100)}%" if spk else "—",
                      {'bg':bg,'fc':P_RED,'bold':True,'align':PP_ALIGN.CENTER}),
                     (f"{_avg(sp):.1f}" if sp else "—",{'bg':bg,'fc':P_TEAL,'align':PP_ALIGN.CENTER}),
                     (f"{sr}%" if sr is not None else "—",{'bg':bg,'fc':P_TEAL,'bold':True,'align':PP_ALIGN.CENTER})])
    _p_table(s,Inches(0.6),Inches(1.6),Inches(12.1),rows,
             [Inches(3.3),Inches(2.2),Inches(2.2),Inches(2.2),Inches(2.2)],row_h=Inches(0.46))

    # SLIDE 6 — Where to focus (dark closer)
    s=_p_slide(prs,PRIMARY)
    _p_text(s,Inches(0.6),Inches(0.4),Inches(12),Inches(0.8),
            "FORMATIONS TO GAME-PLAN AGAINST",36,ACCENT_ON_PRIMARY,bold=True,font="Cambria")
    fgroups={}
    for p in plays:
        f=str(p.get('form','')).strip()
        if f in ('','nan','None'): continue
        fgroups.setdefault(f,[]).append(p)
    ranked=[(f,g) for f,g in fgroups.items() if len(g)>=3]
    ranked.sort(key=lambda kv:-_avg(kv[1]))
    yy=Inches(1.6)
    for f,g in ranked[:7]:
        sr=_sr(g)
        _p_rect(s,Inches(0.8),yy,Inches(4.0),Inches(0.6),P_RED)
        _p_text(s,Inches(0.95),yy,Inches(3.7),Inches(0.6),f,14,P_WHITE,bold=True,anchor=MSO_ANCHOR.MIDDLE)
        _p_rect(s,Inches(4.9),yy,Inches(7.6),Inches(0.6),RGBColor(0x33,0x1E,0x1B))
        txt=(f"{len(g)} snaps   ·   {_avg(g):.1f} yds/play   ·   {sr}% success rate"
             if sr is not None else f"{len(g)} snaps   ·   {_avg(g):.1f} yds/play")
        _p_text(s,Inches(5.1),yy,Inches(7.2),Inches(0.6),txt,14,P_WHITE,anchor=MSO_ANCHOR.MIDDLE)
        yy+=Inches(0.72)
    if not ranked:
        _p_text(s,Inches(0.8),Inches(2),Inches(11),Inches(0.5),
                "Not enough formation data tagged.",14,RGBColor(0xC8,0x9A,0x8E),italic=True)
    _p_text(s,Inches(0.6),Inches(6.95),Inches(12),Inches(0.4),
            "Formations ranked by yards per play (3+ snaps) — these are where the opposing "
            "offense has been most productive. Small samples can mislead — verify against film.",10,
            RGBColor(0xC8,0x9A,0x8E),italic=True)

    buf=io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf.getvalue()

# ── Excel Builder ─────────────────────────────────────────────
def build_excel(plays, opp, week, date):
    FN="Arial"; CW="FFFFFFFF"; CL="FFF5F5F5"; CB="FF16213E"
    CBl="FF1A5276"; CTe="FF0E7060"; CPu="FF4A235A"; CR="FF7B241C"
    CRB="FFFDE8E8"; CPB="FFE8F0FE"
    CDG="FF555555"; CGr="FF1E8449"; CYB="FFFFFBE6"
    # Zone colors
    ZONE_BG={"BZ":"FFFDE8E8","OF":"FFE8F0FE","MF":"FFE8F8E8",
              "FZ":"FFFFFBE6","RZ":"FFFCE4EC","GL":"FFEDE7F6"}
    ZONE_HDR={"BZ":CR,"OF":CBl,"MF":CTe,"FZ":"FF7D6608","RZ":CR,"GL":CPu}

    def fil(c): return PatternFill("solid",fgColor=c)
    def bdr():
        s=Side(style="thin",color="FFB0B0B0")
        return Border(left=s,right=s,top=s,bottom=s)
    def sc(ws,r,c,val="",bold=False,sz=10,fc=CB,bg=None,h="center",v="center",wrap=False,fmt=None):
        cell=ws.cell(row=r,column=c,value=val)
        cell.font=Font(name=FN,bold=bold,size=sz,color=fc)
        if bg: cell.fill=fil(bg)
        cell.alignment=Alignment(horizontal=h,vertical=v,wrap_text=wrap)
        cell.border=bdr()
        if fmt: cell.number_format=fmt
        return cell
    def hdr(ws,r,c,txt,bg=CBl,fc=CW,sz=9,wrap=True,span=1):
        cell=sc(ws,r,c,txt,bold=True,sz=sz,fc=fc,bg=bg,wrap=wrap)
        if span>1: ws.merge_cells(start_row=r,start_column=c,end_row=r,end_column=c+span-1)
        return cell
    def banner(ws,r,txt,nc,bg=CB,fc=CW,sz=13,ht=30):
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=nc)
        c=ws.cell(row=r,column=1,value=txt)
        c.font=Font(name=FN,bold=True,size=sz,color=fc)
        c.fill=fil(bg); c.alignment=Alignment(horizontal="center",vertical="center")
        ws.row_dimensions[r].height=ht
    def widths(ws,lst):
        for i,w in enumerate(lst,1): ws.column_dimensions[gcl(i)].width=w

    zone_list=["BZ","OF","MF","FZ","RZ","GL"]
    zone_names={"BZ":"Backed Up  Own 1–20","OF":"Open Field  Own 21–49",
                "MF":"Midfield  50–Opp 40","FZ":"Fringe  Opp 39–21",
                "RZ":"Red Zone  Opp 20–11","GL":"Goal Line  Opp 10 and in"}
    zone_bgs={"BZ":"FFFDE8E8","OF":"FFE8F0FE","MF":"FFE8F8E8",
              "FZ":"FFFFFBE6","RZ":"FFFCE4EC","GL":"FFEDE7F6"}
    zone_hdrs={"BZ":"FF7B241C","OF":"FF1A5276","MF":"FF0E7060",
               "FZ":"FF7D6608","RZ":"FF7B241C","GL":"FF4A235A"}
    runs=[p for p in plays if p['rp']=='Run']
    passes=[p for p in plays if p['rp']=='Pass']
    total=len(plays)
    has_form=any(p['form'] for p in plays)
    has_back=any(p['backfield'] for p in plays)
    has_motion=any(p['motion'] for p in plays)

    wb2=Workbook()

    # ── Tab 1: Film Log ──────────────────────────────────────
    ws_log=wb2.active; ws_log.title="1. Film Log"
    ws_log.sheet_properties.tabColor="7B241C"
    ws_log.sheet_view.showGridLines=False
    log_cols=[('QTR',6),('DN',6),('DIST',6),('HASH',6),('YARD LN',9),('ZONE',10),
              ('OFF FORM',18),('OFF PLAY',20),('PLAY DIR',9),('PLAY TYPE',10),
              ('GN/LS',8),('RESULT',10),('BACKFIELD',14),('OFF STR',9),('MOTION',12)]
    widths(ws_log,[w for _,w in log_cols])
    banner(ws_log,1,"DEFENSIVE IQ — FILM LOG  |  Tag OFF FORM, BACKFIELD, OFF STR, MOTION in Hudl for full analysis",
           len(log_cols),bg=CR,sz=10,ht=26)
    ws_log.row_dimensions[2].height=36
    for ci,(col,_) in enumerate(log_cols,1):
        if col=="ZONE":    bg="FF000088"
        elif col in("BACKFIELD","OFF STR","MOTION"): bg=CPu
        else: bg=CR
        hdr(ws_log,2,ci,col,bg=bg,sz=8)
    for ri,p in enumerate(plays):
        r=ri+3; ws_log.row_dimensions[r].height=15
        bg=CL if ri%2==0 else CW
        vals={'QTR':'','DN':p['dn'],'DIST':p['dist'],'HASH':p['hash'],
              'YARD LN':'','ZONE':p['zone'],'OFF FORM':p['form'],'OFF PLAY':p['play'],
              'PLAY DIR':p['dir'],'PLAY TYPE':p['rp'],'GN/LS':p['gnls'],
              'RESULT':p['result'],'BACKFIELD':p['backfield'],'OFF STR':p['strength'],'MOTION':p['motion']}
        for ci,(col,_) in enumerate(log_cols,1):
            zbg="FFE8F4FD" if col=="ZONE" else ("FFEDE7F6" if col in("BACKFIELD","OFF STR","MOTION") else bg)
            sc(ws_log,r,ci,vals.get(col,''),sz=9,bg=zbg,fc="FF000000",
               h="left" if col in("OFF FORM","OFF PLAY","RESULT","BACKFIELD") else "center")
    ws_log.freeze_panes="A3"

    # ── Tab 2: Field Zone Tendencies ─────────────────────────
    ws2=wb2.create_sheet("2. Zone x Situation")
    ws2.sheet_properties.tabColor="7B241C"; ws2.sheet_view.showGridLines=False
    ws2.page_setup.orientation="landscape"
    ws2.page_setup.fitToPage=True; ws2.page_setup.fitToWidth=1; ws2.page_setup.fitToHeight=0
    NC2=8
    widths(ws2,[20,8,22,22,10,10,24,12])
    banner(ws2,1,"ZONE x SITUATION  —  What they show by field position & down/distance",NC2,bg=CB,sz=12,ht=30)
    ws2.merge_cells(f"A2:{gcl(NC2)}2")
    leg=ws2.cell(row=2,column=1,
        value="  Situations with no snaps are skipped. Cells marked * have 1-2 snaps — too thin to call a tendency.")
    leg.font=Font(name=FN,size=8,italic=True,color=CDG); leg.fill=fil("FFF0F0F0")
    leg.alignment=Alignment(horizontal="left",vertical="center")
    ws2.row_dimensions[2].height=14

    zs_sits=[("1st & 10",    lambda p: p['dn']==1 and p['dist']>=8),
             ("1st & Short", lambda p: p['dn']==1 and p['dist']<8),
             ("2nd & Long",  lambda p: p['dn']==2 and p['dist']>=7),
             ("2nd & Med",   lambda p: p['dn']==2 and 4<=p['dist']<=6),
             ("2nd & Short", lambda p: p['dn']==2 and p['dist']<=3),
             ("3rd & Long",  lambda p: p['dn']==3 and p['dist']>=7),
             ("3rd & Med",   lambda p: p['dn']==3 and 4<=p['dist']<=6),
             ("3rd & Short", lambda p: p['dn']==3 and p['dist']<=3),
             ("4th Down",    lambda p: p['dn']==4)]

    def _zs_top(lst,key,n=2):
        vals=[str(p.get(key,'')) for p in lst
              if str(p.get(key,'')).strip() not in ('','nan','None','0')]
        if not vals: return "—"
        return ", ".join(f"{v} ({c})" for v,c in Counter(vals).most_common(n))
    def _zs_sr(lst):
        v=[p for p in lst if p.get('succ') is not None]
        return round(sum(1 for p in v if p['succ'])/len(v)*100) if v else None

    row=3
    for zcode in zone_list:
        zp=[p for p in plays if p['zone']==zcode]
        if not zp: continue
        zk=[p for p in zp if p.get('motioned') is not None]
        zb=[p for p in zk if p['motioned'] is True]
        zmr=f"{round(len(zb)/len(zk)*100)}% motion" if zk else "motion n/a"
        ws2.merge_cells(start_row=row,start_column=1,end_row=row,end_column=NC2)
        c=ws2.cell(row=row,column=1,
            value=f"  {zcode}  ·  {zone_names[zcode]}   —   {len(zp)} snaps   ·   {zmr}")
        c.font=Font(name=FN,bold=True,size=11,color=CW)
        c.fill=fil(zone_hdrs[zcode]); c.alignment=Alignment(horizontal="left",vertical="center")
        ws2.row_dimensions[row].height=20
        row+=1
        for cn,txt in [(1,"SITUATION"),(2,"Snaps"),(3,"Top Formations"),(4,"Top Backfields"),
                       (5,"Motion%"),(6,"Yds/Play"),(7,"Named Motions"),(8,"Success%")]:
            hdr(ws2,row,cn,txt,bg=CB,sz=8)
        ws2.row_dimensions[row].height=16
        row+=1
        shown=0
        for si,(lbl,fn) in enumerate(zs_sits):
            try: sp=[p for p in zp if fn(p)]
            except: sp=[]
            if not sp: continue
            shown+=1
            thin=" *" if len(sp)<3 else ""
            bg=zone_bgs[zcode] if si%2==0 else CL
            spk=[p for p in sp if p.get('motioned') is not None]
            sb=[p for p in spk if p['motioned'] is True]
            ws2.row_dimensions[row].height=24
            sc(ws2,row,1,lbl+thin,bold=True,sz=9,fc=CW,bg=CTe,h="left")
            sc(ws2,row,2,len(sp),bold=True,sz=10,fc="FF000000",bg=bg,fmt="0")
            sc(ws2,row,3,_zs_top(sp,'form'),sz=8,fc="FF8B0000",bg=CRB,h="left",wrap=True)
            sc(ws2,row,4,_zs_top(sp,'backfield'),sz=8,fc="FF00008B",bg=CPB,h="left",wrap=True)
            if spk:
                sc(ws2,row,5,float(len(sb))/float(len(spk)),bold=True,sz=10,fc="FF7B241C",bg=CYB,fmt="0%")
            else:
                sc(ws2,row,5,"—",sz=9,fc=CDG,bg=bg)
            sc(ws2,row,6,round(sum(p['gnls'] for p in sp)/len(sp),1),sz=9,fc="FF0E7060",bg="FFE8F8E8",fmt="0.0")
            named=[p for p in sp if p.get('motion') and p['motion']!='No Motion'
                   and not p['motion'].startswith('Motion Left') and not p['motion'].startswith('Motion Right')
                   and not p['motion'].startswith('Motion x')]
            sc(ws2,row,7,_zs_top(named,'motion'),sz=8,fc="FF4A235A",bg="FFEDE7F6",h="left",wrap=True)
            sr=_zs_sr(sp)
            sc(ws2,row,8,(float(sr)/100.0 if sr is not None else "—"),sz=9,fc="FF0E7060",bg="FFE8F8E8",
               fmt="0%" if sr is not None else "General")
            row+=1
        if shown==0:
            ws2.merge_cells(start_row=row,start_column=1,end_row=row,end_column=NC2)
            c=ws2.cell(row=row,column=1,value="   No snaps recorded in this zone.")
            c.font=Font(name=FN,sz=9,italic=True,color=CDG); row+=1
        row+=1

    ws2.cell(row=row,column=1,
        value="* = 1-2 snaps only — too thin to read as a tendency.").font=Font(name=FN,sz=8,italic=True,color=CDG)
    ws2.freeze_panes="A3"

    # ── Tab 3: Formations ─────────────────────────────────────
    ws3=wb2.create_sheet("3. Formations")
    ws3.sheet_properties.tabColor="7B241C"; ws3.sheet_view.showGridLines=False
    NC3=10; widths(ws3,[8,12,9,10,20,20,20,20,20,28])
    banner(ws3,1,"FORMATIONS  —  What formations they line up in by zone",NC3,bg=CR,sz=12,ht=30)
    ws3.row_dimensions[2].height=20
    for c,txt,bg,span in[(1,"ZONE",CB,1),(2,"COUNTS",CB,3),(5,"TOP FORMATIONS USED",CR,3),(8,"TOP FORMATION vs RUN",CR,2),(10,"NOTES",CB,1)]:
        hdr(ws3,2,c,txt,bg=bg,sz=8,span=span)
    ws3.row_dimensions[3].height=36
    for c,txt,bg in[(1,"Zone",CB),(2,"Plays",CB),(3,"Run",CB),(4,"Pass",CB),
                     (5,"#1 Formation",CR),(6,"#2 Formation",CR),(7,"#3 Formation",CR),
                     (8,"#1 vs Run",CR),(9,"#2 vs Run",CR),(10,"Notes",CB)]:
        hdr(ws3,3,c,txt,bg=bg,sz=9,wrap=True)
    for ri,zcode in enumerate(zone_list):
        r=ri+4; zbg=ZONE_BG[zcode]; zhdr=ZONE_HDR[zcode]
        zp=[p for p in plays if p['zone']==zcode]
        zr=[p for p in zp if p['rp']=='Run']
        zpass=[p for p in zp if p['rp']=='Pass']
        ws3.row_dimensions[r].height=28
        sc(ws3,r,1,zcode,bold=True,sz=11,fc=CW,bg=zhdr)
        sc(ws3,r,2,len(zp),bold=True,sz=11,fc="FF000000",bg=zbg,fmt="0")
        sc(ws3,r,3,len(zr),bold=True,sz=11,fc="FF8B0000",bg="FFFDE8E8",fmt="0")
        sc(ws3,r,4,len(zpass),bold=True,sz=11,fc="FF00008B",bg="FFE8F0FE",fmt="0")
        t3f=top3(zp,'form'); t3fr=top3(zr,'form')
        for i,cn in enumerate([5,6,7]): sc(ws3,r,cn,t3f[i]['v']+f" ({t3f[i]['n']})" if i<len(t3f) else "—",sz=9,bg=zbg,wrap=True)
        for i,cn in enumerate([8,9]): sc(ws3,r,cn,t3fr[i]['v']+f" ({t3fr[i]['n']})" if i<len(t3fr) else "—",sz=9,bg=zbg,wrap=True)
        sc(ws3,r,10,"",bg=zbg,sz=9,wrap=True,h="left")
    ws3.freeze_panes="B4"

    # ── Tab 4: Personnel & Backfield ─────────────────────────
    ws4=wb2.create_sheet("4. Personnel & Backfield")
    ws4.sheet_properties.tabColor="1A5276"; ws4.sheet_view.showGridLines=False
    NC4=10; widths(ws4,[8,12,9,10,20,20,20,20,20,28])
    banner(ws4,1,"PERSONNEL & BACKFIELD  —  What sets they use by zone",NC4,bg=CBl,sz=12,ht=30)
    ws4.row_dimensions[2].height=20
    for c,txt,bg,span in[(1,"ZONE",CB,1),(2,"COUNTS",CB,3),(5,"TOP BACKFIELDS",CBl,3),(8,"TOP BACKFIELD vs PASS",CBl,2),(10,"NOTES",CB,1)]:
        hdr(ws4,2,c,txt,bg=bg,sz=8,span=span)
    ws4.row_dimensions[3].height=36
    for c,txt,bg in[(1,"Zone",CB),(2,"Plays",CB),(3,"Run",CB),(4,"Pass",CB),
                     (5,"#1 Backfield",CBl),(6,"#2 Backfield",CBl),(7,"#3 Backfield",CBl),
                     (8,"#1 vs Pass",CBl),(9,"#2 vs Pass",CBl),(10,"Notes",CB)]:
        hdr(ws4,3,c,txt,bg=bg,sz=9,wrap=True)
    for ri,zcode in enumerate(zone_list):
        r=ri+4; zbg=ZONE_BG[zcode]; zhdr=ZONE_HDR[zcode]
        zp=[p for p in plays if p['zone']==zcode]
        zr=[p for p in zp if p['rp']=='Run']
        zpass=[p for p in zp if p['rp']=='Pass']
        ws4.row_dimensions[r].height=28
        sc(ws4,r,1,zcode,bold=True,sz=11,fc=CW,bg=zhdr)
        sc(ws4,r,2,len(zp),bold=True,sz=11,fc="FF000000",bg=zbg,fmt="0")
        sc(ws4,r,3,len(zr),bold=True,sz=11,fc="FF8B0000",bg="FFFDE8E8",fmt="0")
        sc(ws4,r,4,len(zpass),bold=True,sz=11,fc="FF00008B",bg="FFE8F0FE",fmt="0")
        t3c=top3(zp,'backfield'); t3cp=top3(zpass,'backfield')
        for i,cn in enumerate([5,6,7]): sc(ws4,r,cn,t3c[i]['v']+f" ({t3c[i]['n']})" if i<len(t3c) else "—",sz=9,bg=zbg,wrap=True)
        for i,cn in enumerate([8,9]): sc(ws4,r,cn,t3cp[i]['v']+f" ({t3cp[i]['n']})" if i<len(t3cp) else "—",sz=9,bg=zbg,wrap=True)
        sc(ws4,r,10,"",bg=zbg,sz=9,wrap=True,h="left")
    ws4.freeze_panes="B4"

    # ── Tab 5: Motion & Strength ──────────────────────────────
    ws5=wb2.create_sheet("5. Motion & Strength")
    ws5.sheet_properties.tabColor="4A235A"; ws5.sheet_view.showGridLines=False
    NC5=10; widths(ws5,[8,10,10,10,18,18,18,14,14,28])
    banner(ws5,1,"MOTION & STRENGTH  —  When and where they use motion, and their formation strength",NC5,bg=CPu,sz=12,ht=30)
    ws5.row_dimensions[2].height=20
    for c,txt,bg,span in[(1,"ZONE",CB,1),(2,"COUNTS",CB,3),(5,"TOP MOTION TYPES",CPu,3),(8,"HASH MOTION",CPu,2),(10,"NOTES",CB,1)]:
        hdr(ws5,2,c,txt,bg=bg,sz=8,span=span)
    ws5.row_dimensions[3].height=36
    for c,txt,bg in[(1,"Zone",CB),(2,"Plays",CB),(3,"Motion\nCount",CB),(4,"Motion %",CB),
                     (5,"#1 Motion",CPu),(6,"#2 Motion",CPu),(7,"#3 Motion",CPu),
                     (8,"L Hash\nMotion%",CPu),(9,"R Hash\nMotion%",CPu),(10,"Notes",CB)]:
        hdr(ws5,3,c,txt,bg=bg,sz=9,wrap=True)
    for ri,zcode in enumerate(zone_list):
        r=ri+4; zbg=ZONE_BG[zcode]; zhdr=ZONE_HDR[zcode]
        zp=[p for p in plays if p['zone']==zcode]
        zb=[p for p in zp if p.get('motioned') is True]
        zl=[p for p in zp if p['hash']=='L']; zr_h=[p for p in zp if p['hash']=='R']
        zbl=[p for p in zl if p.get('motioned') is True]
        zbr=[p for p in zr_h if p.get('motioned') is True]
        ws5.row_dimensions[r].height=28
        sc(ws5,r,1,zcode,bold=True,sz=11,fc=CW,bg=zhdr)
        sc(ws5,r,2,len(zp),bold=True,sz=11,fc="FF000000",bg=zbg,fmt="0")
        sc(ws5,r,3,len(zb),bold=True,sz=11,fc="FF4A235A",bg="FFEDE7F6",fmt="0")
        sc(ws5,r,4,round(len(zb)/len(zp),2) if zp else "",bold=True,sz=11,fc="FF4A235A",bg="FFEDE7F6",fmt="0%")
        t3b=top3(zb,'motion')
        for i,cn in enumerate([5,6,7]): sc(ws5,r,cn,t3b[i]['v']+f" ({t3b[i]['n']})" if i<len(t3b) else "—",sz=9,bg=zbg,wrap=True)
        sc(ws5,r,8,round(len(zbl)/len(zl),2) if zl else "",sz=10,fc="FF6C3483",bg="FFEAF0FF",fmt="0%")
        sc(ws5,r,9,round(len(zbr)/len(zr_h),2) if zr_h else "",sz=10,fc="FF784212",bg="FFFFF0EA",fmt="0%")
        sc(ws5,r,10,"",bg=zbg,sz=9,wrap=True,h="left")
    ws5.freeze_panes="B4"

    # ── Tab: Formation / Backfield / Motion Tendencies ───────
    def _succ_rate(lst):
        v=[p for p in lst if p.get('succ') is not None]
        return round(sum(1 for p in v if p['succ'])/len(v)*100) if v else None

    def build_matchup_tab(ws, key, title, accent, min_n=3):
        """One row per offensive look, showing how the opposing offense has
        performed out of it — this is scouting film of THEIR offense, so the
        production belongs to them, not to us."""
        ws.sheet_view.showGridLines=False
        NC=12
        widths(ws,[20,8,8,9,9,9,9,9,10,10,20,20])
        banner(ws,1,title,NC,bg=accent,sz=12,ht=30)
        ws.row_dimensions[2].height=40
        for cn,txt,bg in[
            (1,"OFFENSIVE LOOK",CB),(2,"Snaps",CB),(3,"% of\nSnaps",CB),
            (4,"Run%",CR),(5,"Pass%",CBl),
            (6,"Yds/\nPlay",CTe),(7,"Success%",CTe),(8,"Expl%",CTe),
            (9,"Run Yds/\nPlay",CR),(10,"Pass Yds/\nPlay",CBl),
            (11,"Top Play Call",CPu),(12,"Most Seen On",CPu),
        ]:
            hdr(ws,2,cn,txt,bg=bg,sz=8,wrap=True)

        vals=Counter(p[key] for p in plays
                     if str(p.get(key,'')).strip() not in ('','nan','None','0'))
        ranked=[(v,n) for v,n in vals.most_common() if n>=min_n]
        total_all=len(plays)

        def dd_bucket(p):
            dn,dist=p['dn'],p['dist']
            if dn==1: return "1st & 10" if dist>=8 else "1st & Short"
            if dn==2: return "2nd & Long" if dist>=7 else ("2nd & Med" if dist>=4 else "2nd & Short")
            if dn==3: return "3rd & Long" if dist>=7 else ("3rd & Med" if dist>=4 else "3rd & Short")
            if dn==4: return "4th Down"
            return "—"

        for ri,(val,n) in enumerate(ranked):
            r=ri+3; ws.row_dimensions[r].height=26
            bg=CL if ri%2==0 else CW
            vp=[p for p in plays if p[key]==val]
            vruns=[p for p in vp if p['rp']=='Run']
            vpass=[p for p in vp if p['rp']=='Pass']
            avg=sum(p['gnls'] for p in vp)/len(vp) if vp else 0
            expl=len([p for p in vp if p['expl']])
            sr=_succ_rate(vp)
            small=" *" if n<5 else ""
            sc(ws,r,1,str(val)+small,bold=True,sz=9,fc=CW,bg=accent,h="left")
            sc(ws,r,2,n,bold=True,sz=10,fc="FF000000",bg=bg,fmt="0")
            sc(ws,r,3,round(n/total_all,2) if total_all else "",sz=9,fc="FF000000",bg=bg,fmt="0%")
            run_ratio  = (len(vruns)/len(vp)) if vp else None
            pass_ratio = (len(vpass)/len(vp)) if vp else None
            sc(ws,r,4,(round(run_ratio,3) if run_ratio is not None else ""),sz=10,fc="FF8B0000",bg=CRB,fmt="0%")
            sc(ws,r,5,(round(pass_ratio,3) if pass_ratio is not None else ""),sz=10,fc="FF00008B",bg=CPB,fmt="0%")
            sc(ws,r,6,round(avg,1),bold=True,sz=11,fc="FF0E7060",bg="FFE8F8E8",fmt="0.0")
            sc(ws,r,7,(sr/100 if sr is not None else "—"),bold=True,sz=10,fc="FF0E7060",
               bg="FFE8F8E8",fmt="0%" if sr is not None else "General")
            sc(ws,r,8,(round(expl/len(vp),3) if vp else ""),sz=10,fc="FF0E7060",bg="FFE8F8E8",fmt="0%")
            sc(ws,r,9,round(sum(p['gnls'] for p in vruns)/len(vruns),1) if vruns else "—",
               sz=9,fc="FF8B0000",bg=CRB,fmt="0.0" if vruns else "General")
            sc(ws,r,10,round(sum(p['gnls'] for p in vpass)/len(vpass),1) if vpass else "—",
               sz=9,fc="FF00008B",bg=CPB,fmt="0.0" if vpass else "General")
            # top play call out of this look
            pgroups={}
            for p in vp:
                pl=p['play']
                if str(pl).strip() in ('','nan','None'): continue
                pgroups.setdefault(pl,[]).append(p)
            best=[(pl,len(g)) for pl,g in pgroups.items()]
            best.sort(key=lambda t:-t[1])
            sc(ws,r,11,f"{best[0][0]} ({best[0][1]})" if best else "—",
               sz=8,fc="FF4A235A",bg="FFEDE7F6",h="left",wrap=True)
            dd_top=Counter(dd_bucket(p) for p in vp).most_common(1)
            sc(ws,r,12,f"{dd_top[0][0]} ({dd_top[0][1]})" if dd_top else "—",
               sz=8,fc="FF4A235A",bg="FFEDE7F6",h="left",wrap=True)

        if not ranked:
            ws.merge_cells(f"A3:{gcl(NC)}3")
            c=ws.cell(row=3,column=1,value="Not enough tagged data — check this column is filled in your Hudl export.")
            c.font=Font(name=FN,sz=10,italic=True,color=CDG); c.alignment=Alignment(horizontal="center")
        fr=len(ranked)+4
        ws.cell(row=fr,column=1,value="* = small sample (under 5 snaps) — read with caution").font=Font(name=FN,sz=8,italic=True,color=CDG)
        ws.freeze_panes="B3"

    ws_fm=wb2.create_sheet("6. Formation Tendencies")
    ws_fm.sheet_properties.tabColor="8B0000"
    build_matchup_tab(ws_fm,'form',"FORMATION TENDENCIES  —  What each formation has produced","FF8B0000")

    ws_cm=wb2.create_sheet("7. Backfield Tendencies")
    ws_cm.sheet_properties.tabColor="00008B"
    build_matchup_tab(ws_cm,'backfield',"BACKFIELD TENDENCIES  —  What each backfield set has produced","FF00008B")

    ws_bm=wb2.create_sheet("8. Motion Tendencies")
    ws_bm.sheet_properties.tabColor="4A235A"
    build_matchup_tab(ws_bm,'motion',"MOTION TENDENCIES  —  What their motion looks have produced","FF4A235A")

    # ── Tab: Motion vs No-Motion summary ─────────────────────
    ws_bs=wb2.create_sheet("9. Motion Summary")
    ws_bs.sheet_properties.tabColor="7B241C"; ws_bs.sheet_view.showGridLines=False
    widths(ws_bs,[24,10,10,10,10,10,22,22])
    banner(ws_bs,1,"MOTION SUMMARY  —  Motion vs No Motion",8,bg="FF7B241C",sz=12,ht=30)
    ws_bs.row_dimensions[2].height=36
    for cn,txt in [(1,"SITUATION"),(2,"Snaps"),(3,"Rate"),(4,"Yds/Play"),
                   (5,"Success%"),(6,"Expl%"),(7,"Top Play Call"),(8,"Most Seen On")]:
        hdr(ws_bs,2,cn,txt,bg=CB,sz=8,wrap=True)
    motioned=[p for p in plays if p.get('motioned') is True]
    nonmotion=[p for p in plays if p.get('motioned') is False]
    known=len(motioned)+len(nonmotion)
    def _dd_b(p):
        dn,dist=p['dn'],p['dist']
        if dn==1: return "1st & 10" if dist>=8 else "1st & Short"
        if dn==2: return "2nd & Long" if dist>=7 else ("2nd & Med" if dist>=4 else "2nd & Short")
        if dn==3: return "3rd & Long" if dist>=7 else ("3rd & Med" if dist>=4 else "3rd & Short")
        if dn==4: return "4th Down"
        return "—"
    rows_bs=[("WHEN THEY MOTION",motioned,"FF7B241C"),("WHEN THEY DON'T",nonmotion,"FF0E7060")]
    for ri,(lbl,grp,color) in enumerate(rows_bs):
        r=ri+3; ws_bs.row_dimensions[r].height=26
        bg=CL if ri%2==0 else CW
        sc(ws_bs,r,1,lbl,bold=True,sz=10,fc=CW,bg=color,h="left")
        sc(ws_bs,r,2,len(grp),bold=True,sz=11,fc="FF000000",bg=bg,fmt="0")
        sc(ws_bs,r,3,round(len(grp)/known,2) if known else "",bold=True,sz=11,fc="FF000000",bg=bg,fmt="0%")
        sc(ws_bs,r,4,round(sum(p['gnls'] for p in grp)/len(grp),1) if grp else "—",
           bold=True,sz=11,fc="FF0E7060",bg="FFE8F8E8",fmt="0.0" if grp else "General")
        sr=_succ_rate(grp)
        sc(ws_bs,r,5,(sr/100 if sr is not None else "—"),bold=True,sz=10,fc="FF0E7060",
           bg="FFE8F8E8",fmt="0%" if sr is not None else "General")
        sc(ws_bs,r,6,round(len([p for p in grp if p['expl']])/len(grp),2) if grp else "",
           sz=10,fc="FF0E7060",bg="FFE8F8E8",fmt="0%")
        pgroups={}
        for p in grp:
            pl=p['play']
            if str(pl).strip() in ('','nan','None'): continue
            pgroups.setdefault(pl,[]).append(p)
        best=[(pl,len(g)) for pl,g in pgroups.items()]
        best.sort(key=lambda t:-t[1])
        sc(ws_bs,r,7,f"{best[0][0]} ({best[0][1]})" if best else "—",sz=8,bg=bg,h="left",wrap=True)
        ddt=Counter(_dd_b(p) for p in grp).most_common(1)
        sc(ws_bs,r,8,f"{ddt[0][0]} ({ddt[0][1]})" if ddt else "—",sz=8,bg=bg,h="left",wrap=True)
    # motion-by-situation breakdown
    r=6
    ws_bs.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)
    c=ws_bs.cell(row=r,column=1,value="WHEN DO THEY USE MOTION?")
    c.font=Font(name=FN,bold=True,sz=11,color=CW); c.fill=fil(CB); c.alignment=Alignment(horizontal="center")
    ws_bs.row_dimensions[r].height=20
    r+=1
    for cn,txt in [(1,"SITUATION"),(2,"Snaps"),(3,"Motion%"),(4,"Yds/Play"),(5,"Success%"),(6,"Expl%"),(7,""),(8,"")]:
        hdr(ws_bs,r,cn,txt,bg=CB,sz=8,wrap=True)
    r+=1
    sits_b=[("1st & 10",lambda p:p['dn']==1 and p['dist']>=8),
            ("2nd & Long",lambda p:p['dn']==2 and p['dist']>=7),
            ("2nd & Short",lambda p:p['dn']==2 and p['dist']<=3),
            ("3rd & Long",lambda p:p['dn']==3 and p['dist']>=7),
            ("3rd & Med",lambda p:p['dn']==3 and 4<=p['dist']<=6),
            ("3rd & Short",lambda p:p['dn']==3 and p['dist']<=3),
            ("Red Zone",lambda p:p['zone']=='RZ'),
            ("Goal Line",lambda p:p['zone']=='GL')]
    for ri,(lbl,fn) in enumerate(sits_b):
        rr=r+ri; ws_bs.row_dimensions[rr].height=18
        bg=CL if ri%2==0 else CW
        sp=[p for p in plays if fn(p)]
        spk=[p for p in sp if p.get('motioned') is not None]
        sb=[p for p in spk if p['motioned'] is True]
        sc(ws_bs,rr,1,lbl,bold=True,sz=9,fc=CW,bg=CBl,h="left")
        sc(ws_bs,rr,2,len(sp),sz=9,fc="FF000000",bg=bg,fmt="0")
        sc(ws_bs,rr,3,round(len(sb)/len(spk),2) if spk else "",bold=True,sz=10,
           fc="FF7B241C",bg=CRB,fmt="0%")
        sc(ws_bs,rr,4,round(sum(p['gnls'] for p in sp)/len(sp),1) if sp else "—",
           sz=9,bg=bg,fmt="0.0" if sp else "General")
        s2=_succ_rate(sp)
        sc(ws_bs,rr,5,(s2/100 if s2 is not None else "—"),sz=9,fc="FF0E7060",bg="FFE8F8E8",
           fmt="0%" if s2 is not None else "General")
        sc(ws_bs,rr,6,round(len([p for p in sp if p['expl']])/len(sp),2) if sp else "",
           sz=9,fc="FF0E7060",bg="FFE8F8E8",fmt="0%")
        sc(ws_bs,rr,7,"",bg=bg); sc(ws_bs,rr,8,"",bg=bg)

    # ── Tab 6: Hash Tendencies ───────────────────────────────

    ws6=wb2.create_sheet("10. Hash Tendencies")
    ws6.sheet_properties.tabColor="6C3483"; ws6.sheet_view.showGridLines=False
    NC6=13; widths(ws6,[18,10,10,10,10,10,10,18,18,18,18,18,18])
    banner(ws6,1,"HASH TENDENCIES  —  Left · Middle · Right",NC6,bg="FF6C3483",sz=13,ht=32)
    ws6.row_dimensions[2].height=36
    for cn,txt,bg in[(1,"FIELD ZONE",CB),(2,"L Plays","FF6C3483"),(3,"L Run%","FF6C3483"),(4,"L Pass%","FF6C3483"),
                      (5,"M Plays","FF1A5276"),(6,"M Run%","FF1A5276"),(7,"M Pass%","FF1A5276"),
                      (8,"R Plays","FF784212"),(9,"R Run%","FF784212"),(10,"R Pass%","FF784212"),
                      (11,"Top L Form",CB),(12,"Top M Form",CB),(13,"Top R Form",CB)]:
        hdr(ws6,2,cn,txt,bg=bg,sz=8)

    def hash_row(ws,r,label,base,zhdr_col,zbg):
        ws.row_dimensions[r].height=24
        sc(ws,r,1,label,bold=True,sz=10,fc=CW,bg=zhdr_col)
        for h,cols in[('L',(2,3,4)),('M',(5,6,7)),('R',(8,9,10))]:
            hp=[p for p in base if p['hash']==h]
            hr=[p for p in hp if p['rp']=='Run']
            hpass=[p for p in hp if p['rp']=='Pass']
            n=len(hp)
            sc(ws,r,cols[0],n,sz=11,bold=True,fc="FF000000",bg=zbg,fmt="0")
            sc(ws,r,cols[1],round(len(hr)/n,2) if n>0 else "",sz=11,bold=True,fc="FF8B0000",bg=zbg,fmt="0%")
            sc(ws,r,cols[2],round(len(hpass)/n,2) if n>0 else "",sz=11,bold=True,fc="FF00008B",bg=zbg,fmt="0%")
        for col_n,h in[(11,'L'),(12,'M'),(13,'R')]:
            hp=[p for p in base if p['hash']==h]
            tf=top3(hp,'form')
            sc(ws,r,col_n,tf[0]['v']+f" ({tf[0]['n']})" if tf else "—",sz=9,bg=zbg,h="left",wrap=True)

    hash_row(ws6,3,"OVERALL",plays,CB,CL)
    for ri,zcode in enumerate(zone_list):
        hash_row(ws6,ri+4,f"{zcode} — {zone_names[zcode]}",[p for p in plays if p['zone']==zcode],ZONE_HDR[zcode],ZONE_BG[zcode])
    ws6.freeze_panes="A3"


    # ── Tab 7: Down & Distance Tendencies ────────────────────
    ws_dd=wb2.create_sheet("11. Down & Distance")
    ws_dd.sheet_properties.tabColor="F1C40F"
    ws_dd.sheet_view.showGridLines=False
    NC_DD=13
    widths(ws_dd,[20,8,8,8,20,20,20,20,20,20,16,16,16])
    banner(ws_dd,1,"DOWN & DISTANCE  —  Top Formations, Backfields & Motion by Situation",NC_DD,bg=CB,sz=12,ht=32)

    ws_dd.row_dimensions[2].height=38
    for cn,txt,bg in[
        (1,"SITUATION",CB),(2,"Plays",CB),(3,"Run%",CB),(4,"Pass%",CB),
        (5,"#1 Formation",CR),(6,"#2 Formation",CR),(7,"#3 Formation",CR),
        (8,"#1 Backfield",CBl),(9,"#2 Backfield",CBl),(10,"#3 Backfield",CBl),
        (11,"Motion%",CPu),(12,"#1 Motion Type",CPu),(13,"#2 Motion Type",CPu),
    ]:
        hdr(ws_dd,2,cn,txt,bg=bg,sz=9,wrap=True)

    dd_sits=[
        ("1ST & 10",       lambda p: p['dn']==1 and p['dist']>=8),
        ("1ST & SHORT",    lambda p: p['dn']==1 and p['dist']<8),
        ("2ND & LONG",     lambda p: p['dn']==2 and p['dist']>=7),
        ("2ND & MEDIUM",   lambda p: p['dn']==2 and 4<=p['dist']<=6),
        ("2ND & SHORT",    lambda p: p['dn']==2 and p['dist']<=3),
        ("3RD & LONG",     lambda p: p['dn']==3 and p['dist']>=7),
        ("3RD & MEDIUM",   lambda p: p['dn']==3 and 4<=p['dist']<=6),
        ("3RD & SHORT",    lambda p: p['dn']==3 and p['dist']<=3),
        ("4TH DOWN",       lambda p: p['dn']==4),
        ("RED ZONE",       lambda p: p['zone']=='RZ'),
        ("GOAL LINE",      lambda p: p['zone']=='GL'),
        ("BACKED UP",      lambda p: p['zone']=='BZ'),
    ]
    dd_colors=[CTe,CTe,CBl,CBl,CBl,CR,CR,CR,"FF7B241C",CR,CPu,CTe]

    def top3_dd(play_list, key):
        vals=[str(p.get(key,'')) for p in play_list
              if p.get(key) not in(None,'','nan','None')
              and str(p.get(key,'')).strip() not in('','nan','None')]
        if not vals: return ["—","—","—"]
        counts=Counter(vals).most_common(3)
        result=[f"{v} ({n})" for v,n in counts]
        while len(result)<3: result.append("—")
        return result

    for ri,(lbl,fn) in enumerate(dd_sits):
        r=ri+3; ws_dd.row_dimensions[r].height=30
        color=dd_colors[ri]
        bg=CRB if ri%2==0 else CL
        try: sp=[p for p in plays if fn(p)]
        except: sp=[]
        n_run=len([p for p in sp if p['rp']=='Run'])
        n_pass=len([p for p in sp if p['rp']=='Pass'])
        total=len(sp)
        motion_plays=[p for p in sp if p.get('motion','') not in('','nan','None','0','No Motion')]

        sc(ws_dd,r,1,lbl,bold=True,sz=10,fc=CW,bg=color,h="left")
        sc(ws_dd,r,2,total,bold=True,sz=11,fc="FF000000",bg=bg,fmt="0")
        sc(ws_dd,r,3,round(n_run/total,2) if total>0 else "",bold=True,sz=12,fc="FF8B0000",bg=CRB,fmt="0%")
        sc(ws_dd,r,4,round(n_pass/total,2) if total>0 else "",bold=True,sz=12,fc="FF00008B",bg=CPB,fmt="0%")

        # Top formations
        t3f=top3_dd(sp,'form')
        for i,cn in enumerate([5,6,7]):
            c=ws_dd.cell(row=r,column=cn,value=t3f[i])
            c.font=Font(name=FN,sz=9,color="FF8B0000" if t3f[i]!="—" else CDG)
            c.fill=fil(CRB); c.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True); c.border=bdr()

        # Top backfields
        t3c=top3_dd(sp,'backfield')
        for i,cn in enumerate([8,9,10]):
            c=ws_dd.cell(row=r,column=cn,value=t3c[i])
            c.font=Font(name=FN,sz=9,color="FF00008B" if t3c[i]!="—" else CDG)
            c.fill=fil(CPB); c.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True); c.border=bdr()

        # Motion % + top motion types
        motion_pct=round(len(motion_plays)/total,2) if total>0 else ""
        sc(ws_dd,r,11,motion_pct,bold=True,sz=12,fc="FF4A235A",bg="FFEDE7F6",fmt="0%")
        t3b=top3_dd(motion_plays,'motion')
        for i,cn in enumerate([12,13]):
            c=ws_dd.cell(row=r,column=cn,value=t3b[i])
            c.font=Font(name=FN,sz=9,color="FF4A235A" if t3b[i]!="—" else CDG)
            c.fill=fil("FFEDE7F6"); c.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True); c.border=bdr()

    ws_dd.freeze_panes="B3"

    # ── Tab 8: Situational Summary ───────────────────────────
    ws7=wb2.create_sheet("12. Situational Summary")
    ws7.sheet_properties.tabColor="4A235A"; ws7.sheet_view.showGridLines=False
    NC7=11; widths(ws7,[16,9,9,18,18,18,18,18,18,18,28])
    banner(ws7,1,"SITUATIONAL SUMMARY  —  What they show in every situation",NC7,bg="FF4A235A",sz=12,ht=30)
    s7h=["Situation","Run\nCount","Pass\nCount","Top Formation","Top Backfield",
         "Motion %","L Hash\nRun%","M Hash\nRun%","R Hash\nRun%","Top Play", "Notes"]
    ws7.row_dimensions[2].height=38
    for ci,h in enumerate(s7h): hdr(ws7,2,ci+1,h,bg="FF4A235A",sz=9,wrap=True)

    def sit(dn=None,dmin=None,dmax=None,zone=None):
        out=[]
        for p in plays:
            if dn   and p['dn']!=dn:     continue
            if dmin and p['dist']<dmin:  continue
            if dmax and p['dist']>dmax:  continue
            if zone and p['zone']!=zone: continue
            out.append(p)
        return out

    sits=[
        ("1ST DOWN",      dict(dn=1)),
        ("2ND & LONG",    dict(dn=2,dmin=7)),
        ("2ND & MEDIUM",  dict(dn=2,dmin=4,dmax=6)),
        ("2ND & SHORT",   dict(dn=2,dmax=3)),
        ("3RD & LONG",    dict(dn=3,dmin=7)),
        ("3RD & MEDIUM",  dict(dn=3,dmin=4,dmax=6)),
        ("3RD & SHORT",   dict(dn=3,dmax=3)),
        ("4TH DOWN",      dict(dn=4)),
        ("RED ZONE",      dict(zone="RZ")),
        ("GOAL LINE",     dict(zone="GL")),
        ("BACKED UP",     dict(zone="BZ")),
        ("COMING OUT",    dict(zone="OF")),
        ("TWO-MINUTE",    dict(dn=3,dmin=5)),
        ("MUST HAVE",     dict(dn=4)),
    ]
    sit_colors=["FF0E7060","FF1A5276","FF1A5276","FF1A5276",
                "FF7B241C","FF7B241C","FF7B241C","FF7B241C",
                "FF7B241C","FF4A235A","FF0E7060","FF0E7060","FF7D6608","FF16213E"]

    for ri,((lbl,args),color) in enumerate(zip(sits,sit_colors)):
        r=ri+3; ws7.row_dimensions[r].height=34
        sc(ws7,r,1,lbl,bold=True,sz=9,fc=CW,bg=color)
        sp=sit(**args)
        sr=[p for p in sp if p['rp']=='Run']; spass=[p for p in sp if p['rp']=='Pass']
        motion_p=[p for p in sp if p.get('motioned') is True]
        l_p=[p for p in sp if p['hash']=='L']; m_p=[p for p in sp if p['hash']=='M']; r_p=[p for p in sp if p['hash']=='R']
        l_r=len([p for p in l_p if p['rp']=='Run']); m_r=len([p for p in m_p if p['rp']=='Run']); r_r=len([p for p in r_p if p['rp']=='Run'])
        tf=top3(sp,'form'); tc=top3(sp,'backfield'); tp=top3(sp,'play')
        sc(ws7,r,2,len(sr),bold=True,sz=12,fc="FF8B0000",bg="FFFDE8E8",fmt="0")
        sc(ws7,r,3,len(spass),bold=True,sz=12,fc="FF00008B",bg="FFE8F0FE",fmt="0")
        sc(ws7,r,4,tf[0]['v']+f" ({tf[0]['n']})" if tf else "—",sz=9,bg="FFFDE8E8",wrap=True,h="left")
        sc(ws7,r,5,tc[0]['v']+f" ({tc[0]['n']})" if tc else "—",sz=9,bg="FFE8F0FE",wrap=True,h="left")
        sc(ws7,r,6,round(len(motion_p)/len(sp),2) if sp else "",sz=10,fc="FF4A235A",bg="FFEDE7F6",fmt="0%")
        sc(ws7,r,7,round(l_r/len(l_p),2) if l_p else "",sz=10,fc="FF6C3483",bg="FFEAF0FF",fmt="0%")
        sc(ws7,r,8,round(m_r/len(m_p),2) if m_p else "",sz=10,fc="FF1A5276",bg="FFE8F0FE",fmt="0%")
        sc(ws7,r,9,round(r_r/len(r_p),2) if r_p else "",sz=10,fc="FF784212",bg="FFFFF0EA",fmt="0%")
        sc(ws7,r,10,tp[0]['v']+f" ({tp[0]['n']})" if tp else "—",sz=9,bg=CL,wrap=True,h="left")
        sc(ws7,r,11,"",bg=CL if ri%2==0 else CW,sz=9,wrap=True,v="top")
    ws7.freeze_panes="D3"

    # ── Tab 8: DC Call Sheet Builder ─────────────────────────
    ws8=wb2.create_sheet("13. DC Call Sheet Builder")
    ws8.sheet_properties.tabColor="0E7060"; ws8.sheet_view.showGridLines=False
    NC8=11; widths(ws8,[20,8,18,18,10,9,20,20,18,9,26])
    banner(ws8,1,"DC CALL SHEET BUILDER  —  Left side auto-filled from film · Yellow = your calls",NC8,bg=CTe,sz=12,ht=30)
    c8h=["Situation","Snaps","Expected Formation","Expected Backfield","Motion %","Their\nSuccess%",
         "Run Fit / Front","Coverage Call","Pressure Package","Priority","Notes"]
    ws8.row_dimensions[2].height=38
    for ci,h in enumerate(c8h): hdr(ws8,2,ci+1,h,bg=CTe,sz=9,wrap=True)
    def _cs_sr(lst):
        v=[p for p in lst if p.get('succ') is not None]
        return round(sum(1 for p in v if p['succ'])/len(v)*100) if v else None
    def _cs_top(lst,key):
        vals=[str(p.get(key,'')) for p in lst if str(p.get(key,'')).strip() not in ('','nan','None','0')]
        c=Counter(vals).most_common(1)
        return c[0][0] if c else "—"

    # Situations mapped to real filters so we can PRE-FILL what the offense does
    cs_filters=[
        ("1st & 10",            lambda p: p['dn']==1 and p['dist']>=8),
        ("1st & 10 (Own Half)", lambda p: p['dn']==1 and p['dist']>=8 and p['zone'] in ('BZ','OF')),
        ("1st & 10 (Opp Half)", lambda p: p['dn']==1 and p['dist']>=8 and p['zone'] in ('MF','FZ')),
        ("2nd & Long (8+)",     lambda p: p['dn']==2 and p['dist']>=8),
        ("2nd & Medium (4-7)",  lambda p: p['dn']==2 and 4<=p['dist']<=7),
        ("2nd & Short (1-3)",   lambda p: p['dn']==2 and p['dist']<=3),
        ("3rd & Long (7+)",     lambda p: p['dn']==3 and p['dist']>=7),
        ("3rd & Medium (4-6)",  lambda p: p['dn']==3 and 4<=p['dist']<=6),
        ("3rd & Short (1-3)",   lambda p: p['dn']==3 and p['dist']<=3),
        ("4th Down",            lambda p: p['dn']==4),
        ("Red Zone",            lambda p: p['zone']=='RZ'),
        ("Goal Line",           lambda p: p['zone']=='GL'),
        ("Backed Up",           lambda p: p['zone']=='BZ'),
        ("Coming Out",          lambda p: p['zone']=='OF'),
        ("Midfield",            lambda p: p['zone']=='MF'),
        ("Fringe",              lambda p: p['zone']=='FZ'),
        ("Left Hash",           lambda p: p['hash']=='L'),
        ("Middle of Field",     lambda p: p['hash']=='M'),
        ("Right Hash",          lambda p: p['hash']=='R'),
    ]
    blank_rows=["Two-Minute (Ahead)","Two-Minute (Behind)","Must-Stop Plays",
                "Two-Point Defense","Overtime","Openers","Shot Play Alert"]

    r=3
    for ri,(lbl,fn) in enumerate(cs_filters):
        ws8.row_dimensions[r].height=22
        bg="FFF0FFF0" if ri%2==0 else CW
        try: sp=[p for p in plays if fn(p)]
        except: sp=[]
        spk=[p for p in sp if p.get('motioned') is not None]
        sb=[p for p in spk if p['motioned'] is True]
        motion_txt=f"{round(len(sb)/len(spk)*100)}%" if spk else "—"
        sr=_cs_sr(sp)
        sc(ws8,r,1,lbl,bold=True,sz=9,fc=CW,bg=CGr,h="left")
        sc(ws8,r,2,len(sp),sz=9,bg=bg,fmt="0")
        sc(ws8,r,3,_cs_top(sp,'form'),sz=8,bg=bg,h="left",wrap=True)
        sc(ws8,r,4,_cs_top(sp,'backfield'),sz=8,bg=bg,h="left",wrap=True)
        sc(ws8,r,5,motion_txt,bold=True,sz=9,fc="FF7B241C",bg="FFFDE8E8")
        sc(ws8,r,6,(sr/100 if sr is not None else "—"),sz=9,fc="FF0E7060",bg="FFE8F8E8",
           fmt="0%" if sr is not None else "General")
        # blank columns for the DC's own calls
        for ci in range(7,NC8+1):
            sc(ws8,r,ci,"",bg="FFFFFBE6",sz=9,wrap=True,v="top")
        r+=1
    # extra blank situations the DC fills in entirely
    for ri,lbl in enumerate(blank_rows):
        ws8.row_dimensions[r].height=22
        sc(ws8,r,1,lbl,bold=True,sz=9,fc=CW,bg=CBl,h="left")
        for ci in range(2,NC8+1):
            sc(ws8,r,ci,"",bg="FFFFFBE6",sz=9,wrap=True,v="top")
        r+=1
    # spare rows
    for ri in range(8):
        ws8.row_dimensions[r].height=22
        for ci in range(1,NC8+1):
            sc(ws8,r,ci,"",bg=CW if ri%2 else "FFF7F7F7",sz=9,wrap=True,v="top")
        r+=1
    ws8.freeze_panes="B3"

    # ── Tab 9: Coordinator Summary ───────────────────────────
    ws9=wb2.create_sheet("14. Coordinator Summary")
    ws9.sheet_properties.tabColor="F1C40F"; ws9.sheet_view.showGridLines=False
    ws9.page_setup.paperSize=1; ws9.page_setup.orientation="landscape"
    ws9.page_setup.fitToPage=True; ws9.page_setup.fitToWidth=1; ws9.page_setup.fitToHeight=1
    widths(ws9,[22,32,3,22,32,3,22,32])
    banner(ws9,1,"DEFENSIVE COORDINATOR SUMMARY  ·  PRINT LANDSCAPE",8,bg=CB,sz=14,ht=36)
    ws9.row_dimensions[2].height=18
    for lbl,ci in[("OPP:",1),("WEEK:",3),("DATE:",5),("DC:",7)]:
        sc(ws9,2,ci,lbl,bold=True,sz=9,fc=CW,bg=CBl,h="right")
        val={"OPP:":opp,"WEEK:":week,"DATE:":date}.get(lbl,"")
        sc(ws9,2,ci+1,val,bg=CL,sz=9)
    for r in range(3,66):
        ws9.row_dimensions[r].height=16
        for dc in[3,6]: ws9.cell(row=r,column=dc).fill=fil("FFCCCCCC")

    rz=[p for p in plays if p['zone']=='RZ']; gl=[p for p in plays if p['zone']=='GL']
    t3d=[p for p in plays if p['dn']==3]
    all_motion=[p for p in plays if p.get('motioned') is True]

    def blk(ws,sr,ca,cb,title,data,hc):
        ws.row_dimensions[sr].height=17
        ws.merge_cells(start_row=sr,start_column=ca,end_row=sr,end_column=cb)
        c=ws.cell(row=sr,column=ca,value=title)
        c.font=Font(name=FN,bold=True,sz=9,color=CW); c.fill=fil(hc)
        c.alignment=Alignment(horizontal="center",vertical="center"); c.border=bdr()
        for i,(lbl,val) in enumerate(data):
            rr=sr+1+i; bg=CL if i%2==0 else CW; ws.row_dimensions[rr].height=16
            la=ws.cell(row=rr,column=ca,value=lbl)
            la.font=Font(name=FN,bold=True,sz=8,color=CDG)
            la.fill=fil(bg); la.alignment=Alignment(horizontal="right",vertical="center"); la.border=bdr()
            vb=ws.cell(row=rr,column=cb,value=val)
            vb.font=Font(name=FN,sz=9,color="FF000000")
            vb.fill=fil(bg); vb.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True); vb.border=bdr()
            if ca+1<cb:
                ws.merge_cells(start_row=rr,start_column=ca+1,end_row=rr,end_column=cb-1)
                for mc in range(ca+1,cb): ws.cell(row=rr,column=mc).fill=fil(bg)
        return sr+1+len(data)

    r=3
    r=blk(ws9,r,1,2,"SNAP COUNTS",[
        ("Total Plays",total),("Total Runs",len(runs)),("Total Passes",len(passes)),
        ("Run %",f"{pct(len(runs),total)}%"),("Pass %",f"{pct(len(passes),total)}%"),
    ],CB)
    r+=1
    r=blk(ws9,r,1,2,"MOTION OVERVIEW",[
        ("Total Motion Plays",len(all_motion)),
        ("Motion %",f"{pct(len(all_motion),total)}%"),
        ("RZ Motion %",f"{pct(len([p for p in rz if p.get('motioned') is True]),len([p for p in rz if p.get('motioned') is not None]))}%"),
        ("3rd Down Motion %",f"{pct(len([p for p in t3d if p.get('motioned') is True]),len([p for p in t3d if p.get('motioned') is not None]))}%"),
    ],"FF4A235A")
    r+=1
    blk(ws9,r,1,2,"RED ZONE / GL",[
        ("RZ Runs",len([p for p in rz if p['rp']=='Run'])),
        ("RZ Passes",len([p for p in rz if p['rp']=='Pass'])),
        ("GL Runs",len([p for p in gl if p['rp']=='Run'])),
        ("GL Passes",len([p for p in gl if p['rp']=='Pass'])),
    ],CR)

    r=3
    r=blk(ws9,r,4,5,"TOP FORMATIONS PLAYED",[(f"#{i+1}",x['v']+f" ({x['n']})") for i,x in enumerate(top3(plays,'form'))],CR)
    r+=1
    r=blk(ws9,r,4,5,"TOP BACKFIELDS FACED",[(f"#{i+1}",x['v']+f" ({x['n']})") for i,x in enumerate(top3(plays,'backfield'))],CBl)
    r+=1
    blk(ws9,r,4,5,"TOP MOTION TYPES",[(f"#{i+1}",x['v']+f" ({x['n']})") for i,x in enumerate(top3(all_motion,'motion'))],CPu)

    r=3
    r=blk(ws9,r,7,8,"BEST RUN FITS vs THEIR O  (fill in)",[("1.",""),("2.",""),("3.",""),("4.",""),("5.","")],CTe)
    r+=1
    r=blk(ws9,r,7,8,"BEST COVERAGE CALLS  (fill in)",[("1.",""),("2.",""),("3.",""),("4.",""),("5.","")],CBl)
    r+=1
    r=blk(ws9,r,7,8,"PRESSURES THAT WORK  (fill in)",[("1.",""),("2.",""),("3.",""),("4.","")],CPu)
    r+=1
    blk(ws9,r,7,8,"MUST-STOP PLAYS  (fill in)",[("Call 1.",""),("Call 2.",""),("Call 3.",""),("Call 4.","")],CB)

    buf=io.BytesIO(); wb2.save(buf); buf.seek(0)
    return buf.getvalue()

# ── HTML Report ───────────────────────────────────────────────
def build_html(plays, opp, week, date):
    zc={"BZ":"#7b241c","OF":"#1a5276","MF":"#0e7060","FZ":"#7d6608","RZ":"#7b241c","GL":"#4a235a"}
    zn={"BZ":"Backed Up","OF":"Open Field","MF":"Midfield","FZ":"Fringe","RZ":"Red Zone","GL":"Goal Line"}
    zone_list=["BZ","OF","MF","FZ","RZ","GL"]
    total=len(plays); runs=[p for p in plays if p['rp']=='Run']; passes=[p for p in plays if p['rp']=='Pass']
    all_motion=[p for p in plays if p.get('motioned') is True]
    rz=[p for p in plays if p['zone']=='RZ']; gl=[p for p in plays if p['zone']=='GL']

    def tags(items,cls=''):
        if not items: return '<span class="ctag">—</span>'
        return ''.join(f'<span class="ctag {cls}">{x["v"]} ({x["n"]})</span>' for x in items)

    zone_cards=''
    for z in zone_list:
        zp=[p for p in plays if p['zone']==z]
        if not zp: continue
        zr2=[p for p in zp if p['rp']=='Run']; zpas=[p for p in zp if p['rp']=='Pass']
        rp=pct(len(zr2),len(zp)); pp=pct(len(zpas),len(zp))
        zb=[p for p in zp if p.get('motioned') is True]
        zone_cards+=f'''<div class="zone-card">
          <div class="zone-hdr" style="background:{zc[z]}20;border-bottom:2px solid {zc[z]}">
            <div><div class="zone-badge" style="color:{zc[z]}">{z}</div><div class="zone-sub">{zn[z]}</div></div>
            <div class="zone-plays">{len(zp)} plays</div>
          </div>
          <div class="zone-body">
            <div class="bar-row">
              <div class="bar-labels"><span style="color:#e8a095">RUN {rp}%</span><span style="color:#93d4f0">PASS {pp}%</span></div>
              <div class="bar-bg"><div class="bar-fill" style="background:#7b241c;width:{rp}%"></div></div>
            </div>
            <div class="zone-tags">
              <div class="tag-lbl">Top Formations</div>{tags(top3(zp,"form"),"f")}
              <div class="tag-lbl" style="margin-top:6px">Top Backfields</div>{tags(top3(zp,"backfield"),"c")}
              <div class="tag-lbl" style="margin-top:6px">Motion % — {pct(len(zb),len(zp))}%</div>{tags(top3(zb,"motion"),"b")}
            </div>
          </div>
        </div>'''

    hash_cards=''
    for h,lbl,cls,color in[('L','Left Hash','hl','#b388d4'),('M','Middle','hm','#5dade2'),('R','Right Hash','hr','#e59866')]:
        hp=[p for p in plays if p['hash']==h]
        if not hp: hash_cards+=f'<div class="hash-card {cls}"><div class="hc-title" style="color:{color}">{lbl}</div><p style="color:rgba(240,237,232,.25);text-align:center;font-size:12px">No data</p></div>'; continue
        hr2=[p for p in hp if p['rp']=='Run']; hpass=[p for p in hp if p['rp']=='Pass']
        hb=[p for p in hp if p.get('motioned') is True]
        tf=top3(hp,'form')
        hash_cards+=f'''<div class="hash-card {cls}">
          <div class="hc-title" style="color:{color}">{lbl}</div>
          <div class="hbig" style="color:{color}">{len(hp)}</div>
          <div class="hsub">total plays</div>
          <div class="hrp">
            <div class="hrp-item"><div class="hrp-lbl">Run %</div><div class="hrp-val" style="color:#7b241c">{pct(len(hr2),len(hp))}%</div></div>
            <div class="hrp-item"><div class="hrp-lbl">Motion %</div><div class="hrp-val" style="color:#b388d4">{pct(len(hb),len(hp))}%</div></div>
          </div>
          <div class="htc">Top Formation: <span style="color:#d4a017">{tf[0]["v"]+" ("+str(tf[0]["n"])+")" if tf else "—"}</span></div>
        </div>'''

    sit_rows=''
    sits2=[("1ST DOWN",lambda p:p['dn']==1),("2ND & LONG",lambda p:p['dn']==2 and p['dist']>=7),
           ("2ND & MED",lambda p:p['dn']==2 and 4<=p['dist']<=6),("2ND & SHORT",lambda p:p['dn']==2 and p['dist']<=3),
           ("3RD & LONG",lambda p:p['dn']==3 and p['dist']>=7),("3RD & MED",lambda p:p['dn']==3 and 4<=p['dist']<=6),
           ("3RD & SHORT",lambda p:p['dn']==3 and p['dist']<=3),("4TH DOWN",lambda p:p['dn']==4),
           ("RED ZONE",lambda p:p['zone']=='RZ'),("GOAL LINE",lambda p:p['zone']=='GL'),("BACKED UP",lambda p:p['zone']=='BZ')]
    for lbl,fn in sits2:
        sp=[p for p in plays if fn(p)]
        sr2=[p for p in sp if p['rp']=='Run']; spass=[p for p in sp if p['rp']=='Pass']
        sb=[p for p in sp if p.get('motioned') is True]
        tf=top3(sp,'form'); tc=top3(sp,'backfield')
        sit_rows+=f'''<tr>
          <td class="sit-lbl">{lbl}</td>
          <td style="text-align:center;font-family:Barlow Condensed,sans-serif;font-weight:800;font-size:20px;color:#7b241c">{len(sr2)}</td>
          <td style="text-align:center;font-family:Barlow Condensed,sans-serif;font-weight:800;font-size:20px;color:#5dade2">{len(spass)}</td>
          <td style="text-align:center;font-family:Barlow Condensed,sans-serif;font-weight:700;font-size:16px;color:#b388d4">{str(pct(len(sb),len(sp)))+"%"  if sp else "—"}</td>
          <td style="font-family:Share Tech Mono,monospace;font-size:9px;color:#d4a017">{tf[0]["v"]+" ("+str(tf[0]["n"])+")" if tf else "—"}</td>
          <td style="font-family:Share Tech Mono,monospace;font-size:9px;color:#d4a017">{tc[0]["v"]+" ("+str(tc[0]["n"])+")" if tc else "—"}</td>
        </tr>'''

    def con_rows(items,color,max_n):
        if not items: return '<div style="color:rgba(240,237,232,.3);font-size:12px">No data — tag in Hudl to see trends</div>'
        return ''.join(f'<div style="padding:9px 0;border-bottom:1px solid rgba(240,237,232,.1)"><div style="display:flex;justify-content:space-between"><span style="font-family:Share Tech Mono,monospace;font-size:11px;color:{color}">{x["v"]}</span><span style="font-family:Barlow Condensed,sans-serif;font-weight:700;font-size:20px;color:{color}">{x["n"]}</span></div><div style="background:rgba(240,237,232,.06);height:3px;margin-top:4px"><div style="background:{color};height:3px;width:{round(x["n"]/max_n*100)}%"></div></div></div>' for x in items)

    mf=top3(plays,'form')[0]['n'] if top3(plays,'form') else 1
    mc=top3(plays,'backfield')[0]['n'] if top3(plays,'backfield') else 1
    mb=top3(all_motion,'motion')[0]['n'] if top3(all_motion,'motion') else 1

    return f'''<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>DefensiveIQ — {opp}</title>
<link href="https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>:root{{--field:#0a1628;--chalk:#f0ede8;--red:#7b241c;--gold:#d4a017;--blue:#1a5276;--mid:#1e2d3d;--line:rgba(240,237,232,0.1);}}*{{box-sizing:border-box;margin:0;padding:0;}}body{{background:var(--field);color:var(--chalk);font-family:Inter,sans-serif;font-size:17px;line-height:1.5;}}nav{{display:flex;align-items:center;justify-content:space-between;padding:14px 40px;border-bottom:1px solid var(--line);background:rgba(10,22,40,.97);}}.logo{{font-family:Oswald,sans-serif;font-weight:900;font-size:22px;}}.logo span{{color:#7b241c;}}.wrap{{max-width:1200px;margin:0 auto;padding:40px;}}.eyebrow{{font-family:Inter,sans-serif;font-size:19px;letter-spacing:.2em;color:var(--gold);text-transform:uppercase;margin-bottom:10px;}}.rpt-hdr{{display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:32px;padding-bottom:18px;border-bottom:1px solid var(--line);}}.rpt-title{{font-family:Oswald,sans-serif;font-weight:900;font-size:42px;text-transform:uppercase;}}.rpt-meta{{font-family:Inter,sans-serif;font-size:19px;color:var(--gold);text-align:right;}}.sum-grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:36px;}}.sum-card{{background:var(--mid);border:1px solid var(--line);padding:16px;}}.sum-lbl{{font-size:13px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:rgba(240,237,232,.65);margin-bottom:5px;}}.sum-val{{font-family:Oswald,sans-serif;font-weight:800;font-size:37px;line-height:1;}}.stitle{{font-family:Oswald,sans-serif;font-weight:800;font-size:21px;text-transform:uppercase;letter-spacing:.06em;margin-bottom:16px;margin-top:36px;padding-bottom:8px;border-bottom:1px solid var(--line);}}.zone-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;}}.zone-card{{background:var(--mid);border:1px solid var(--line);overflow:hidden;}}.zone-hdr{{padding:10px 14px;display:flex;justify-content:space-between;align-items:center;}}.zone-badge{{font-family:Oswald,sans-serif;font-weight:900;font-size:19px;}}.zone-sub{{font-size:13px;color:rgba(240,237,232,.65);}}.zone-plays{{font-family:Inter,sans-serif;font-size:13px;color:var(--gold);}}.zone-body{{padding:11px 14px;}}.bar-row{{margin-bottom:8px;}}.bar-labels{{display:flex;justify-content:space-between;font-size:13px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;margin-bottom:3px;}}.bar-bg{{background:rgba(240,237,232,.07);height:5px;position:relative;}}.bar-fill{{height:5px;position:absolute;left:0;top:0;}}.zone-tags{{margin-top:8px;border-top:1px solid var(--line);padding-top:8px;}}.tag-lbl{{font-size:12px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:rgba(240,237,232,.82);margin-bottom:3px;}}.ctag{{display:inline-block;background:rgba(240,237,232,.05);border:1px solid rgba(240,237,232,.1);font-family:Inter,sans-serif;font-size:12px;padding:2px 4px;margin:1px 1px 1px 0;color:rgba(240,237,232,.82);}}.ctag.f{{border-color:rgba(123,36,28,.4);color:#e8a095;}}.ctag.c{{border-color:rgba(93,173,226,.35);color:#93d4f0;}}.ctag.b{{border-color:rgba(180,136,212,.4);color:#b388d4;}}.hash-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;}}.hash-card{{background:var(--mid);border:1px solid var(--line);padding:18px;text-align:center;}}.hc-title{{font-family:Oswald,sans-serif;font-weight:800;font-size:15px;letter-spacing:.14em;text-transform:uppercase;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid var(--line);}}.hbig{{font-family:Oswald,sans-serif;font-weight:900;font-size:50px;line-height:1;margin-bottom:2px;}}.hsub{{font-size:13px;color:rgba(240,237,232,.62);margin-bottom:10px;}}.hrp{{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:8px;}}.hrp-item{{background:rgba(240,237,232,.04);padding:7px;}}.hrp-lbl{{font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:rgba(240,237,232,.62);}}.hrp-val{{font-family:Oswald,sans-serif;font-weight:700;font-size:21px;}}.htc{{font-family:Inter,sans-serif;font-size:13px;color:rgba(240,237,232,.62);}}.sit-table{{width:100%;border-collapse:collapse;font-size:15px;}}.sit-table th{{background:var(--field);padding:7px 10px;font-family:Oswald,sans-serif;font-weight:700;font-size:13px;letter-spacing:.1em;text-transform:uppercase;color:rgba(240,237,232,.7);border:1px solid var(--line);text-align:center;}}.sit-table th:first-child{{text-align:left;}}.sit-table td{{border:1px solid var(--line);padding:8px 12px;}}.sit-table tr:nth-child(odd) td{{background:rgba(240,237,232,.02);}}.sit-table tr:nth-child(even) td{{background:var(--mid);}}.sit-lbl{{font-family:Oswald,sans-serif;font-weight:700;font-size:15px;white-space:nowrap;}}.con-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:24px;}}</style>
</head><body>
<nav><div class="logo">DEFENSIVE<span>IQ</span></div><div style="font-family:Share Tech Mono,monospace;font-size:10px;color:rgba(240,237,232,.4)">OFFENSIVE TENDENCY REPORT</div></nav>
<div class="wrap">
<div class="rpt-hdr"><div><div class="eyebrow">// Defensive Coordinator — Offensive Tendency Report</div><div class="rpt-title">{opp} — Offensive Analysis</div></div><div class="rpt-meta">WEEK {week}{("<br>"+date) if date else ""}<br>{total} PLAYS ANALYZED</div></div>
<div class="sum-grid">
  <div class="sum-card"><div class="sum-lbl">Total Plays</div><div class="sum-val">{total}</div></div>
  <div class="sum-card"><div class="sum-lbl">Their Run %</div><div class="sum-val" style="color:#7b241c">{pct(len(runs),total)}%</div></div>
  <div class="sum-card"><div class="sum-lbl">Their Pass %</div><div class="sum-val" style="color:#5dade2">{pct(len(passes),total)}%</div></div>
  <div class="sum-card"><div class="sum-lbl">Motion %</div><div class="sum-val" style="color:#b388d4">{pct(len(all_motion),total)}%</div></div>
  <div class="sum-card"><div class="sum-lbl">RZ Run %</div><div class="sum-val" style="color:#d4a017">{pct(len([p for p in rz if p["rp"]=="Run"]),len(rz))}%</div></div>
</div>
<div class="stitle">Field Zone Breakdown</div><div class="zone-grid">{zone_cards}</div>
<div class="stitle">Hash Tendencies</div><div class="hash-grid">{hash_cards}</div>
<div class="stitle">Situational Summary</div>
<table class="sit-table">
  <tr><th style="text-align:left">Situation</th><th>Runs</th><th>Passes</th><th>Motion %</th><th style="text-align:left">Top Formation</th><th style="text-align:left">Top Backfield</th></tr>
  {sit_rows}
</table>
<div class="stitle">Offensive Tendencies</div>
<div class="con-grid">
  <div><div class="eyebrow" style="margin-bottom:12px">// Top Formations</div>{con_rows(top3(plays,"form"),"#e8a095",mf)}</div>
  <div><div class="eyebrow" style="margin-bottom:12px">// Top Backfields</div>{con_rows(top3(plays,"backfield"),"#93d4f0",mc)}</div>
  <div><div class="eyebrow" style="margin-bottom:12px">// Top Motion Types</div>{con_rows(top3(all_motion,"motion"),"#b388d4",mb)}</div>
</div>
</div></body></html>'''

# ── STREAMLIT UI ──────────────────────────────────────────────
st.markdown('<div class="main-title">Defensive<span style="color:#7b241c">IQ</span></div>', unsafe_allow_html=True)
st.markdown('<div style="font-size:16px;color:rgba(240,237,232,.55);margin-bottom:24px;font-weight:300">Scout your next opponent\'s offense. Upload offensive film and uncover every tendency—formations, personnel, backfield sets, motion, run/pass, and situations—to build your defensive game plan.</div>', unsafe_allow_html=True)

st.divider()
col1,col2,col3=st.columns(3)
with col1: opp  = st.text_input("Opponent Name", placeholder="e.g. Lincoln High School")
with col2: week = st.text_input("Week", placeholder="e.g. 3")
with col3: date = st.text_input("Game Date", placeholder="e.g. Sept 5, 2026")

st.markdown("**Presentation Colors** — used for your defensive scouting PowerPoint")
cc1, cc2 = st.columns(2)
with cc1: team_primary = st.color_picker("Primary (headers/titles)", "#7B241C")
with cc2: team_accent  = st.color_picker("Accent (highlights)", "#C9A227")

st.markdown("---")
uploaded=st.file_uploader("Upload Hudl Playlist Export (.xlsx or .csv)", type=['xlsx','xls','csv'],
                           help="Export the opponent's offensive playlist from Hudl as Excel or CSV and upload here.")

if uploaded and st.button("⚡ RUN ANALYSIS"):
    with st.spinner("Analyzing offensive tendencies..."):
        try:
            if uploaded.name.lower().endswith('.csv'):
                df=pd.read_csv(uploaded)
            else:
                df=pd.read_excel(uploaded)

            # Flexible header mapping — handles differently-named columns
            df, matched, missing = map_columns(df)
            with st.expander("📋 Column mapping — what we found in your file"):
                st.write("**Matched:** " + (", ".join(matched.keys()) if matched else "none"))
                if missing:
                    st.info("Not found (those sections will be blank): " + ", ".join(missing))

            # Hard stop with a CLEAR message if a truly required column is absent
            req_missing = check_required(matched)
            if req_missing:
                st.error("Your file is missing required column(s): " + ", ".join(req_missing))
                st.info("These are needed to analyze anything. Check that your Hudl export includes "
                        "Play Type (Run/Pass), Yard Line, Down, and Distance — then re-export and try again.")
                st.stop()

            plays=load_plays(df)
            if len(plays)==0:
                st.error("No Run/Pass plays found in this file.")
                st.info("Common causes: PLAY TYPE uses different words than 'Run'/'Pass', or YARD LN is blank. "
                        "Open the column mapping above to see what we detected.")
            else:
                opp_name=opp or "Opponent"
                prog=st.progress(0,"Reading data...")
                prog.progress(30,"Calculating zone tendencies...")
                excel_bytes=build_excel(plays,opp_name,week,date)
                prog.progress(70,"Building HTML report...")
                html_bytes=build_html(plays,opp_name,week,date).encode('utf-8')
                prog.progress(85,"Building scouting presentation...")
                pptx_bytes=build_pptx(plays,opp_name,week,date,team_primary,team_accent)
                prog.progress(100,"Complete!")

                runs=[p for p in plays if p['rp']=='Run']; passes=[p for p in plays if p['rp']=='Pass']
                motion_yes=[p for p in plays if p.get('motioned') is True]
                motion_known=[p for p in plays if p.get('motioned') is not None]
                rz=[p for p in plays if p['zone']=='RZ']

                st.success(f"✅ Analysis complete — {len(plays)} plays analyzed")
                st.divider()

                m1,m2,m3,m4,m5=st.columns(5)
                m1.metric("Total Plays",   len(plays))
                m2.metric("Run %",         f"{pct(len(runs),len(plays))}%")
                m3.metric("Pass %",        f"{pct(len(passes),len(plays))}%")
                m4.metric("Motion %",      f"{pct(len(motion_yes),len(motion_known))}%" if motion_known else "—")
                m5.metric("RZ Run %",      f"{pct(len([p for p in rz if p['rp']=='Run']),len(rz))}%")

                st.divider()
                st.markdown("### Download Your Reports")
                d1,d2,d3=st.columns(3)
                fname=(opp_name+"_" if opp_name else "")+(f"Week{week}_" if week else "")+"DefensiveIQ"
                with d1:
                    st.download_button("📊 Excel Workbook",data=excel_bytes,
                        file_name=f"{fname}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with d2:
                    st.download_button("🌐 HTML Report",data=html_bytes,
                        file_name=f"{fname}_Report.html",mime="text/html")
                with d3:
                    st.download_button("📽️ Scouting Presentation",data=pptx_bytes,
                        file_name=f"{fname}_Scouting.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

                st.divider()
                st.markdown("### Quick Summary")
                z1,z2,z3=st.columns(3)
                with z1:
                    st.markdown("**Top Formations Faced**")
                    for x in top3(plays,'form'): st.markdown(f"- {x['v']} ({x['n']} plays)")
                    if not any(p['form'] for p in plays): st.markdown("*Tag OFF FORM in Hudl*")
                with z2:
                    st.markdown("**Top Backfields**")
                    for x in top3(plays,'backfield'): st.markdown(f"- {x['v']} ({x['n']} plays)")
                    if not any(p['backfield'] for p in plays): st.markdown("*Tag BACKFIELD in Hudl*")
                with z3:
                    st.markdown("**Top Motion Types**")
                    for x in top3(motion_yes,'motion'): st.markdown(f"- {x['v']} ({x['n']})")
                    if not motion_yes: st.markdown("*Tag MOTION in Hudl*")

        except Exception as e:
            import traceback
            st.error("Something went wrong reading this file — it may be formatted differently than expected.")
            st.info("Check that this is a Hudl playlist export. If it looks right, send the file and the "
                    "details below to support so it can be fixed.")
            with st.expander("🔧 Technical details"):
                st.code(traceback.format_exc())
            st.info("Make sure this is a Hudl playlist export with PLAY TYPE, YARD LN, OFF FORM, DN, DIST, HASH columns.")

st.divider()
st.markdown('<div style="font-family:Share Tech Mono,monospace;font-size:10px;color:rgba(240,237,232,.25);text-align:center;padding:20px 0">© 2026 DEFENSIVEIQ · BUILT FOR DEFENSIVE COORDINATORS</div>',unsafe_allow_html=True)
