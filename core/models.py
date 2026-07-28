from dataclasses import dataclass
from typing import Optional


@dataclass
class Play:
    play_number: int = 0

    quarter: Optional[int] = None
    clock: str = ""

    down: Optional[int] = None
    distance: Optional[int] = None

    hash_mark: str = ""
    field_position: Optional[int] = None

    offense_personnel: str = ""
    defense_personnel: str = ""

    formation: str = ""
    strength: str = ""

    motion: str = ""
    shift: str = ""

    play_type: str = ""
    run_pass: str = ""

    concept: str = ""

    gain: int = 0

    result: str = ""

    explosive: bool = False
    success: bool = False

    red_zone: bool = False
    goal_line: bool = False

    notes: str = ""
    @dataclass
class AnalyticsSummary:
    total_plays: int = 0
    runs: int = 0
    passes: int = 0

    run_pct: float = 0.0
    pass_pct: float = 0.0

    avg_gain: float = 0.0

    explosive_plays: int = 0

    success_rate: float = 0.0