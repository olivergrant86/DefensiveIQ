from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QDockWidget,
    QToolBar,
    QStatusBar,
)
from PySide6.QtCore import Qt

from gui.dashboard import Dashboard
from gui.log_panel import LogPanel


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("DefensiveIQ")
        self.resize(1400, 850)

        self.build_toolbar()
        self.build_statusbar()
        self.build_dashboard()
        self.build_log_panel()

    def build_toolbar(self):

        toolbar = QToolBar("Main Toolbar")

        self.addToolBar(toolbar)

    def build_statusbar(self):

        status = QStatusBar()

        status.showMessage("Ready")

        self.setStatusBar(status)

    def build_dashboard(self):

        central = QWidget()

        layout = QVBoxLayout()

        dashboard = Dashboard()

        layout.addWidget(dashboard)

        central.setLayout(layout)

        self.setCentralWidget(central)

    def build_log_panel(self):

        dock = QDockWidget("Application Log")

        dock.setAllowedAreas(Qt.RightDockWidgetArea)

        dock.setWidget(LogPanel())

        self.addDockWidget(Qt.RightDockWidgetArea, dock)
        self.import_action = QAction("Open Playlist", self)

self.toolbar.addAction(self.import_action)