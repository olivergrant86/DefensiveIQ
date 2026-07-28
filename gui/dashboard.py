from PySide6.QtWidgets import (
    QWidget,
    QLabel,
    QGridLayout,
    QGroupBox,
    QVBoxLayout,
)
from PySide6.QtCore import Qt


class Dashboard(QWidget):
    def __init__(self):
        super().__init__()

        self.total_label = QLabel("0")
        self.runs_label = QLabel("0")
        self.passes_label = QLabel("0")
        self.run_pct_label = QLabel("0%")
        self.pass_pct_label = QLabel("0%")
        self.avg_gain_label = QLabel("0.0")

        for label in (
            self.total_label,
            self.runs_label,
            self.passes_label,
            self.run_pct_label,
            self.pass_pct_label,
            self.avg_gain_label,
        ):
            label.setAlignment(Qt.AlignCenter)
            label.setStyleSheet(
                "font-size: 28px; font-weight: bold;"
            )

        layout = QGridLayout()
        layout.addWidget(self._card("Total Plays", self.total_label), 0, 0)
        layout.addWidget(self._card("Runs", self.runs_label), 0, 1)
        layout.addWidget(self._card("Passes", self.passes_label), 0, 2)
        layout.addWidget(self._card("Run %", self.run_pct_label), 1, 0)
        layout.addWidget(self._card("Pass %", self.pass_pct_label), 1, 1)
        layout.addWidget(self._card("Avg Gain", self.avg_gain_label), 1, 2)

        self.setLayout(layout)

    def _card(self, title: str, value: QLabel):
        box = QGroupBox(title)

        layout = QVBoxLayout()
        layout.addStretch()
        layout.addWidget(value)
        layout.addStretch()

        box.setLayout(layout)

        return box

    def update_summary(self, summary):
        self.total_label.setText(str(summary.total_plays))
        self.runs_label.setText(str(summary.runs))
        self.passes_label.setText(str(summary.passes))
        self.run_pct_label.setText(f"{summary.run_pct:.1f}%")
        self.pass_pct_label.setText(f"{summary.pass_pct:.1f}%")
        self.avg_gain_label.setText(f"{summary.avg_gain:.1f}")