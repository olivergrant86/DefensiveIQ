import sys
from PySide6.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout


class DefensiveIQ(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("DefensiveIQ")
        self.resize(500, 300)

        layout = QVBoxLayout()

        title = QLabel("DefensiveIQ")
        layout.addWidget(title)

        button = QPushButton("Load Hudl Playlist")
        layout.addWidget(button)

        self.setLayout(layout)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = DefensiveIQ()
    window.show()

    sys.exit(app.exec())