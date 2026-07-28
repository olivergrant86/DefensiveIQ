from PySide6.QtWidgets import QTextEdit


class LogPanel(QTextEdit):
    def __init__(self):
        super().__init__()

        self.setReadOnly(True)
        self.append("DefensiveIQ started.")

    def log(self, message: str):
        self.append(message)