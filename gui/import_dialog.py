from PySide6.QtWidgets import QFileDialog


class ImportDialog:

    @staticmethod
    def get_filename(parent=None):

        return QFileDialog.getOpenFileName(

            parent,

            "Open Playlist",

            "",

            "Hudl Playlist (*.xlsx *.csv)"
        )