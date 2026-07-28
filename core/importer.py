from pathlib import Path
from typing import List

from core.models import Play


class PlaylistImporter:

    def import_playlist(self, filename: str) -> List[Play]:
        raise NotImplementedError