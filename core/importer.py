from pathlib import Path
from typing import List

from core.models import Play


class PlaylistImporter:

    SUPPORTED_TYPES = [".xlsx", ".csv"]

    def import_playlist(self, filename: str) -> List[Play]:

        path = Path(filename)

        if not path.exists():
            raise FileNotFoundError(filename)

        if path.suffix.lower() not in self.SUPPORTED_TYPES:
            raise ValueError(
                f"Unsupported file type: {path.suffix}"
            )

        if path.suffix.lower() == ".csv":
            return self._import_csv(path)

        return self._import_excel(path)

    def _import_csv(self, path: Path) -> List[Play]:
        raise NotImplementedError

    def _import_excel(self, path: Path) -> List[Play]:
        raise NotImplementedError