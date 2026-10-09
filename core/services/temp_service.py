from pathlib import Path
import tempfile

class TempService:
    @staticmethod
    def create_dir(prefix="mrc_"):
        return Path(tempfile.mkdtemp(prefix=prefix))
