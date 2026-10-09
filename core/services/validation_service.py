from pathlib import Path

class ValidationService:
    @staticmethod
    def require_file(path, label="File"):
        if not path or not Path(path).is_file():
            raise ValueError(f"{label} does not exist: {path}")

    @staticmethod
    def require_folder(path, label="Folder"):
        if not path or not Path(path).is_dir():
            raise ValueError(f"{label} does not exist: {path}")
