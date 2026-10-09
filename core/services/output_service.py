from pathlib import Path
from datetime import datetime
import re

class OutputService:
    @staticmethod
    def safe_name(text):
        value = re.sub(r'[<>:"/\|?*]+', '_', str(text)).strip()
        return value or "Blank"

    @staticmethod
    def timestamped_path(folder, stem, suffix=".xlsx"):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return str(Path(folder) / f"{stem}_{ts}{suffix}")
