from pathlib import Path

class FileEngine:
    @staticmethod
    def excel_files(folder):
        p = Path(folder)
        return sorted([x for x in p.glob("*.xlsx") if not x.name.startswith("~$")])
