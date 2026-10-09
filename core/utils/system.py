import os
from pathlib import Path

def open_folder(path):
    p = Path(path)
    if hasattr(os, "startfile"):
        os.startfile(str(p))
