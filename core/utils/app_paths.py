from __future__ import annotations
import os
from pathlib import Path
APP_DATA_VENDOR = "GoldenFir"
APP_DATA_PRODUCT = "MasterReportCenter"
def local_app_data_root() -> Path:
    local = os.environ.get("LOCALAPPDATA")
    if local:
        return Path(local) / APP_DATA_VENDOR / APP_DATA_PRODUCT
    return Path.home() / ".goldenfir" / APP_DATA_PRODUCT
def logs_dir() -> Path:
    return local_app_data_root() / "logs"
def user_config_dir() -> Path:
    return local_app_data_root() / "config"
def default_documents_dir() -> Path:
    docs = Path.home() / "Documents"
    return docs if docs.exists() else Path.home()
