import json
from pathlib import Path

class ConfigService:
    def __init__(self, config_dir: Path):
        self.config_dir = config_dir
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.app_config = self._load("app.json")
        self.defaults = self._load("defaults.json")

    def _load(self, name):
        path = self.config_dir / name
        if not path.exists(): return {}
        return json.loads(path.read_text(encoding="utf-8"))

    def get(self, key, default=None):
        return self.app_config.get(key, default)
