import logging
from datetime import datetime
from pathlib import Path


class LoggingService:
    def __init__(self, log_dir: Path):
        log_dir.mkdir(parents=True, exist_ok=True)
        name = datetime.now().strftime("master_report_center_%Y%m%d.log")
        self.path = log_dir / name
        self.logger = logging.getLogger("MasterReportCenter")
        self.logger.setLevel(logging.INFO)
        if not self.logger.handlers:
            formatter = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s")
            fh = logging.FileHandler(self.path, encoding="utf-8")
            sh = logging.StreamHandler()
            fh.setFormatter(formatter)
            sh.setFormatter(formatter)
            self.logger.addHandler(fh)
            self.logger.addHandler(sh)

    def info(self, message): self.logger.info(message)
    def warning(self, message): self.logger.warning(message)
    def error(self, message): self.logger.error(message)
