import json
import math
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path


class TrialService:
    """Local trial + activated-license state for Master Report Center Publish."""

    TRIAL_DAYS = 12
    REG_PATH = r"Software\GoldenFir\MasterReportCenter"
    FILE_NAME = "trial_state.json"
    LICENSE_FILE_NAME = "license_state.json"

    def __init__(self):
        local = os.environ.get("LOCALAPPDATA")
        if local:
            self.state_dir = Path(local) / "GoldenFir" / "MasterReportCenter"
        else:
            self.state_dir = Path.home() / ".goldenfir" / "MasterReportCenter"
        self.state_file = self.state_dir / self.FILE_NAME
        self.license_file = self.state_dir / self.LICENSE_FILE_NAME

    @staticmethod
    def _utc_now():
        return datetime.now(timezone.utc)

    @staticmethod
    def _parse(value):
        if not value:
            return None
        try:
            dt = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc)
        except Exception:
            return None

    @staticmethod
    def _iso(dt):
        return dt.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")

    def _read_json(self, path):
        try:
            if not path.exists():
                return {}
            data = json.loads(path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def _read_file(self):
        return self._read_json(self.state_file)

    def _write_file(self, first_run, last_seen):
        try:
            self.state_dir.mkdir(parents=True, exist_ok=True)
            payload = {
                "first_run_utc": self._iso(first_run),
                "last_seen_utc": self._iso(last_seen),
                "trial_days": self.TRIAL_DAYS,
            }
            self.state_file.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _read_registry(self):
        if os.name != "nt":
            return {}
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.REG_PATH) as key:
                first, _ = winreg.QueryValueEx(key, "TrialFirstRunUtc")
                last, _ = winreg.QueryValueEx(key, "TrialLastSeenUtc")
                return {"first_run_utc": first, "last_seen_utc": last}
        except Exception:
            return {}

    def _write_registry(self, first_run, last_seen):
        if os.name != "nt":
            return
        try:
            import winreg
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.REG_PATH) as key:
                winreg.SetValueEx(key, "TrialFirstRunUtc", 0, winreg.REG_SZ, self._iso(first_run))
                winreg.SetValueEx(key, "TrialLastSeenUtc", 0, winreg.REG_SZ, self._iso(last_seen))
        except Exception:
            pass

    def _read_license_state(self):
        data = self._read_json(self.license_file)
        if str(data.get("status", "")).lower() not in {"active", "refunded", "revoked"}:
            return None
        return data

    def save_license_state(self, state: dict):
        payload = dict(state or {})
        payload["status"] = str(payload.get("status", "active")).lower()
        payload["saved_at_utc"] = self._iso(self._utc_now())
        try:
            self.state_dir.mkdir(parents=True, exist_ok=True)
            self.license_file.write_text(
                json.dumps(payload, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
        except Exception:
            pass

        if os.name == "nt":
            try:
                import winreg
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.REG_PATH) as key:
                    winreg.SetValueEx(
                        key,
                        "LicenseStateJson",
                        0,
                        winreg.REG_SZ,
                        json.dumps(payload, ensure_ascii=False),
                    )
            except Exception:
                pass

    def save_active_license(self, state: dict):
        payload = dict(state or {})
        payload["status"] = "active"
        self.save_license_state(payload)

    def _read_license_state_registry(self):
        if os.name != "nt":
            return None
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.REG_PATH) as key:
                raw, _ = winreg.QueryValueEx(key, "LicenseStateJson")
            data = json.loads(raw)
            if isinstance(data, dict) and str(data.get("status", "")).lower() in {"active", "refunded", "revoked"}:
                return data
        except Exception:
            return None
        return None

    def get_state(self):
        license_state = self._read_license_state() or self._read_license_state_registry()
        if license_state:
            license_state.setdefault("days_remaining", 0)
            return license_state

        now = self._utc_now()
        file_state = self._read_file()
        reg_state = self._read_registry()

        first_candidates = [
            self._parse(file_state.get("first_run_utc")),
            self._parse(reg_state.get("first_run_utc")),
        ]
        first_candidates = [x for x in first_candidates if x is not None]
        first_run = min(first_candidates) if first_candidates else now

        last_candidates = [
            self._parse(file_state.get("last_seen_utc")),
            self._parse(reg_state.get("last_seen_utc")),
        ]
        last_candidates = [x for x in last_candidates if x is not None]
        last_seen = max(last_candidates) if last_candidates else first_run

        effective_now = max(now, last_seen)
        expiry = first_run + timedelta(days=self.TRIAL_DAYS)
        seconds_left = (expiry - effective_now).total_seconds()
        days_remaining = max(0, int(math.ceil(seconds_left / 86400.0)))
        status = "trial" if seconds_left > 0 else "expired"
        clock_rollback_detected = now < (last_seen - timedelta(minutes=5))

        stored_last_seen = max(last_seen, now)
        self._write_file(first_run, stored_last_seen)
        self._write_registry(first_run, stored_last_seen)

        return {
            "status": status,
            "days_remaining": days_remaining,
            "license_type": None,
            "trial_days": self.TRIAL_DAYS,
            "first_run_utc": self._iso(first_run),
            "expires_utc": self._iso(expiry),
            "clock_rollback_detected": clock_rollback_detected,
        }
