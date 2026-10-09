from __future__ import annotations

import json
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path
from urllib import error, request


DEFAULT_API_URL = "https://mrc-licensing-api-production.samuelshanjin.workers.dev"
INSTALL_ID_FILE = "install_id.txt"


class BackendClientError(RuntimeError):
    """Raised when the licensing backend cannot complete a request."""


class LicensingBackendClient:
    def __init__(self, base_url: str | None = None, timeout_seconds: float = 10.0):
        self.base_url = (base_url or os.getenv("MRC_LICENSING_API_URL") or DEFAULT_API_URL).rstrip("/")
        self.timeout_seconds = timeout_seconds

    @staticmethod
    def get_or_create_install_id() -> str:
        local = os.environ.get("LOCALAPPDATA")
        if local:
            state_dir = Path(local) / "GoldenFir" / "MasterReportCenter"
        else:
            state_dir = Path.home() / ".goldenfir" / "MasterReportCenter"
        path = state_dir / INSTALL_ID_FILE

        try:
            if path.exists():
                value = path.read_text(encoding="utf-8").strip()
                if value:
                    return value[:128]
        except Exception:
            pass

        value = f"MRCI-{uuid.uuid4().hex.upper()}"
        try:
            state_dir.mkdir(parents=True, exist_ok=True)
            path.write_text(value, encoding="utf-8")
        except Exception:
            pass
        return value

    def _request_json(self, method: str, path: str, payload: dict | None = None, extra_headers: dict | None = None) -> dict:
        data = None if method == "GET" else json.dumps(payload or {}).encode("utf-8")
        headers = {
            "Accept": "application/json",
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/140.0.0.0 Safari/537.36 "
                "GoldenFir-MRC/2.0.2"
            ),
            "Accept-Language": "en-AU,en;q=0.9",
        }
        if method != "GET":
            headers["Content-Type"] = "application/json"
        if extra_headers:
            headers.update(extra_headers)
        req = request.Request(
            f"{self.base_url}{path}",
            data=data,
            headers=headers,
            method=method,
        )
        try:
            with request.urlopen(req, timeout=self.timeout_seconds) as response:
                raw = response.read().decode("utf-8")
                return json.loads(raw) if raw else {}
        except error.HTTPError as exc:
            detail = ""
            try:
                body = exc.read().decode("utf-8")
                parsed = json.loads(body)
                detail = parsed.get("detail") or body
                if isinstance(detail, list):
                    detail = "; ".join(
                        str(item.get("msg", item)) if isinstance(item, dict) else str(item)
                        for item in detail
                    )
            except Exception:
                detail = str(exc.reason or exc)
            raise BackendClientError(
                f"Licensing service rejected the request ({exc.code}). {detail}".strip()
            ) from exc
        except (error.URLError, TimeoutError, OSError) as exc:
            reason = getattr(exc, "reason", exc)
            raise BackendClientError(
                f"Could not reach the Golden Fir licensing service at {self.base_url}. {reason}"
            ) from exc
        except (json.JSONDecodeError, ValueError) as exc:
            raise BackendClientError("The licensing service returned an invalid response.") from exc

    def _post_json(self, path: str, payload: dict | None = None, extra_headers: dict | None = None) -> dict:
        return self._request_json("POST", path, payload, extra_headers)

    def _get_json(self, path: str, extra_headers: dict | None = None) -> dict:
        return self._request_json("GET", path, None, extra_headers)

    def create_purchase_intent(
        self,
        *,
        email: str,
        product_code: str,
        terms_version: str,
        app_version: str | None,
    ) -> dict:
        payload = {
            "email": email.strip().lower(),
            "product_code": product_code,
            "terms_version": terms_version,
            "terms_accepted_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
            "app_version": app_version,
            "install_id": self.get_or_create_install_id(),
        }
        result = self._post_json("/purchase-intents", payload)
        intent_id = str(result.get("purchase_intent_id", "")).strip()
        if not intent_id:
            raise BackendClientError("The licensing service did not return a Purchase Intent ID.")
        return result

    def create_paypal_order(self, purchase_intent_id: str) -> dict:
        result = self._post_json(
            f"/purchase-intents/{purchase_intent_id}/paypal-order",
            {},
        )
        order_id = str(result.get("paypal_order_id", "")).strip()
        approval_url = str(result.get("approval_url", "")).strip()
        if not order_id or not approval_url:
            raise BackendClientError(
                "The licensing service did not return a complete PayPal order."
            )
        return result

    def get_purchase_status(self, purchase_intent_id: str) -> dict:
        install_id = self.get_or_create_install_id()
        return self._get_json(
            f"/purchase-intents/{purchase_intent_id}/status",
            {"X-MRC-Install-ID": install_id},
        )

    def activate_paid_purchase(self, purchase_intent_id: str) -> dict:
        install_id = self.get_or_create_install_id()
        result = self._post_json(
            f"/purchase-intents/{purchase_intent_id}/activate",
            {"install_id": install_id},
        )
        if str(result.get("status", "")).upper() != "ACTIVE":
            raise BackendClientError("The licensing service did not activate the license.")
        return result


    def recover_existing_license(self, *, email: str, recovery_code: str, mode: str = "recover") -> dict:
        install_id = self.get_or_create_install_id()
        result = self._post_json(
            "/licenses/recover",
            {
                "email": email.strip().lower(),
                "recovery_code": recovery_code.strip().upper(),
                "install_id": install_id,
                "mode": "activate" if mode == "activate" else "recover",
            },
        )
        if str(result.get("status", "")).upper() != "ACTIVE":
            raise BackendClientError("The licensing service did not activate the license.")
        return result

    def submit_refund_request(
        self,
        *,
        email: str,
        recovery_code: str | None,
        no_recovery_code: bool,
        purchase_reference: str,
        reason: str,
        details: str,
    ) -> dict:
        return self._post_json(
            "/refund-requests",
            {
                "email": email.strip().lower(),
                "recovery_code": (recovery_code or "").strip().upper() or None,
                "no_recovery_code": bool(no_recovery_code),
                "purchase_reference": purchase_reference.strip(),
                "reason": reason.strip(),
                "details": details.strip(),
                "install_id": self.get_or_create_install_id(),
            },
        )

    def get_license_status(self, license_id: str) -> dict:
        install_id = self.get_or_create_install_id()
        return self._get_json(
            f"/licenses/{license_id}/status",
            {"X-MRC-Install-ID": install_id},
        )

