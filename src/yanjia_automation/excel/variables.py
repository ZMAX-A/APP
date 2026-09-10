from __future__ import annotations

import json
import os
import re
from datetime import date, datetime
from hashlib import sha256

from dotenv import dotenv_values

from yanjia_automation.config import Settings

VARIABLE_PATTERN = re.compile(r"\$\{([A-Za-z_][A-Za-z0-9_]*)\}")
SENSITIVE_KEY_PARTS = ("PASSWORD", "USERNAME", "PHONE", "CUSTOMER", "TOKEN", "SECRET")


class VariableResolutionError(ValueError):
    """Raised when an Excel value references an undefined variable."""


class VariableResolver:
    def __init__(self, values: dict[str, str], *, sensitive_values: set[str]) -> None:
        self._values = values
        # Tracebacks render locals with repr(), while JSON may escape quotes,
        # backslashes and non-ASCII characters. Redact those representations too.
        self._sensitive_values = {
            representation
            for value in sensitive_values
            if value
            for representation in (
                value,
                repr(value)[1:-1],
                json.dumps(value, ensure_ascii=False)[1:-1],
                json.dumps(value, ensure_ascii=True)[1:-1],
            )
        }

    @classmethod
    def from_settings(
        cls,
        settings: Settings,
        *,
        run_id: str,
        require_credentials: bool = True,
    ) -> VariableResolver:
        values: dict[str, str] = {
            key: str(value)
            for key, value in dotenv_values(settings.project_root / ".env").items()
            if value is not None
        }
        values.update({key: value for key, value in os.environ.items() if value is not None})
        run_token = re.sub(r"[^A-Za-z0-9]", "", run_id)[-6:] or "RUN"
        run_phone_suffix = int.from_bytes(
            sha256(run_id.encode("utf-8")).digest()[:4], "big"
        ) % 1_000_000
        values.update(
            {
                "RUN_ID": run_id,
                "RUN_TOKEN": run_token,
                "RUN_PHONE": f"13900{run_phone_suffix:06d}",
                "TODAY": date.today().isoformat(),
                "NOW": datetime.now().isoformat(timespec="seconds"),
                "YANJIA_STORE_OR_FIRST": settings.store_name or "",
                "YANJIA_APP_PACKAGE": settings.app_package,
            }
        )
        values.setdefault("INVALID_USERNAME", "__invalid_yanjia_user__")
        values.setdefault("INVALID_PASSWORD", "__invalid_yanjia_password__")
        values.setdefault(
            "NON_EXISTENT_CASE_TAG",
            f"__missing_case_tag_{run_token}__",
        )
        values.setdefault("SPACE", " ")

        try:
            credentials = settings.credentials()
        except RuntimeError:
            if require_credentials:
                raise
        else:
            credential_aliases = {
                "YANJIA_USERNAME": credentials.username,
                "YANJIA_PASSWORD": credentials.password,
                "TEST_USERNAME": credentials.username,
                "TEST_PASSWORD": credentials.password,
            }
            for name, value in credential_aliases.items():
                # The example .env deliberately includes blank TEST_* keys.
                # Treat blanks as unset so migrated cases inherit YANJIA_*.
                if not values.get(name, "").strip():
                    values[name] = value

        sensitive_values = {
            value
            for key, value in values.items()
            if any(part in key.upper() for part in SENSITIVE_KEY_PARTS)
        }
        return cls(values, sensitive_values=sensitive_values)

    def resolve(self, value: str | None) -> str | None:
        if value is None:
            return None
        if value.strip().upper() in {"<EMPTY>", "<NULL>", "<NONE>"}:
            return ""

        missing: set[str] = set()

        def replace(match: re.Match[str]) -> str:
            name = match.group(1)
            resolved = self._values.get(name)
            if resolved is None:
                missing.add(name)
                return match.group(0)
            return resolved

        resolved = VARIABLE_PATTERN.sub(replace, value)
        if missing:
            names = ", ".join(sorted(missing))
            raise VariableResolutionError(f"Excel引用了未配置的变量：{names}")
        return resolved

    def redact(self, value: object) -> str:
        text = str(value)
        for secret in sorted(self._sensitive_values, key=len, reverse=True):
            text = text.replace(secret, "***")
        return text

    @property
    def available_names(self) -> frozenset[str]:
        """Return configured variable names without exposing their values."""
        return frozenset(self._values)
