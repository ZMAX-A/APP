from __future__ import annotations

import os
import re
from datetime import date, datetime

from dotenv import dotenv_values

from yanjia_automation.config import Settings

VARIABLE_PATTERN = re.compile(r"\$\{([A-Za-z_][A-Za-z0-9_]*)\}")
SENSITIVE_KEY_PARTS = ("PASSWORD", "USERNAME", "PHONE", "CUSTOMER", "TOKEN", "SECRET")


class VariableResolutionError(ValueError):
    """Raised when an Excel value references an undefined variable."""


class VariableResolver:
    def __init__(self, values: dict[str, str], *, sensitive_values: set[str]) -> None:
        self._values = values
        self._sensitive_values = {value for value in sensitive_values if value}

    @classmethod
    def from_settings(cls, settings: Settings, *, run_id: str) -> VariableResolver:
        values: dict[str, str] = {
            key: str(value)
            for key, value in dotenv_values(settings.project_root / ".env").items()
            if value is not None
        }
        values.update({key: value for key, value in os.environ.items() if value is not None})
        values.update(
            {
                "RUN_ID": run_id,
                "TODAY": date.today().isoformat(),
                "NOW": datetime.now().isoformat(timespec="seconds"),
                "YANJIA_STORE_OR_FIRST": settings.store_name or "",
                "YANJIA_APP_PACKAGE": settings.app_package,
            }
        )

        credentials = settings.credentials()
        values.setdefault("YANJIA_USERNAME", credentials.username)
        values.setdefault("YANJIA_PASSWORD", credentials.password)

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
