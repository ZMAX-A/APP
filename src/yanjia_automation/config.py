from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import dotenv_values

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _as_bool(value: str | None, *, default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True, repr=False)
class Credentials:
    username: str
    password: str


@dataclass(frozen=True)
class Settings:
    appium_server_url: str
    app_package: str
    app_activity: str
    udid: str | None
    store_name: str | None
    no_reset: bool
    run_seeded: bool
    allow_mutation: bool
    allow_destructive: bool
    capture_sensitive_artifacts: bool
    skip_device_initialization: bool
    skip_server_installation: bool
    project_root: Path = PROJECT_ROOT

    def credentials(self) -> Credentials:
        username = _setting("YANJIA_USERNAME") or _setting("TEST_USERNAME")
        password = _setting("YANJIA_PASSWORD") or _setting("TEST_PASSWORD")
        if username and password:
            return Credentials(username=username, password=password)

        legacy_path = self.project_root / "env.txt"
        if legacy_path.exists():
            values = [
                line.strip()
                for line in legacy_path.read_text(encoding="utf-8-sig").splitlines()
                if line.strip() and not line.lstrip().startswith("#")
            ]
            if len(values) >= 2:
                username = _legacy_value(
                    values[0],
                    aliases={"账号", "帐号", "用户名", "username", "user", "account"},
                )
                password = _legacy_value(
                    values[1],
                    aliases={"密码", "password", "pwd"},
                )
                return Credentials(username=username, password=password)

        raise RuntimeError(
            "Missing test credentials. Set YANJIA_USERNAME and YANJIA_PASSWORD "
            "in .env, or keep the two-line legacy env.txt file locally."
        )


_DOTENV = dotenv_values(PROJECT_ROOT / ".env")


def _legacy_value(line: str, *, aliases: set[str]) -> str:
    for separator in ("：", ":", "="):
        if separator not in line:
            continue
        label, value = line.split(separator, 1)
        if label.strip().lower() in aliases:
            stripped = value.strip()
            if not stripped:
                raise RuntimeError("env.txt contains an empty credential value.")
            return stripped
    return line.strip()


def _setting(name: str, default: str | None = None) -> str | None:
    value = os.getenv(name)
    if value is None:
        dotenv_value = _DOTENV.get(name)
        value = str(dotenv_value) if dotenv_value is not None else None
    if value is None:
        return default
    stripped = value.strip()
    return stripped or default


def load_settings() -> Settings:
    return Settings(
        appium_server_url=_setting("APPIUM_SERVER_URL", "http://127.0.0.1:4723") or "",
        app_package=_setting("YANJIA_APP_PACKAGE", "com.xiaofutech.yanjia_ai") or "",
        app_activity=_setting("YANJIA_APP_ACTIVITY", ".activity.SplashActivity") or "",
        udid=_setting("ANDROID_UDID"),
        store_name=_setting("YANJIA_STORE_NAME"),
        no_reset=_as_bool(_setting("YANJIA_NO_RESET"), default=True),
        run_seeded=_as_bool(_setting("YANJIA_RUN_SEEDED")),
        allow_mutation=_as_bool(_setting("YANJIA_ALLOW_MUTATION")),
        allow_destructive=_as_bool(_setting("YANJIA_ALLOW_DESTRUCTIVE")),
        capture_sensitive_artifacts=_as_bool(
            _setting("YANJIA_CAPTURE_SENSITIVE_ARTIFACTS")
        ),
        skip_device_initialization=_as_bool(
            _setting("APPIUM_SKIP_DEVICE_INITIALIZATION")
        ),
        skip_server_installation=_as_bool(
            _setting("APPIUM_SKIP_SERVER_INSTALLATION")
        ),
    )
