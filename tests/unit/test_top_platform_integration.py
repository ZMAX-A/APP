from __future__ import annotations

import json
from pathlib import Path

import pytest

from yanjia_automation.config import Settings
from yanjia_automation.excel.variables import VariableResolutionError, VariableResolver

ROOT = Path(__file__).resolve().parents[2]


def _settings(project_root: Path) -> Settings:
    return Settings(
        appium_server_url="http://127.0.0.1:4723",
        app_package="com.xiaofutech.yanjia_ai",
        app_activity=".activity.SplashActivity",
        udid=None,
        store_name=None,
        no_reset=True,
        run_seeded=False,
        allow_mutation=False,
        allow_destructive=False,
        capture_sensitive_artifacts=False,
        skip_device_initialization=False,
        skip_server_installation=False,
        project_root=project_root,
    )


def test_testops_package_manifest_matches_android_runner() -> None:
    manifest = json.loads((ROOT / "testops-package.json").read_text(encoding="utf-8"))

    assert manifest == {
        "schema_version": "1.0",
        "name": "yanjia-android-appium",
        "version": "1.0.1",
        "runner_type": "ANDROID_APPIUM",
        "entrypoint": "scripts/run-excel.ps1",
        "workbook": "test_case.xlsx",
        "target_type": "APP",
    }


def test_testops_execution_can_ignore_local_dotenv(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    (tmp_path / ".env").write_text("SEEDED_CUSTOMER_PHONE=private-local-value\n", encoding="utf-8")
    monkeypatch.setenv("YANJIA_IGNORE_DOTENV", "true")

    resolver = VariableResolver.from_settings(
        _settings(tmp_path),
        run_id="top-run",
        require_credentials=False,
    )

    with pytest.raises(VariableResolutionError, match="SEEDED_CUSTOMER_PHONE"):
        resolver.resolve("${SEEDED_CUSTOMER_PHONE}")


def test_run_excel_accepts_testops_output_isolation_parameters() -> None:
    script = (ROOT / "scripts" / "run-excel.ps1").read_text(encoding="utf-8-sig")

    for parameter in ("[string]$ReportsRoot", "[string]$RunId", "[switch]$IgnoreDotEnv"):
        assert parameter in script
    assert "$env:YANJIA_IGNORE_DOTENV = 'true'" in script
    assert 'allure-results\\$effectiveRunId' in script
