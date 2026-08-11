from __future__ import annotations

import json
from collections.abc import Generator, Iterator
from pathlib import Path
from typing import Any

import pytest
from appium.webdriver.webdriver import WebDriver

from yanjia_automation.config import Settings, load_settings
from yanjia_automation.driver import DriverManager
from yanjia_automation.flows.navigation import ensure_home
from yanjia_automation.screens.home import HomeScreen


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    settings = load_settings()
    run_seeded = bool(config.getoption("--run-seeded")) or settings.run_seeded
    allow_mutation = bool(config.getoption("--allow-mutation")) or settings.allow_mutation
    allow_destructive = bool(config.getoption("--allow-destructive")) or settings.allow_destructive

    skip_seeded = pytest.mark.skip(reason="requires --run-seeded or YANJIA_RUN_SEEDED=true")
    skip_mutation = pytest.mark.skip(
        reason="requires --allow-mutation or YANJIA_ALLOW_MUTATION=true"
    )
    skip_destructive = pytest.mark.skip(
        reason=(
            "requires both mutation and destructive authorization; destructive tests are "
            "disabled by default"
        )
    )

    for item in items:
        if item.get_closest_marker("excel_driven"):
            # Excel-driven tests record safety skips themselves so SKIP is written back.
            continue
        if item.get_closest_marker("requires_seed") and not run_seeded:
            item.add_marker(skip_seeded)
        if item.get_closest_marker("mutating") and not allow_mutation:
            item.add_marker(skip_mutation)
        if item.get_closest_marker("destructive") and not (
            allow_mutation and allow_destructive
        ):
            item.add_marker(skip_destructive)


@pytest.fixture(scope="session")
def settings() -> Settings:
    return load_settings()


@pytest.fixture(scope="session")
def driver_manager(settings: Settings) -> Iterator[DriverManager]:
    manager = DriverManager(settings)
    try:
        yield manager
    finally:
        manager.quit()


@pytest.fixture()
def driver(driver_manager: DriverManager) -> WebDriver:
    return driver_manager.get()


@pytest.fixture()
def home(driver: WebDriver, settings: Settings) -> HomeScreen:
    return ensure_home(driver, settings)


@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(
    item: pytest.Item,
    call: pytest.CallInfo[Any],
) -> Generator[None, Any, None]:
    outcome = yield
    report = outcome.get_result()
    setattr(item, f"report_{report.when}", report)


@pytest.fixture(autouse=True)
def failure_evidence(
    request: pytest.FixtureRequest,
    settings: Settings,
) -> Iterator[None]:
    yield

    report = getattr(request.node, "report_call", None)
    if not report or not report.failed or "driver" not in request.fixturenames:
        return

    active_driver: WebDriver = request.getfixturevalue("driver")
    case_dir = _artifact_directory(settings.project_root, request.node.nodeid)
    metadata = {
        "nodeid": request.node.nodeid,
        "activity": active_driver.current_activity,
        "package": settings.app_package,
        "sensitive_artifacts_enabled": settings.capture_sensitive_artifacts,
    }
    (case_dir / "metadata.json").write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    if settings.capture_sensitive_artifacts:
        active_driver.save_screenshot(str(case_dir / "screenshot.png"))
        (case_dir / "page_source.xml").write_text(active_driver.page_source, encoding="utf-8")


def _artifact_directory(project_root: Path, nodeid: str) -> Path:
    safe_name = "".join(character if character.isalnum() else "_" for character in nodeid)
    path = project_root / "artifacts" / safe_name[:180]
    path.mkdir(parents=True, exist_ok=True)
    return path
