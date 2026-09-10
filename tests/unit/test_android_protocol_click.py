from __future__ import annotations

from typing import cast

import pytest
from appium.webdriver.webdriver import WebDriver

from yanjia_automation.config import Settings
from yanjia_automation.excel.models import ExcelCase, ExcelStep
from yanjia_automation.excel.runner import ExcelCaseRunner, _parse_relative_point
from yanjia_automation.excel.variables import VariableResolver


class _Element:
    rect = {"x": 1135, "y": 1973, "width": 529, "height": 49}

    def __init__(self) -> None:
        self.clicked = False

    def is_displayed(self) -> bool:
        return True

    def is_enabled(self) -> bool:
        return True

    def click(self) -> None:
        self.clicked = True


class _Driver:
    def __init__(self) -> None:
        self.scripts: list[tuple[str, dict[str, int]]] = []

    def execute_script(self, script: str, args: dict[str, int]) -> None:
        self.scripts.append((script, args))


def _settings() -> Settings:
    return Settings(
        appium_server_url="http://127.0.0.1:4723",
        app_package="com.xiaofutech.yanjia_ai",
        app_activity=".activity.SplashActivity",
        udid="device",
        store_name=None,
        no_reset=True,
        run_seeded=False,
        allow_mutation=False,
        allow_destructive=False,
        capture_sensitive_artifacts=False,
        skip_device_initialization=True,
        skip_server_installation=True,
    )


def _step(input_value: str | None) -> ExcelStep:
    return ExcelStep(
        source_row=1,
        case_id="TC-LOGIN-006",
        order=2,
        name="点击用户协议文字链接",
        action="click",
        locator="id=com.xiaofutech.yanjia_ai:id/login_protocol_tv",
        input_value=input_value,
        assertion=None,
        expected=None,
        timeout=1,
        element_index=0,
        continue_on_failure=False,
        enabled=True,
        note=None,
    )


def _runner(driver: _Driver) -> ExcelCaseRunner:
    return ExcelCaseRunner(
        cast(WebDriver, driver),
        _settings(),
        VariableResolver({}, sensitive_values=set()),
    )


def test_click_with_relative_point_taps_clickable_span_hotspot(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    driver = _Driver()
    element = _Element()
    runner = _runner(driver)
    monkeypatch.setattr(runner, "_element", lambda *_args, **_kwargs: element)

    runner._click(cast(ExcelCase, None), _step("0.84,0.5"))

    assert driver.scripts == [("mobile: clickGesture", {"x": 1579, "y": 1997})]
    assert element.clicked is False


def test_click_without_relative_point_preserves_normal_element_click(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    driver = _Driver()
    element = _Element()
    runner = _runner(driver)
    monkeypatch.setattr(runner, "_element", lambda *_args, **_kwargs: element)

    runner._click(cast(ExcelCase, None), _step(None))

    assert element.clicked is True
    assert driver.scripts == []


@pytest.mark.parametrize("value", ["0.5", "left,0.5", "1.01,0.5", "nan,0.5"])
def test_relative_click_rejects_invalid_coordinates(value: str) -> None:
    with pytest.raises(ValueError):
        _parse_relative_point(value)
