from __future__ import annotations

import time
from collections.abc import Callable
from contextlib import AbstractContextManager
from typing import Any, cast

import allure
from appium.webdriver.common.appiumby import AppiumBy
from appium.webdriver.webdriver import WebDriver
from appium.webdriver.webelement import WebElement
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait

from yanjia_automation.config import Settings
from yanjia_automation.excel.locators import (
    LocatorCandidates,
    parse_locator_candidates,
)
from yanjia_automation.excel.models import ExcelCase, ExcelStep
from yanjia_automation.excel.variables import VariableResolver
from yanjia_automation.flows.navigation import ensure_home
from yanjia_automation.flows.recovery import restart_to_home, restart_to_login
from yanjia_automation.screens.login import LoginScreen

ASSERTION_NAMES = {
    "activity_equals",
    "activity_endswith",
    "all_elements_visible",
    "attribute_contains",
    "attribute_equals",
    "element_count_equals",
    "element_count_gte",
    "element_enabled",
    "element_exists",
    "element_not_visible",
    "element_selected",
    "element_visible",
    "orientation_equals",
    "package_equals",
    "page_source_contains",
    "text_contains",
    "text_equals",
    "text_not_empty",
}

SUPPORTED_ACTIONS = {
    "activate_app",
    "assert",
    "back",
    "background_app",
    "clear",
    "click",
    "ensure_home",
    "hide_keyboard",
    "input",
    "long_click",
    "noop",
    "pause",
    "press_keycode",
    "restart_to_home",
    "restart_to_login",
    "select_first_store",
    "scroll_to_text",
    "set_orientation",
    "swipe",
    "terminate_app",
    "verify",
    "wait_invisible",
    "wait_visible",
}

LOCATOR_REQUIRED_ACTIONS = {
    "clear",
    "click",
    "input",
    "long_click",
    "wait_invisible",
    "wait_visible",
}

LOCATOR_REQUIRED_ASSERTIONS = {
    "all_elements_visible",
    "attribute_contains",
    "attribute_equals",
    "element_count_equals",
    "element_count_gte",
    "element_enabled",
    "element_exists",
    "element_not_visible",
    "element_selected",
    "element_visible",
    "text_contains",
    "text_equals",
    "text_not_empty",
}


class ExcelCaseExecutionError(AssertionError):
    """Raised after one or more soft-failure Excel steps fail."""


class ExcelCaseRunner:
    def __init__(
        self,
        driver: WebDriver,
        settings: Settings,
        variables: VariableResolver,
    ) -> None:
        self.driver = driver
        self.settings = settings
        self.variables = variables

    def run(self, case: ExcelCase) -> None:
        failures: list[str] = []
        for step in case.steps:
            title = f"{step.order}. {step.name} [{step.action}]"
            try:
                step_context = cast(AbstractContextManager[Any], allure.step(title))
                with step_context:
                    self._execute(case, step)
            except Exception as error:
                safe_error = self.variables.redact(error)
                allure.attach(
                    safe_error,
                    name=f"步骤{step.order}错误",
                    attachment_type=allure.attachment_type.TEXT,
                )
                if not step.continue_on_failure:
                    raise
                failures.append(f"步骤{step.order}：{safe_error}")

        if failures:
            raise ExcelCaseExecutionError("；".join(failures))

    def attach_failure_evidence(self) -> None:
        allure.attach(
            self.driver.current_activity or "<unknown>",
            name="失败时Activity",
            attachment_type=allure.attachment_type.TEXT,
        )
        if not self.settings.capture_sensitive_artifacts:
            return
        allure.attach(
            self.driver.get_screenshot_as_png(),
            name="失败截图-可能包含敏感信息",
            attachment_type=allure.attachment_type.PNG,
        )
        allure.attach(
            self.driver.page_source,
            name="页面结构-可能包含敏感信息",
            attachment_type=allure.attachment_type.XML,
        )

    def _execute(self, case: ExcelCase, step: ExcelStep) -> None:
        actions: dict[str, Callable[[ExcelCase, ExcelStep], None]] = {
            "activate_app": self._activate_app,
            "assert": self._assert_step,
            "back": self._back,
            "background_app": self._background_app,
            "clear": self._clear,
            "click": self._click,
            "ensure_home": self._ensure_home,
            "hide_keyboard": self._hide_keyboard,
            "input": self._input,
            "long_click": self._long_click,
            "noop": self._noop,
            "pause": self._pause,
            "press_keycode": self._press_keycode,
            "restart_to_home": self._restart_to_home,
            "restart_to_login": self._restart_to_login,
            "select_first_store": self._select_first_store,
            "scroll_to_text": self._scroll_to_text,
            "set_orientation": self._set_orientation,
            "swipe": self._swipe,
            "terminate_app": self._terminate_app,
            "verify": self._assert_step,
            "wait_invisible": self._wait_invisible,
            "wait_visible": self._wait_visible,
        }
        if step.action in ASSERTION_NAMES and not step.assertion:
            self._apply_assertion(case, step, step.action)
            return
        action = actions.get(step.action)
        if action is None:
            supported = ", ".join(sorted(set(actions) | ASSERTION_NAMES))
            raise ValueError(f"不支持的操作类型：{step.action!r}；支持：{supported}")
        action(case, step)

    def _activate_app(self, case: ExcelCase, step: ExcelStep) -> None:
        self.driver.activate_app(self.settings.app_package)

    def _terminate_app(self, case: ExcelCase, step: ExcelStep) -> None:
        self.driver.terminate_app(self.settings.app_package)

    def _restart_to_home(self, case: ExcelCase, step: ExcelStep) -> None:
        restart_to_home(self.driver, self.settings)

    def _restart_to_login(self, case: ExcelCase, step: ExcelStep) -> None:
        restart_to_login(self.driver, self.settings)

    def _select_first_store(self, case: ExcelCase, step: ExcelStep) -> None:
        LoginScreen(self.driver).select_first_store(timeout=step.timeout)

    def _ensure_home(self, case: ExcelCase, step: ExcelStep) -> None:
        ensure_home(self.driver, self.settings)

    def _click(self, case: ExcelCase, step: ExcelStep) -> None:
        element = self._element(step, require_visible=True)
        WebDriverWait(self.driver, step.timeout).until(
            lambda _: element.is_displayed() and element.is_enabled()
        )
        element.click()

    def _long_click(self, case: ExcelCase, step: ExcelStep) -> None:
        duration = self._resolved_number(step.input_value, default=1.0)
        element = self._element(step, require_visible=True)
        ActionChains(self.driver).click_and_hold(element).pause(duration).release().perform()

    def _input(self, case: ExcelCase, step: ExcelStep) -> None:
        value = self._resolved_input(case, step)
        element = self._element(step, require_visible=True)
        element.clear()
        if value:
            element.send_keys(value)

    def _clear(self, case: ExcelCase, step: ExcelStep) -> None:
        self._element(step, require_visible=True).clear()

    def _noop(self, case: ExcelCase, step: ExcelStep) -> None:
        return

    def _back(self, case: ExcelCase, step: ExcelStep) -> None:
        self.driver.back()

    def _hide_keyboard(self, case: ExcelCase, step: ExcelStep) -> None:
        try:
            self.driver.hide_keyboard()
        except Exception as error:
            if "keyboard" not in str(error).lower():
                raise

    def _press_keycode(self, case: ExcelCase, step: ExcelStep) -> None:
        keycode = int(self._required_resolved(step.input_value, "按键码"))
        self.driver.press_keycode(keycode)

    def _pause(self, case: ExcelCase, step: ExcelStep) -> None:
        seconds = self._resolved_number(step.input_value, default=1.0)
        if seconds < 0 or seconds > 30:
            raise ValueError("pause只允许0至30秒")
        time.sleep(seconds)

    def _background_app(self, case: ExcelCase, step: ExcelStep) -> None:
        seconds = self._resolved_number(step.input_value, default=1.0)
        self.driver.background_app(round(seconds))

    def _set_orientation(self, case: ExcelCase, step: ExcelStep) -> None:
        orientation = self._required_resolved(step.input_value, "屏幕方向").upper()
        if orientation not in {"LANDSCAPE", "PORTRAIT"}:
            raise ValueError("屏幕方向只能是LANDSCAPE或PORTRAIT")
        self.driver.orientation = orientation

    def _swipe(self, case: ExcelCase, step: ExcelStep) -> None:
        raw = self._required_resolved(step.input_value, "滑动参数")
        parts = [part.strip() for part in raw.split(",")]
        if len(parts) not in {4, 5}:
            raise ValueError("swipe输入格式：起点X,起点Y,终点X,终点Y[,时长毫秒]")
        coordinates = [float(part) for part in parts[:4]]
        size = self.driver.get_window_size()

        def pixels(value: float, total: int) -> int:
            return round(value * total) if 0 <= value <= 1 else round(value)

        start_x = pixels(coordinates[0], size["width"])
        start_y = pixels(coordinates[1], size["height"])
        end_x = pixels(coordinates[2], size["width"])
        end_y = pixels(coordinates[3], size["height"])
        duration = int(parts[4]) if len(parts) == 5 else 500
        self.driver.swipe(start_x, start_y, end_x, end_y, duration)

    def _scroll_to_text(self, case: ExcelCase, step: ExcelStep) -> None:
        text = self._required_resolved(step.input_value, "滚动目标文字")
        escaped = text.replace("\\", "\\\\").replace('"', '\\"')
        selector = (
            "new UiScrollable(new UiSelector().scrollable(true))"
            f'.scrollIntoView(new UiSelector().text("{escaped}"))'
        )
        WebDriverWait(self.driver, step.timeout).until(
            lambda current: current.find_element(AppiumBy.ANDROID_UIAUTOMATOR, selector)
        )

    def _wait_visible(self, case: ExcelCase, step: ExcelStep) -> None:
        self._element(step, require_visible=True)

    def _wait_invisible(self, case: ExcelCase, step: ExcelStep) -> None:
        locators = parse_locator_candidates(step.locator)
        WebDriverWait(self.driver, step.timeout).until(
            lambda current: not any(
                element.is_displayed()
                for locator in locators
                for element in current.find_elements(*locator)
            )
        )

    def _assert_step(self, case: ExcelCase, step: ExcelStep) -> None:
        if not step.assertion:
            raise ValueError("assert/verify步骤必须填写断言类型")
        self._apply_assertion(case, step, step.assertion)

    def _apply_assertion(self, case: ExcelCase, step: ExcelStep, assertion: str) -> None:
        expected = self.variables.resolve(step.expected)
        if assertion == "activity_endswith":
            suffix = self._required(expected, "期望Activity后缀")
            self._wait_condition(
                lambda: bool(self.driver.current_activity or "")
                and (self.driver.current_activity or "").endswith(suffix),
                step,
                "当前Activity不符合期望后缀",
            )
        elif assertion == "activity_equals":
            target = self._required(expected, "期望Activity")
            self._wait_condition(
                lambda: self.driver.current_activity == target,
                step,
                "当前Activity不等于期望值",
            )
        elif assertion == "package_equals":
            target = self._required(expected, "期望包名")
            self._wait_condition(
                lambda: self.driver.current_package == target,
                step,
                "当前包名不等于期望值",
            )
        elif assertion == "orientation_equals":
            target = self._required(expected, "期望屏幕方向").upper()
            self._wait_condition(
                lambda: self.driver.orientation.upper() == target,
                step,
                "当前屏幕方向不等于期望值",
            )
        elif assertion == "element_visible":
            self._element(step, require_visible=True)
        elif assertion == "element_exists":
            self._element(step, require_visible=False)
        elif assertion == "element_not_visible":
            self._wait_invisible(case, step)
        elif assertion == "all_elements_visible":
            locators = parse_locator_candidates(step.locator)

            def all_visible() -> bool:
                for locator in locators:
                    elements = self.driver.find_elements(*locator)
                    if elements and all(item.is_displayed() for item in elements):
                        return True
                return False

            self._wait_condition(
                all_visible,
                step,
                "并非所有匹配元素都可见",
            )
        elif assertion in {"element_count_gte", "element_count_equals"}:
            target_count = int(self._required(expected, "期望元素数量"))
            locators = parse_locator_candidates(step.locator)
            compare = (
                (lambda count: count >= target_count)
                if assertion == "element_count_gte"
                else (lambda count: count == target_count)
            )
            self._wait_condition(
                lambda: any(
                    compare(len(self.driver.find_elements(*locator)))
                    for locator in locators
                ),
                step,
                "元素数量不符合期望",
            )
        elif assertion == "text_not_empty":
            element = self._element(step, require_visible=True)
            self._wait_condition(
                lambda: bool(element.text.strip()), step, "元素文本为空"
            )
        elif assertion in {"text_equals", "text_contains"}:
            target = self._required(expected, "期望文本")
            element = self._element(step, require_visible=True)
            compare_text = (
                (lambda text: text == target)
                if assertion == "text_equals"
                else (lambda text: target in text)
            )
            self._wait_condition(
                lambda: compare_text(element.text), step, "元素文本不符合期望"
            )
        elif assertion in {"attribute_equals", "attribute_contains"}:
            raw = self._required(expected, "属性断言参数")
            if "|" not in raw:
                raise ValueError("属性断言期望值格式：属性名|期望值")
            attribute, target = raw.split("|", 1)
            element = self._element(step, require_visible=False)
            compare_attribute = (
                (lambda value: value == target)
                if assertion == "attribute_equals"
                else (lambda value: target in (value or ""))
            )
            self._wait_condition(
                lambda: compare_attribute(element.get_attribute(attribute)),
                step,
                "元素属性不符合期望",
            )
        elif assertion == "element_enabled":
            element = self._element(step, require_visible=True)
            self._wait_condition(element.is_enabled, step, "元素未启用")
        elif assertion == "element_selected":
            element = self._element(step, require_visible=False)
            self._wait_condition(element.is_selected, step, "元素未选中")
        elif assertion == "page_source_contains":
            target = self._required(expected, "期望页面文本")
            self._wait_condition(
                lambda: target in self.driver.page_source,
                step,
                "页面结构不包含期望内容",
            )
        else:
            supported = ", ".join(sorted(ASSERTION_NAMES))
            raise ValueError(f"不支持的断言类型：{assertion!r}；支持：{supported}")

    def _element(self, step: ExcelStep, *, require_visible: bool) -> WebElement:
        locators = parse_locator_candidates(step.locator)

        def find(current: WebDriver) -> WebElement | bool:
            for locator in locators:
                elements = current.find_elements(*locator)
                index = step.element_index
                if index >= 0:
                    if len(elements) <= index:
                        continue
                elif len(elements) < abs(index):
                    continue
                element = elements[index]
                if require_visible and not element.is_displayed():
                    continue
                return element
            return False

        try:
            result = WebDriverWait(self.driver, step.timeout).until(find)
            return cast(WebElement, result)
        except TimeoutException as error:
            raise AssertionError(
                f"在{step.timeout:g}秒内未找到符合条件的元素：{step.locator}"
            ) from error

    def _wait_condition(
        self,
        condition: Callable[[], bool],
        step: ExcelStep,
        message: str,
    ) -> None:
        try:
            WebDriverWait(self.driver, step.timeout).until(lambda _: condition())
        except TimeoutException as error:
            raise AssertionError(f"{message}（超时{step.timeout:g}秒）") from error

    def _resolved_input(self, case: ExcelCase, step: ExcelStep) -> str:
        value = step.input_value if step.input_value is not None else case.input_data
        resolved = self.variables.resolve(value)
        return resolved or ""

    def _required_resolved(self, value: str | None, label: str) -> str:
        return self._required(self.variables.resolve(value), label)

    def _resolved_number(self, value: str | None, *, default: float) -> float:
        resolved = self.variables.resolve(value)
        return default if resolved is None or resolved == "" else float(resolved)

    @staticmethod
    def _required(value: str | None, label: str) -> str:
        if value is None or value == "":
            raise ValueError(f"Excel步骤缺少{label}")
        return value

    @staticmethod
    def locator_for_debug(step: ExcelStep) -> LocatorCandidates | None:
        return parse_locator_candidates(step.locator) if step.locator else None
