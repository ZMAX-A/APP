from __future__ import annotations

import re
import time
from collections.abc import Callable
from contextlib import AbstractContextManager
from datetime import date
from math import isfinite
from typing import Any, cast

import allure
from appium.webdriver.common.appiumby import AppiumBy
from appium.webdriver.webdriver import WebDriver
from appium.webdriver.webelement import WebElement
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait

from yanjia_automation.config import Settings
from yanjia_automation.date_picker import date_after_picker_change, swipe_number_picker_steps
from yanjia_automation.excel.locators import (
    LocatorCandidates,
    parse_locator_candidates,
)
from yanjia_automation.excel.models import ExcelCase, ExcelStep
from yanjia_automation.excel.safety import (
    CUSTOMER_MUTATION_CASE_IDS,
    CUSTOMER_MUTATION_TAIL_ACTIONS,
    REMARK_MUTATION_CASE_IDS,
    REMARK_MUTATION_SCOPES,
    REMARK_MUTATION_TAIL_ACTIONS,
    TAG_MUTATION_CASE_IDS,
    TAG_MUTATION_SCOPES,
    TAG_MUTATION_TAIL_ACTIONS,
)
from yanjia_automation.excel.variables import VariableResolver
from yanjia_automation.flows.customer_mutation import DedicatedCustomerMutationFlow
from yanjia_automation.flows.navigation import ensure_home
from yanjia_automation.flows.recovery import restart_to_home, restart_to_login
from yanjia_automation.flows.remark_mutation import (
    DedicatedRemarkMutationFlow,
    RemarkScope,
    RemarkSnapshot,
)
from yanjia_automation.flows.tag_mutation import (
    DedicatedTagMutationFlow,
    TagScope,
    TagSnapshot,
)
from yanjia_automation.screens.customer import CustomerDetailScreen, CustomerEditScreen
from yanjia_automation.screens.login import LoginScreen

ASSERTION_NAMES = {
    "activity_equals",
    "activity_endswith",
    "activity_not_endswith",
    "all_elements_visible",
    "attribute_contains",
    "attribute_equals",
    "element_count_equals",
    "element_count_gte",
    "element_disabled",
    "element_enabled",
    "element_exists",
    "element_not_visible",
    "element_selected",
    "element_visible",
    "element_count_unchanged_after_click",
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
    "delete_current_run_remark",
    "delete_current_run_tag",
    "ensure_home",
    "hide_keyboard",
    "input",
    "long_click",
    "noop",
    "open_first_unlinked_image",
    "pause",
    "press_keycode",
    "restart_to_home",
    "restart_to_login",
    "select_first_store",
    "scroll_to_text",
    "scroll_profile_remark",
    "select_birthday",
    "select_date_range",
    "set_orientation",
    "swipe",
    "terminate_app",
    "verify",
    "verify_selected_image_consultation",
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
    "element_disabled",
    "element_enabled",
    "element_exists",
    "element_not_visible",
    "element_selected",
    "element_visible",
    "element_count_unchanged_after_click",
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
        self._active_tag_mutation: tuple[DedicatedTagMutationFlow, TagSnapshot] | None = None
        self._active_remark_mutation: tuple[DedicatedRemarkMutationFlow, RemarkSnapshot] | None = (
            None
        )
        self._selected_unlinked_image_date: str | None = None

    def run(self, case: ExcelCase) -> None:
        if case.case_id in REMARK_MUTATION_CASE_IDS:
            self._run_dedicated_remark_mutation(case)
            return
        if case.case_id in TAG_MUTATION_CASE_IDS:
            self._run_dedicated_tag_mutation(case)
            return
        if case.case_id in CUSTOMER_MUTATION_CASE_IDS:
            self._run_dedicated_customer_mutation(case)
            return
        self._run_steps(case, case.steps)

    def _run_steps(self, case: ExcelCase, steps: tuple[ExcelStep, ...]) -> None:
        failures: list[str] = []
        for step in steps:
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

    def _run_dedicated_customer_mutation(self, case: ExcelCase) -> None:
        expected_actions = CUSTOMER_MUTATION_TAIL_ACTIONS[case.case_id]
        mutation_steps = case.steps[-len(expected_actions) :]
        actions = tuple(step.action for step in mutation_steps)
        if actions != expected_actions:
            raise ValueError(f"顾客写入用例 {case.case_id} 的受控步骤尾部不符合恢复契约")

        flow = DedicatedCustomerMutationFlow(
            self.driver,
            self.settings,
            timeout=max(step.timeout for step in case.steps),
        )
        with flow.restoration_session() as restoration:
            restoration.run(lambda: self._run_steps(case, mutation_steps))

    def _run_dedicated_tag_mutation(self, case: ExcelCase) -> None:
        expected_actions = TAG_MUTATION_TAIL_ACTIONS[case.case_id]
        mutation_steps = case.steps[-len(expected_actions) :]
        actions = tuple(step.action for step in mutation_steps)
        if actions != expected_actions:
            raise ValueError(f"标签写入用例 {case.case_id} 的受控步骤尾部不符合恢复契约")

        flow = DedicatedTagMutationFlow(
            self.driver,
            self.settings,
            timeout=max(step.timeout for step in case.steps),
        )
        scope = cast(TagScope, TAG_MUTATION_SCOPES[case.case_id])
        with flow.restoration_session(scope) as restoration:
            self._active_tag_mutation = (flow, restoration.snapshot)
            try:
                restoration.run(lambda: self._run_steps(case, mutation_steps))
            finally:
                self._active_tag_mutation = None

    def _run_dedicated_remark_mutation(self, case: ExcelCase) -> None:
        expected_actions = REMARK_MUTATION_TAIL_ACTIONS[case.case_id]
        mutation_steps = case.steps[-len(expected_actions) :]
        actions = tuple(step.action for step in mutation_steps)
        if actions != expected_actions:
            raise ValueError(f"备注写入用例 {case.case_id} 的受控步骤尾部不符合恢复契约")

        flow = DedicatedRemarkMutationFlow(
            self.driver,
            self.settings,
            timeout=max(step.timeout for step in case.steps),
        )
        scope = cast(RemarkScope, REMARK_MUTATION_SCOPES[case.case_id])
        with flow.restoration_session(scope) as restoration:
            self._active_remark_mutation = (flow, restoration.snapshot)
            try:
                restoration.run(lambda: self._run_steps(case, mutation_steps))
            finally:
                self._active_remark_mutation = None

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
            "delete_current_run_remark": self._delete_current_run_remark,
            "delete_current_run_tag": self._delete_current_run_tag,
            "ensure_home": self._ensure_home,
            "hide_keyboard": self._hide_keyboard,
            "input": self._input,
            "long_click": self._long_click,
            "noop": self._noop,
            "open_first_unlinked_image": self._open_first_unlinked_image,
            "pause": self._pause,
            "press_keycode": self._press_keycode,
            "restart_to_home": self._restart_to_home,
            "restart_to_login": self._restart_to_login,
            "select_first_store": self._select_first_store,
            "scroll_to_text": self._scroll_to_text,
            "scroll_profile_remark": self._scroll_profile_remark,
            "select_birthday": self._select_birthday,
            "select_date_range": self._select_date_range,
            "set_orientation": self._set_orientation,
            "swipe": self._swipe,
            "terminate_app": self._terminate_app,
            "verify": self._assert_step,
            "verify_selected_image_consultation": (self._verify_selected_image_consultation),
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
        restart_to_home(self.driver, self.settings, timeout=step.timeout)

    def _restart_to_login(self, case: ExcelCase, step: ExcelStep) -> None:
        restart_to_login(self.driver, self.settings, timeout=step.timeout)

    def _select_first_store(self, case: ExcelCase, step: ExcelStep) -> None:
        LoginScreen(self.driver).select_first_store(timeout=step.timeout)

    def _select_birthday(self, case: ExcelCase, step: ExcelStep) -> None:
        del case
        birthday = self._required_resolved(step.input_value, "生日")
        CustomerEditScreen(self.driver, timeout=step.timeout).select_birthday(birthday)

    def _ensure_home(self, case: ExcelCase, step: ExcelStep) -> None:
        ensure_home(self.driver, self.settings)

    def _click(self, case: ExcelCase, step: ExcelStep) -> None:
        element = self._element(step, require_visible=True)
        WebDriverWait(self.driver, step.timeout).until(
            lambda _: element.is_displayed() and element.is_enabled()
        )
        relative_point = self.variables.resolve(step.input_value)
        if not relative_point:
            element.click()
            return

        relative_x, relative_y = _parse_relative_point(relative_point)
        rect = element.rect
        width = float(rect["width"])
        height = float(rect["height"])
        if width <= 0 or height <= 0:
            raise AssertionError(f"元素尺寸无效，无法执行相对点击：{rect}")
        target_x = round(float(rect["x"]) + (width - 1) * relative_x)
        target_y = round(float(rect["y"]) + (height - 1) * relative_y)
        self.driver.execute_script(
            "mobile: clickGesture",
            {"x": target_x, "y": target_y},
        )

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

    def _open_first_unlinked_image(self, case: ExcelCase, step: ExcelStep) -> None:
        del case
        detail = CustomerDetailScreen(self.driver, timeout=step.timeout)
        images = detail.find_all(detail.images)
        image_dates = [
            _normalized_element_text(item) for item in detail.find_all(detail.image_dates)
        ]
        consultation_dates = _normalized_element_texts(
            detail.find_all(detail.consultation_detection_times)
        )
        if not images or len(images) != len(image_dates):
            raise AssertionError("历史影像与检测时间数量不一致，无法安全选择未关联影像")
        candidate_index = _first_unlinked_image_index(image_dates, consultation_dates)
        if candidate_index is None:
            raise AssertionError("当前可见历史影像均已关联咨询单，未执行新增")
        self._selected_unlinked_image_date = image_dates[candidate_index]
        images[candidate_index].click()

    def _verify_selected_image_consultation(self, case: ExcelCase, step: ExcelStep) -> None:
        del case
        selected_date = self._selected_unlinked_image_date
        if not selected_date:
            raise RuntimeError("尚未通过 open_first_unlinked_image 记录目标影像")
        detail = CustomerDetailScreen(self.driver, timeout=step.timeout)
        WebDriverWait(self.driver, step.timeout).until(
            lambda _: (
                selected_date
                in _normalized_element_texts(detail.find_all(detail.consultation_detection_times))
            )
        )
        self._selected_unlinked_image_date = None

    def _delete_current_run_tag(self, case: ExcelCase, step: ExcelStep) -> None:
        del case, step
        if self._active_tag_mutation is None:
            raise RuntimeError("delete_current_run_tag 只能在受控标签恢复会话中执行")
        flow, snapshot = self._active_tag_mutation
        flow.delete_created_tag(snapshot)

    def _delete_current_run_remark(self, case: ExcelCase, step: ExcelStep) -> None:
        del case, step
        if self._active_remark_mutation is None:
            raise RuntimeError("delete_current_run_remark 只能在受控备注恢复会话中执行")
        flow, snapshot = self._active_remark_mutation
        flow.delete_created_remark(snapshot)

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

    def _select_date_range(self, case: ExcelCase, step: ExcelStep) -> None:
        del case
        target_start, target_end = _parse_date_range(
            self._required_resolved(step.input_value, "日期范围")
        )
        package = self.settings.app_package
        summary_id = f"{package}:id/customer_records_date_tv"
        summaries = self.driver.find_elements(AppiumBy.ID, summary_id)
        if not summaries or not summaries[0].text:
            raise AssertionError("无法读取 Android 日期筛选当前范围")
        current_start, current_end = _parse_date_range(summaries[0].text)

        wheel_ids = (
            f"{package}:id/date_picker_year_wheel",
            f"{package}:id/date_picker_month_wheel",
            f"{package}:id/date_picker_day_wheel",
        )
        if target_start > current_end:
            self._select_date_endpoint(
                wheel_ids,
                1,
                current_end,
                target_end,
                minimum=current_start,
            )
            self._select_date_endpoint(
                wheel_ids,
                0,
                current_start,
                target_start,
                maximum=target_end,
            )
        else:
            self._select_date_endpoint(
                wheel_ids,
                0,
                current_start,
                target_start,
                maximum=current_end,
            )
            self._select_date_endpoint(
                wheel_ids,
                1,
                current_end,
                target_end,
                minimum=target_start,
            )

    def _select_date_endpoint(
        self,
        wheel_ids: tuple[str, str, str],
        element_index: int,
        current: date,
        target: date,
        *,
        minimum: date | None = None,
        maximum: date | None = None,
    ) -> None:
        self._scroll_date_wheel(wheel_ids[0], element_index, current.year, target.year)
        current = date_after_picker_change(
            current,
            year=target.year,
            minimum=minimum,
            maximum=maximum,
        )
        self._scroll_date_wheel(wheel_ids[1], element_index, current.month, target.month)
        current = date_after_picker_change(
            current,
            month=target.month,
            minimum=minimum,
            maximum=maximum,
        )
        self._scroll_date_wheel(wheel_ids[2], element_index, current.day, target.day)

    def _scroll_date_wheel(
        self,
        resource_id: str,
        element_index: int,
        current_value: int,
        target_value: int,
    ) -> None:
        delta = target_value - current_value
        if delta == 0:
            return
        wheels = self.driver.find_elements(AppiumBy.ID, resource_id)
        if len(wheels) <= element_index:
            raise AssertionError(f"日期滚轮数量不足：{resource_id}，需要索引{element_index}")
        swipe_number_picker_steps(self.driver, wheels[element_index].rect, delta)

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

    def _scroll_profile_remark(self, case: ExcelCase, step: ExcelStep) -> None:
        del case
        CustomerEditScreen(self.driver, timeout=step.timeout).scroll_to_remark()

    def _wait_visible(self, case: ExcelCase, step: ExcelStep) -> None:
        self._element(step, require_visible=True)

    def _wait_invisible(self, case: ExcelCase, step: ExcelStep) -> None:
        locators = parse_locator_candidates(step.locator)
        WebDriverWait(self.driver, step.timeout).until(
            lambda current: (
                not any(
                    element.is_displayed()
                    for locator in locators
                    for element in current.find_elements(*locator)
                )
            )
        )

    def _assert_step(self, case: ExcelCase, step: ExcelStep) -> None:
        if not step.assertion:
            raise ValueError("assert/verify步骤必须填写断言类型")
        self._apply_assertion(case, step, step.assertion)

    def _apply_assertion(self, case: ExcelCase, step: ExcelStep, assertion: str) -> None:
        expected = self.variables.resolve(step.expected)
        if assertion in {"activity_endswith", "activity_not_endswith"}:
            suffix = self._required(expected, "期望Activity后缀")
            matches = (
                (lambda activity: activity.endswith(suffix))
                if assertion == "activity_endswith"
                else (lambda activity: not activity.endswith(suffix))
            )
            self._wait_condition(
                lambda: (
                    bool(self.driver.current_activity or "")
                    and matches(self.driver.current_activity or "")
                ),
                step,
                (
                    "当前Activity不符合期望后缀"
                    if assertion == "activity_endswith"
                    else "当前Activity仍为不允许的后缀"
                ),
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
                    compare(len(self.driver.find_elements(*locator))) for locator in locators
                ),
                step,
                "元素数量不符合期望",
            )
        elif assertion == "text_not_empty":
            element = self._element(step, require_visible=True)
            self._wait_condition(lambda: bool(element.text.strip()), step, "元素文本为空")
        elif assertion in {"text_equals", "text_contains"}:
            target = self._required(expected, "期望文本")
            if (
                assertion == "text_contains"
                and self._active_tag_mutation is not None
                and parse_locator_candidates(step.locator)
                == ((AppiumBy.ID, f"{self.settings.app_package}:id/a_records_remark_tag_tv"),)
            ):
                flow, snapshot = self._active_tag_mutation
                flow.assert_created_tag(snapshot, target, timeout=step.timeout)
                return
            element = self._element(step, require_visible=True)
            compare_text = (
                (lambda text: text == target)
                if assertion == "text_equals"
                else (lambda text: target in text)
            )
            self._wait_condition(lambda: compare_text(element.text), step, "元素文本不符合期望")
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
        elif assertion == "element_disabled":
            element = self._element(step, require_visible=True)
            self._wait_condition(lambda: not element.is_enabled(), step, "元素仍处于启用状态")
        elif assertion == "element_enabled":
            element = self._element(step, require_visible=True)
            self._wait_condition(element.is_enabled, step, "元素未启用")
        elif assertion == "element_selected":
            element = self._element(step, require_visible=False)
            self._wait_condition(element.is_selected, step, "元素未选中")
        elif assertion == "element_count_unchanged_after_click":
            observed_locator = self._required(expected, "需要观察的元素定位器")
            self._assert_element_count_unchanged_after_click(step, observed_locator)
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

    def _assert_element_count_unchanged_after_click(
        self,
        step: ExcelStep,
        observed_locator: str,
    ) -> None:
        observed = parse_locator_candidates(observed_locator)
        baseline = tuple(len(self.driver.find_elements(*locator)) for locator in observed)
        starting_activity = self.driver.current_activity
        button = self._element(step, require_visible=True)
        rect = button.rect
        width = float(rect["width"])
        height = float(rect["height"])
        if width <= 0 or height <= 0:
            raise AssertionError(f"元素尺寸无效，无法执行提交点击：{rect}")
        self.driver.execute_script(
            "mobile: clickGesture",
            {
                "x": round(float(rect["x"]) + (width - 1) / 2),
                "y": round(float(rect["y"]) + (height - 1) / 2),
            },
        )

        deadline = time.monotonic() + step.timeout
        while True:
            current = tuple(len(self.driver.find_elements(*locator)) for locator in observed)
            if current != baseline:
                raise AssertionError("点击提交后备注记录数量发生变化")
            if self.driver.current_activity != starting_activity:
                raise AssertionError("点击提交后离开了备注编辑页面")
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return
            time.sleep(min(0.25, remaining))

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


def _parse_date_range(value: str) -> tuple[date, date]:
    raw = value.split("|")[-1].strip()
    match = re.fullmatch(
        r"(\d{4}[/-]\d{1,2}[/-]\d{1,2})\s*(?:~|-)\s*"
        r"(\d{4}[/-]\d{1,2}[/-]\d{1,2})",
        raw,
    )
    if match is None:
        raise ValueError(f"日期范围格式错误：{value!r}")
    dates: list[date] = []
    for part in match.groups():
        components = re.fullmatch(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", part)
        if components is None:
            raise ValueError(f"日期格式错误：{part!r}")
        try:
            dates.append(date(*(int(item) for item in components.groups())))
        except ValueError as error:
            raise ValueError(f"日期不是有效日期：{part!r}") from error
    if dates[0] > dates[1]:
        raise ValueError(f"日期范围起始日期晚于结束日期：{value!r}")
    return dates[0], dates[1]


def _normalized_element_text(element: WebElement) -> str:
    return "".join(str(element.get_attribute("text") or "").split())


def _normalized_element_texts(elements: list[WebElement]) -> set[str]:
    return {text for item in elements if (text := _normalized_element_text(item))}


def _first_unlinked_image_index(image_dates: list[str], consultation_dates: set[str]) -> int | None:
    normalized_consultation_dates = {
        "".join(value.split()) for value in consultation_dates if value.strip()
    }
    for index, value in enumerate(image_dates):
        normalized = "".join(value.split())
        if normalized and normalized not in normalized_consultation_dates:
            return index
    return None


def _parse_relative_point(value: str) -> tuple[float, float]:
    parts = [part.strip() for part in value.split(",")]
    if len(parts) != 2:
        raise ValueError("click输入格式：元素内相对X,元素内相对Y，例如0.84,0.5")
    try:
        coordinates = (float(parts[0]), float(parts[1]))
    except ValueError as error:
        raise ValueError("click相对坐标必须是数字，例如0.84,0.5") from error
    if any(not isfinite(value) or not 0 <= value <= 1 for value in coordinates):
        raise ValueError("click相对坐标必须位于0到1之间")
    return coordinates
