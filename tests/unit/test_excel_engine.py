from __future__ import annotations

from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import cast

import pytest
from appium.webdriver.common.appiumby import AppiumBy
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from selenium.common.exceptions import InvalidSessionIdException

from yanjia_automation.driver import is_recoverable_driver_error
from yanjia_automation.excel.locators import (
    LocatorFormatError,
    parse_locator,
    parse_locator_candidates,
)
from yanjia_automation.excel.results import ExcelResult, ExcelResultWriter
from yanjia_automation.excel.validation import validate_cases
from yanjia_automation.excel.variables import VariableResolver
from yanjia_automation.excel.workbook import ExcelCaseRepository


def test_repository_loads_enabled_case_and_ordered_steps(tmp_path: Path) -> None:
    workbook_path = tmp_path / "cases.xlsx"
    _create_workbook(workbook_path)

    cases = ExcelCaseRepository(workbook_path).load_cases()

    assert [case.case_id for case in cases] == ["TC-DEMO-001"]
    assert [step.order for step in cases[0].steps] == [1, 2]
    assert cases[0].tags == ("smoke", "readonly")


def test_variable_resolution_and_redaction() -> None:
    resolver = VariableResolver(
        {"NAME": "演示顾客", "PASSWORD": "secret-value"},
        sensitive_values={"secret-value"},
    )

    assert resolver.resolve("你好，${NAME}") == "你好，演示顾客"
    assert resolver.resolve("<EMPTY>") == ""
    assert resolver.redact("password=secret-value") == "password=***"


def test_locator_parser() -> None:
    assert parse_locator("id=com.example:id/button") == (
        AppiumBy.ID,
        "com.example:id/button",
    )
    assert parse_locator("accessibility_id=返回") == (AppiumBy.ACCESSIBILITY_ID, "返回")
    assert parse_locator_candidates(
        "id=com.example:id/button || accessibility_id=保存"
    ) == (
        (AppiumBy.ID, "com.example:id/button"),
        (AppiumBy.ACCESSIBILITY_ID, "保存"),
    )
    with pytest.raises(LocatorFormatError):
        parse_locator_candidates("id=button || ")


def test_validator_reports_missing_variable_and_bad_action(tmp_path: Path) -> None:
    workbook_path = tmp_path / "cases.xlsx"
    _create_workbook(workbook_path)
    case = ExcelCaseRepository(workbook_path).load_cases()[0]
    bad_step = replace(
        case.steps[0],
        action="unsupported_action",
        expected="$" + "{MISSING_VALUE}",
    )
    report = validate_cases(
        [replace(case, steps=(bad_step,))],
        VariableResolver({}, sensitive_values=set()),
    )

    assert not report.is_valid
    assert any("不支持的操作类型" in issue.message for issue in report.errors)
    assert any("MISSING_VALUE" in issue.message for issue in report.errors)


def test_infrastructure_error_classifier_never_retries_assertions() -> None:
    assert is_recoverable_driver_error(InvalidSessionIdException("invalid session id"))
    assert not is_recoverable_driver_error(AssertionError("页面业务断言失败"))


def test_result_writer_updates_latest_result_and_history(tmp_path: Path) -> None:
    workbook_path = tmp_path / "cases.xlsx"
    _create_workbook(workbook_path)
    writer = ExcelResultWriter(
        source_path=workbook_path,
        output_path=None,
        run_id="run-001",
        allure_report_dir=tmp_path / "allure-report",
        enabled=True,
    )
    writer.prepare(("TC-DEMO-001",))
    started = datetime(2026, 7, 30, 10, 0, 0)
    finished = datetime(2026, 7, 30, 10, 0, 1)

    writer.record(
        ExcelResult(
            case_id="TC-DEMO-001",
            status="PASS",
            started_at=started,
            finished_at=finished,
            duration_seconds=1.25,
        )
    )

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        cases = workbook["自动化测试用例"]
        headers: dict[str, int] = {}
        for cell in cases[1]:
            if isinstance(cell.value, str) and isinstance(cell.column, int):
                headers[cell.value] = cell.column
        assert cases.cell(2, headers["实际结果"]).value == "PASS"
        assert cases.cell(2, headers["运行编号"]).value == "run-001"
        history = workbook["执行记录"]
        assert history.cell(2, 2).value == "TC-DEMO-001"
        assert history.cell(2, 5).value == "PASS"
    finally:
        workbook.close()


def _create_workbook(path: Path) -> None:
    workbook = Workbook()
    cases = cast(Worksheet, workbook.active)
    cases.title = "自动化测试用例"
    cases.append(
        (
            "用例ID",
            "模块",
            "测试场景",
            "测试点",
            "优先级",
            "前置条件",
            "输入数据",
            "期望结果",
            "超时(秒)",
            "实际结果",
            "自动化状态",
            "标签",
            "是否执行",
        )
    )
    cases.append(
        (
            "TC-DEMO-001",
            "演示模块",
            "演示场景",
            "演示测试点",
            "P1",
            "已进入首页",
            None,
            "按钮可见",
            5,
            "NOT_RUN",
            "AUTOMATED",
            "smoke,readonly",
            "是",
        )
    )
    cases.append(
        (
            "TC-DEMO-002",
            "演示模块",
            "禁用场景",
            "禁用测试点",
            "P2",
            None,
            None,
            None,
            5,
            "NOT_RUN",
            "BACKLOG",
            "readonly",
            "否",
        )
    )

    steps = workbook.create_sheet("自动化执行步骤")
    steps.append(
        (
            "用例ID",
            "步骤序号",
            "步骤名称",
            "操作类型",
            "元素定位器",
            "输入数据",
            "断言类型",
            "期望值",
            "超时(秒)",
            "元素索引",
            "失败继续",
            "是否执行",
            "备注",
        )
    )
    steps.append(
        (
            "TC-DEMO-001",
            2,
            "断言按钮",
            "assert",
            "id=com.example:id/button",
            None,
            "element_visible",
            None,
            5,
            0,
            "否",
            "是",
            None,
        )
    )
    steps.append(
        (
            "TC-DEMO-001",
            1,
            "恢复首页",
            "restart_to_home",
            None,
            None,
            None,
            None,
            30,
            0,
            "否",
            "是",
            None,
        )
    )
    workbook.save(path)
    workbook.close()
