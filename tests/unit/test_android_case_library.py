from __future__ import annotations

from pathlib import Path
from typing import cast

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from yanjia_automation.excel.workbook import ExcelCaseRepository

HEADERS = (
    "用例ID",
    "模块",
    "测试场景",
    "测试点",
    "优先级",
    "前置条件",
    "输入数据",
    "期望结果",
    "超时(秒)",
    "是否执行",
)


def test_repository_migrates_readonly_case_library_p1_slice(tmp_path: Path) -> None:
    workbook_path = tmp_path / "case-library-p1.xlsx"
    workbook = Workbook()
    sheet = cast(Worksheet, workbook.active)
    sheet.title = "自动化测试用例"
    sheet.append(HEADERS)
    sheet.append(
        (
            "TC-CASE-002",
            "案例库",
            "标签搜索-无结果",
            "验证空结果",
            "P2",
            "已登录",
            "不存在的标签xyz",
            "显示空结果且没有案例卡片",
            10,
            "是",
        )
    )
    sheet.append(
        (
            "TC-CASE-003",
            "案例库",
            "标签搜索-有结果",
            "验证标签搜索",
            "P1",
            "已登录且存在火标签案例",
            "火",
            "结果卡片包含火标签",
            10,
            "是",
        )
    )
    sheet.append(
        (
            "TC-CASE-004",
            "案例库",
            "标签筛选-单选",
            "验证标签筛选",
            "P1",
            "已登录且存在火标签案例",
            None,
            "结果卡片包含火标签",
            10,
            "是",
        )
    )
    workbook.save(workbook_path)
    workbook.close()

    cases = ExcelCaseRepository(workbook_path).load_cases(
        case_patterns=("TC-CASE-002", "TC-CASE-003", "TC-CASE-004")
    )
    by_id = {case.case_id: case for case in cases}

    empty = by_id["TC-CASE-002"]
    assert empty.steps[4].input_value == "不存在的标签xyz"
    assert empty.steps[-2].locator == ("id=com.xiaofutech.yanjia_ai:id/empty_tv")
    assert empty.steps[-2].assertion == "element_visible"
    assert empty.steps[-1].locator == ("id=com.xiaofutech.yanjia_ai:id/a_case_image_cl")
    assert empty.steps[-1].assertion == "element_not_visible"

    search = by_id["TC-CASE-003"]
    assert [step.action for step in search.steps] == [
        "restart_to_home",
        "click",
        "assert",
        "assert",
        "input",
        "click",
        "assert",
        "assert",
    ]
    assert search.steps[4].locator == ("id=com.xiaofutech.yanjia_ai:id/case_search_et")
    assert search.steps[4].input_value == "火"
    assert search.steps[-1].locator == (
        "xpath=//*[@resource-id='com.xiaofutech.yanjia_ai:id/a_case_image_cl']//*[@text='火']"
    )
    assert search.steps[-1].assertion == "element_count_gte"
    assert search.steps[-1].expected == "1"

    case_filter = by_id["TC-CASE-004"]
    assert case_filter.steps[4].locator == ("id=com.xiaofutech.yanjia_ai:id/case_tag_expand_tv")
    assert case_filter.steps[5].locator == (
        "xpath=//*[@resource-id='com.xiaofutech.yanjia_ai:id/a_case_tag_tv' and @text='火']"
    )
    assert case_filter.steps[-1].locator == search.steps[-1].locator
    assert search.tags == (
        "tablet",
        "readonly",
        "case-library",
        "requires_seed",
    )
    assert case_filter.tags == search.tags
