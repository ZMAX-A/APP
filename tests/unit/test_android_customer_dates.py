from __future__ import annotations

from pathlib import Path
from typing import cast

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from yanjia_automation.excel.workbook import ExcelCaseRepository


def test_repository_migrates_customer_date_ranges_to_native_wheels(
    tmp_path: Path,
) -> None:
    workbook_path = tmp_path / "customer-dates.xlsx"
    workbook = Workbook()
    sheet = cast(Worksheet, workbook.active)
    sheet.title = "自动化测试用例"
    sheet.append(
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
            "是否执行",
        )
    )
    definitions = {
        "TC-CUSTOMER-007": "|2015/02/04-2015/02/13",
        "TC-CUSTOMER-008": "2026-01-01~2026-06-28",
        "TC-CUSTOMER-017": "|2015/02/04-2015/02/28",
        "TC-CUSTOMER-018": "|2026/01/01-2026/07/30",
    }
    for case_id, date_range in definitions.items():
        sheet.append(
            (
                case_id,
                "顾客列表",
                "日期筛选",
                "验证日期筛选结果",
                "P1",
                "已登录并准备受控顾客种子",
                date_range,
                "日期范围符合条件",
                10,
                "是",
            )
        )
    workbook.save(workbook_path)
    workbook.close()

    cases = ExcelCaseRepository(workbook_path).load_cases(
        case_patterns=(
            "TC-CUSTOMER-007",
            "TC-CUSTOMER-008",
            "TC-CUSTOMER-017",
            "TC-CUSTOMER-018",
        )
    )
    by_id = {case.case_id: case for case in cases}

    assert list(by_id) == list(definitions)
    assert by_id["TC-CUSTOMER-008"].steps[5].action == "select_date_range"
    assert by_id["TC-CUSTOMER-008"].steps[5].input_value == ("2026-01-01~2026-06-28")
    assert by_id["TC-CUSTOMER-008"].steps[7].assertion == "text_equals"
    assert by_id["TC-CUSTOMER-008"].steps[7].expected == "2026-01-01~2026-06-28"
    assert by_id["TC-CUSTOMER-008"].steps[-1].assertion == "element_count_gte"
    early_empty = by_id["TC-CUSTOMER-007"]
    assert len(early_empty.steps) == 9
    assert early_empty.steps[5].action == "select_date_range"
    assert early_empty.steps[5].input_value == "2015-02-04~2015-02-13"
    assert early_empty.steps[7].expected == "2015-02-04~2015-02-13"
    assert early_empty.steps[-1].locator == "id=com.xiaofutech.yanjia_ai:id/empty_tv"
    assert early_empty.steps[-1].assertion == "element_visible"
    assert by_id["TC-CUSTOMER-017"].steps[-1].locator == ("id=com.xiaofutech.yanjia_ai:id/empty_tv")
    assert by_id["TC-CUSTOMER-017"].steps[-1].assertion == "element_visible"
    assert by_id["TC-CUSTOMER-017"].steps[7].expected == "2015-02-04~2015-02-28"
    assert [step.action for step in by_id["TC-CUSTOMER-017"].steps[-6:]] == [
        "click",
        "assert",
        "click",
        "click",
        "click",
        "assert",
    ]
    assert by_id["TC-CUSTOMER-017"].steps[10].locator == (
        "id=com.xiaofutech.yanjia_ai:id/customer_records_other_sex_man_tv"
    )
    assert by_id["TC-CUSTOMER-017"].steps[11].locator == (
        "id=com.xiaofutech.yanjia_ai:id/customer_records_other_age_1_tv"
    )
