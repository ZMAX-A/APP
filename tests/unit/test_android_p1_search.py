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


def test_repository_migrates_read_only_search_slice(tmp_path: Path) -> None:
    workbook_path = tmp_path / "p1-search.xlsx"
    workbook = Workbook()
    sheet = cast(Worksheet, workbook.active)
    sheet.title = "自动化测试用例"
    sheet.append(HEADERS)
    definitions = (
        ("TC-HOME-002", "首页搜索", "不存在的用户xyz123"),
        ("TC-HOME-003", "首页搜索", "100000000000"),
        ("TC-HOME-004", "首页搜索", "t"),
        ("TC-HOME-005", "首页搜索", "田"),
        ("TC-HOME-006", "首页搜索", "186"),
        ("TC-HOME-007", "首页搜索", "${SEEDED_CUSTOMER_PHONE}"),
        ("TC-CUSTOMER-001", "顾客列表", "不存在的用户xyz123"),
        ("TC-CUSTOMER-002", "顾客列表", "t"),
        ("TC-CUSTOMER-003", "顾客列表", "田"),
        ("TC-CUSTOMER-004", "顾客列表", "100000000000"),
        ("TC-CUSTOMER-005", "顾客列表", "186"),
        ("TC-CUSTOMER-006", "顾客列表", "${SEEDED_CUSTOMER_PHONE}"),
    )
    for case_id, module, input_data in definitions:
        sheet.append(
            (
                case_id,
                module,
                "只读搜索",
                "验证搜索结果",
                "P1",
                "已登录并按用例准备页面或种子数据",
                input_data,
                "展示预期搜索结果",
                10,
                "是",
            )
        )
    workbook.save(workbook_path)
    workbook.close()

    cases = ExcelCaseRepository(workbook_path).load_cases(
        case_patterns=("TC-HOME-*", "TC-CUSTOMER-*")
    )
    by_id = {case.case_id: case for case in cases}

    assert list(by_id) == [case_id for case_id, _, _ in definitions]
    assert by_id["TC-HOME-002"].steps[1].locator == (
        "id=com.xiaofutech.yanjia_ai:id/main_search_et"
    )
    assert by_id["TC-HOME-002"].steps[1].input_value == "不存在的用户xyz123"
    assert by_id["TC-HOME-002"].steps[-1].locator == ("id=com.xiaofutech.yanjia_ai:id/empty_tv")
    assert by_id["TC-HOME-002"].steps[-1].assertion == "element_visible"
    assert by_id["TC-HOME-004"].steps[-1].assertion == "element_count_gte"
    assert by_id["TC-HOME-004"].steps[-1].expected == "1"
    assert "requires_seed" in by_id["TC-HOME-004"].tags
    assert by_id["TC-HOME-006"].steps[-1].locator == (
        "id=com.xiaofutech.yanjia_ai:id/a_records_unique_tv"
    )
    assert by_id["TC-HOME-006"].steps[-1].assertion == "text_contains"
    assert by_id["TC-HOME-006"].steps[-1].expected == "186"
    assert by_id["TC-HOME-007"].steps[-2].assertion == "element_count_equals"
    assert by_id["TC-HOME-007"].steps[-2].expected == "1"
    assert by_id["TC-HOME-007"].steps[-1].locator == (
        "id=com.xiaofutech.yanjia_ai:id/a_records_unique_tv"
    )
    assert by_id["TC-HOME-007"].steps[-1].assertion == "text_equals"
    assert by_id["TC-HOME-007"].steps[-1].expected == "${SEEDED_CUSTOMER_PHONE}"
    assert "requires_seed" in by_id["TC-HOME-007"].tags

    assert by_id["TC-CUSTOMER-001"].steps[1].locator == (
        "id=com.xiaofutech.yanjia_ai:id/main_records_ll"
    )
    assert by_id["TC-CUSTOMER-001"].steps[3].locator == (
        "id=com.xiaofutech.yanjia_ai:id/customer_records_search_et"
    )
    assert by_id["TC-CUSTOMER-001"].steps[-1].locator == ("id=com.xiaofutech.yanjia_ai:id/empty_tv")
    assert by_id["TC-CUSTOMER-006"].steps[-1].assertion == "element_count_gte"
    assert by_id["TC-CUSTOMER-006"].steps[-1].expected == "1"
    assert "requires_seed" in by_id["TC-CUSTOMER-006"].tags
