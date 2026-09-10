from __future__ import annotations

from pathlib import Path
from typing import cast

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from yanjia_automation.excel.workbook import ExcelCaseRepository


def test_repository_migrates_customer_gender_and_age_filters(tmp_path: Path) -> None:
    workbook_path = tmp_path / "customer-filters.xlsx"
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
        "TC-CUSTOMER-009": "男",
        "TC-CUSTOMER-010": "女",
        "TC-CUSTOMER-012": "18-25岁",
        "TC-CUSTOMER-013": "26-35岁",
        "TC-CUSTOMER-014": "36-45岁",
        "TC-CUSTOMER-015": "45岁以上",
        "TC-CUSTOMER-016": "18岁以下",
    }
    for case_id in definitions:
        sheet.append(
            (
                case_id,
                "顾客列表",
                "顾客筛选",
                "验证筛选结果",
                "P1",
                "已登录并准备受控顾客种子",
                definitions[case_id],
                "筛选结果符合条件",
                10,
                "是",
            )
        )
    workbook.save(workbook_path)
    workbook.close()

    cases = ExcelCaseRepository(workbook_path).load_cases(case_patterns=("TC-CUSTOMER-*",))
    by_id = {case.case_id: case for case in cases}

    assert list(by_id) == list(definitions)
    for case_id, _filter_value in definitions.items():
        case = by_id[case_id]
        assert [step.action for step in case.steps] == [
            "restart_to_home",
            "click",
            "assert",
            "click",
            "assert",
            "click",
            "click",
            "assert",
        ]
        assert case.steps[3].locator == ("id=com.xiaofutech.yanjia_ai:id/customer_records_other_tv")
        assert case.steps[4].locator == ("id=com.xiaofutech.yanjia_ai:id/customer_records_other_cl")
        assert case.steps[5].locator is not None
        assert case.steps[5].locator.startswith(
            "id=com.xiaofutech.yanjia_ai:id/customer_records_other_"
        )
        assert case.steps[6].locator == (
            "id=com.xiaofutech.yanjia_ai:id/customer_records_other_confirm_tv"
        )
        if case_id == "TC-CUSTOMER-015":
            assert case.steps[-1].locator == ("id=com.xiaofutech.yanjia_ai:id/empty_tv")
            assert case.steps[-1].assertion == "element_visible"
            assert case.steps[-1].expected is None
        else:
            assert case.steps[-1].assertion == "element_count_gte"
            assert case.steps[-1].expected == "1"
        assert "requires_seed" in case.tags
        assert "mutating" not in case.tags


def test_repository_migrates_customer_filter_reset_contract(tmp_path: Path) -> None:
    workbook_path = tmp_path / "customer-filter-reset.xlsx"
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
    sheet.append(
        (
            "TC-CUSTOMER-022",
            "顾客列表",
            "筛选重置功能",
            "验证重置按钮取消筛选条件",
            "P1",
            "已登录并进入顾客列表页",
            "123",
            "全部筛选条件已取消",
            10,
            "是",
        )
    )
    workbook.save(workbook_path)
    workbook.close()

    [case] = ExcelCaseRepository(workbook_path).load_cases(
        case_patterns=("TC-CUSTOMER-022",)
    )

    assert [step.action for step in case.steps] == [
        "restart_to_home",
        "click",
        "assert",
        "click",
        "assert",
        "click",
        "assert",
        "click",
        "assert",
        "input",
        "assert",
        "click",
        "assert",
        "assert",
        "assert",
        "assert",
    ]
    assert case.steps[6].assertion == "element_selected"
    assert case.steps[8].assertion == "element_selected"
    assert case.steps[9].input_value == "123"
    assert case.steps[10].expected == "text|123"
    assert case.steps[11].locator == (
        "id=com.xiaofutech.yanjia_ai:id/customer_records_other_reset_tv"
    )
    assert [step.expected for step in case.steps[12:15]] == [
        "selected|false",
        "selected|false",
        "text|搜索备注记录",
    ]
    assert case.steps[15].assertion == "element_visible"
    assert case.steps[15].locator == (
        "id=com.xiaofutech.yanjia_ai:id/customer_records_other_cl"
    )
    assert "requires_seed" in case.tags
    assert "mutating" not in case.tags
