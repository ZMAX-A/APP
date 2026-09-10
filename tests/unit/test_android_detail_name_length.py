from __future__ import annotations

from scripts.prepare_excel import CURRENT_STEPS
from yanjia_automation.excel.safety import CUSTOMER_MUTATION_CASE_IDS
from yanjia_automation.excel.single_sheet import _detail_validation_steps


def _step(name: str) -> tuple[object, ...]:
    return next(step for step in CURRENT_STEPS["TC-DETAIL-003"] if step[0] == name)


def test_detail_name_length_uses_dedicated_customer_and_recovery_guard() -> None:
    steps = CURRENT_STEPS["TC-DETAIL-003"]
    assert len(steps) == 13
    assert "TC-DETAIL-003" in CUSTOMER_MUTATION_CASE_IDS
    assert _step("输入专用顾客查询")[3] == "${YANJIA_MUTATION_CUSTOMER_QUERY}"
    assert _step("校验专用顾客唯一匹配")[4:6] == ("element_count_equals", "1")
    assert [step[1] for step in steps[-3:]] == ["input", "click", "assert"]


def test_detail_name_length_matches_probed_native_contract() -> None:
    boundary = _step("输入33字符姓名")
    assert boundary[2] == (
        "id=com.xiaofutech.yanjia_ai:id/customer_edit_username_et"
    )
    assert len(str(boundary[3])) == 33
    assert "强制恢复" in str(boundary[10])

    validation = _step("校验姓名长度提示")
    assert validation[4:6] == ("page_source_contains", "客户名称长度不合法")
    assert "重新打开编辑页比对" in str(validation[10])


def test_single_sheet_adapter_builds_the_same_guarded_mutation_tail() -> None:
    steps = _detail_validation_steps(
        row=45,
        case_id="TC-DETAIL-003",
        input_data="DQWDQWDWDWDWDWDWDWDDDD11111112EF1",
        timeout=10,
    )

    assert len(steps) == 13
    assert [step.action for step in steps[-3:]] == ["input", "click", "assert"]
    assert steps[-3].input_value == "DQWDQWDWDWDWDWDWDWDDDD11111112EF1"
    assert steps[-1].assertion == "page_source_contains"
    assert steps[-1].expected == "客户名称长度不合法"
