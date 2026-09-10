from __future__ import annotations

import pytest

from scripts.prepare_excel import CURRENT_STEPS
from yanjia_automation.excel.models import ExcelCase
from yanjia_automation.excel.safety import execution_skip_reason
from yanjia_automation.excel.single_sheet import _detail_validation_steps


@pytest.mark.parametrize(
    ("case_id", "expected_length", "input_data", "entry_locator", "has_space"),
    [
        (
            "TC-DETAIL-019",
            19,
            None,
            "a_records_detail_all_remark_ifv",
            True,
        ),
        (
            "TC-DETAIL-032",
            19,
            "咨询单特定| ",
            "a_consultation_result_remark_ifv",
            True,
        ),
        (
            "TC-DETAIL-034",
            18,
            "咨询单特定|",
            "a_consultation_result_remark_ifv",
            False,
        ),
    ],
)
def test_tag_validation_steps_show_prompt_cancel_and_never_delete(
    case_id: str,
    expected_length: int,
    input_data: str | None,
    entry_locator: str,
    has_space: bool,
) -> None:
    canonical = CURRENT_STEPS[case_id]
    migrated = _detail_validation_steps(
        row=61,
        case_id=case_id,
        input_data=input_data,
        timeout=10,
    )

    assert len(canonical) == expected_length
    assert len(migrated) == expected_length
    for steps in (canonical, migrated):
        actions = [step[1] if isinstance(step, tuple) else step.action for step in steps]
        locators = [step[2] if isinstance(step, tuple) else step.locator for step in steps]
        assertions = [step[4] if isinstance(step, tuple) else step.assertion for step in steps]
        expected = [step[5] if isinstance(step, tuple) else step.expected for step in steps]
        inputs = [step[3] if isinstance(step, tuple) else step.input_value for step in steps]

        assert any(str(locator or "").endswith(entry_locator) for locator in locators)
        assert actions[-5:] == ["click", "assert", "click", "click", "assert"]
        assert assertions[-4] == "page_source_contains"
        assert expected[-4] == "请输入标签内容"
        assert str(locators[-3] or "").endswith("records_remark_tag_cancel_tv")
        assert str(locators[-2] or "").endswith("records_remark_close_ifv")
        assert assertions[-1] == "activity_endswith"
        assert expected[-1] == ".CustomerDetailActivity"
        assert not any("delete" in str(locator or "") for locator in locators)
        if has_space:
            assert "input" in actions[-7:]
            assert "${SPACE}" in inputs
        else:
            assert "input" not in actions[8:]


def test_image_whitespace_tag_uses_dedicated_query_without_mutation_gate() -> None:
    case = ExcelCase(
        source_row=61,
        case_id="TC-DETAIL-019",
        module="顾客详情",
        scenario="影像标签空格",
        test_point="校验空格标签提示",
        priority="P1",
        precondition=None,
        input_data=None,
        expected_result=None,
        timeout=10,
        tags=("readonly", "requires_seed"),
        enabled=True,
        automation_status="AUTOMATED",
        steps=(),
    )

    assert execution_skip_reason(
        case,
        run_seeded=True,
        allow_mutation=False,
        allow_destructive=False,
        mutation_customer_query=None,
        customer_preflight_verified=False,
    ) == "专用顾客用例需要配置 YANJIA_MUTATION_CUSTOMER_QUERY"
    assert execution_skip_reason(
        case,
        run_seeded=True,
        allow_mutation=False,
        allow_destructive=False,
        mutation_customer_query="private-customer-query",
        customer_preflight_verified=False,
    ) is None


@pytest.mark.parametrize("case_id", ["TC-DETAIL-032", "TC-DETAIL-034"])
def test_consultation_tag_validation_requires_seed_but_not_dedicated_query(
    case_id: str,
) -> None:
    case = ExcelCase(
        source_row=74,
        case_id=case_id,
        module="顾客详情",
        scenario="咨询单标签校验",
        test_point="校验空标签提示",
        priority="P1",
        precondition=None,
        input_data="咨询单特定|",
        expected_result=None,
        timeout=10,
        tags=("readonly", "requires_seed"),
        enabled=True,
        automation_status="AUTOMATED",
        steps=(),
    )

    assert execution_skip_reason(
        case,
        run_seeded=True,
        allow_mutation=False,
        allow_destructive=False,
        mutation_customer_query=None,
        customer_preflight_verified=False,
    ) is None
