from __future__ import annotations

from unittest.mock import Mock, call

import pytest

from scripts.prepare_excel import CURRENT_STEPS
from yanjia_automation.excel.models import ExcelCase
from yanjia_automation.excel.safety import (
    CUSTOMER_MUTATION_TAIL_ACTIONS,
    execution_skip_reason,
)
from yanjia_automation.excel.single_sheet import (
    SingleSheetMigrationError,
    _detail_validation_steps,
)
from yanjia_automation.screens.customer import CustomerEditScreen


def test_marital_steps_use_guarded_write_reopen_and_restore_contract() -> None:
    steps = CURRENT_STEPS["TC-DETAIL-015"]

    assert len(steps) == 15
    assert tuple(step[1] for step in steps[-5:]) == CUSTOMER_MUTATION_TAIL_ACTIONS[
        "TC-DETAIL-015"
    ]
    assert steps[10][2].endswith("customer_edit_marital_tv")
    assert steps[11][2].endswith("customer_edit_marital_secret_tv")
    assert steps[-1][4:6] == ("text_equals", "保密")
    assert "恢复运行前姓名、手机号、性别、生日、邮箱、婚姻状态、地址和个人备注" in str(
        steps[-1][10]
    )


def test_empty_tag_steps_validate_and_cancel_without_creating_data() -> None:
    steps = CURRENT_STEPS["TC-DETAIL-017"]

    assert len(steps) == 18
    assert steps[13][2].endswith("records_remark_tag_save_tv")
    assert steps[14][4:6] == ("page_source_contains", "请输入标签内容")
    assert steps[15][2].endswith("records_remark_tag_cancel_tv")
    assert not any(step[1] == "input" for step in steps[8:])
    assert "不创建标签" in str(steps[13][10])


@pytest.mark.parametrize(
    ("case_id", "input_data", "expected_length", "tail_actions"),
    [
        (
            "TC-DETAIL-015",
            "保密",
            15,
            ("click", "click", "click", "click", "assert"),
        ),
        ("TC-DETAIL-017", None, 18, None),
    ],
)
def test_single_sheet_boundary_steps_match_canonical_contract(
    case_id: str,
    input_data: str | None,
    expected_length: int,
    tail_actions: tuple[str, ...] | None,
) -> None:
    steps = _detail_validation_steps(
        row=57,
        case_id=case_id,
        input_data=input_data,
        timeout=10,
    )

    assert len(steps) == expected_length
    if tail_actions is not None:
        assert tuple(step.action for step in steps[-len(tail_actions) :]) == tail_actions
    else:
        assert steps[-4].assertion == "page_source_contains"
        assert steps[-4].expected == "请输入标签内容"


def test_single_sheet_marital_rejects_non_contract_value() -> None:
    with pytest.raises(SingleSheetMigrationError, match="必须为保密"):
        _detail_validation_steps(
            row=57,
            case_id="TC-DETAIL-015",
            input_data="已婚",
            timeout=10,
        )


def test_customer_edit_screen_selects_supported_marital_state() -> None:
    screen = CustomerEditScreen(Mock(), timeout=1)
    screen.click = Mock()
    screen.marital_value = Mock(return_value="保密")

    screen.select_marital("保密")

    assert screen.click.call_args_list == [
        call(screen.marital),
        call(screen.marital_secret),
    ]


def test_customer_edit_screen_rejects_unknown_marital_state() -> None:
    screen = CustomerEditScreen(Mock(), timeout=1)

    with pytest.raises(ValueError, match="保密、未婚或已婚"):
        screen.select_marital("未知")


def test_empty_tag_case_requires_dedicated_query_without_mutation_authorization() -> None:
    case = ExcelCase(
        source_row=59,
        case_id="TC-DETAIL-017",
        module="顾客详情",
        scenario="标签为空",
        test_point="校验空标签提示",
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
