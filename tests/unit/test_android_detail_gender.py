from __future__ import annotations

from unittest.mock import Mock, call

import pytest

from scripts.prepare_excel import CURRENT_STEPS
from yanjia_automation.excel.safety import (
    CUSTOMER_MUTATION_CASE_IDS,
    CUSTOMER_MUTATION_TAIL_ACTIONS,
)
from yanjia_automation.excel.single_sheet import (
    SingleSheetMigrationError,
    _detail_validation_steps,
)
from yanjia_automation.screens.customer import CustomerEditScreen


@pytest.mark.parametrize(
    ("case_id", "gender", "option_id"),
    [
        ("TC-DETAIL-004", "男", "customer_edit_sex_man_tv"),
        ("TC-DETAIL-005", "女", "customer_edit_sex_female_tv"),
    ],
)
def test_detail_gender_steps_use_native_controls_and_guarded_restore(
    case_id: str,
    gender: str,
    option_id: str,
) -> None:
    steps = CURRENT_STEPS[case_id]

    assert len(steps) == 15
    assert case_id in CUSTOMER_MUTATION_CASE_IDS
    assert CUSTOMER_MUTATION_TAIL_ACTIONS[case_id] == (
        "click",
        "click",
        "click",
        "click",
        "assert",
    )
    assert steps[10][2].endswith("customer_edit_sex_tv")
    assert steps[11][2].endswith(option_id)
    assert steps[12][2].endswith("customer_edit_save_tv")
    assert steps[13][2].endswith("customer_detail_edit_tv")
    assert steps[14][2].endswith("customer_edit_sex_tv")
    assert steps[14][4:6] == ("text_equals", gender)
    assert "强制恢复原姓名、手机号和性别" in str(steps[14][10])


@pytest.mark.parametrize(
    ("case_id", "gender", "option_id"),
    [
        ("TC-DETAIL-004", "男", "customer_edit_sex_man_tv"),
        ("TC-DETAIL-005", "女", "customer_edit_sex_female_tv"),
    ],
)
def test_single_sheet_gender_steps_match_canonical_contract(
    case_id: str,
    gender: str,
    option_id: str,
) -> None:
    steps = _detail_validation_steps(
        row=46,
        case_id=case_id,
        input_data=gender,
        timeout=10,
    )

    assert len(steps) == 15
    assert [step.action for step in steps[-5:]] == [
        "click",
        "click",
        "click",
        "click",
        "assert",
    ]
    assert steps[-4].locator is not None and steps[-4].locator.endswith(option_id)
    assert steps[-1].assertion == "text_equals"
    assert steps[-1].expected == gender


def test_single_sheet_gender_rejects_mismatched_case_input() -> None:
    with pytest.raises(SingleSheetMigrationError, match="必须为男"):
        _detail_validation_steps(
            row=46,
            case_id="TC-DETAIL-004",
            input_data="女",
            timeout=10,
        )


def test_customer_edit_screen_selects_only_supported_native_gender() -> None:
    screen = CustomerEditScreen(Mock(), timeout=1)
    screen.click = Mock()
    screen.gender_value = Mock(return_value="女")

    screen.select_gender("女")

    assert screen.click.call_args_list == [call(screen.gender), call(screen.gender_female)]
    with pytest.raises(ValueError, match="仅支持男或女"):
        screen.select_gender("保密")
