from __future__ import annotations

from unittest.mock import Mock, call

import pytest

from scripts.prepare_excel import CURRENT_STEPS
from yanjia_automation.excel.safety import CUSTOMER_MUTATION_TAIL_ACTIONS
from yanjia_automation.excel.single_sheet import (
    SingleSheetMigrationError,
    _detail_validation_steps,
)
from yanjia_automation.screens.customer import CustomerEditScreen


@pytest.mark.parametrize(
    ("case_id", "action", "input_value", "assertion", "expected"),
    [
        ("TC-DETAIL-013", "clear", None, "attribute_equals", "text|"),
        ("TC-DETAIL-014", "input", "……&&*……*&", "text_equals", "……&&*……*&"),
    ],
)
def test_profile_remark_steps_scroll_write_reopen_and_restore(
    case_id: str,
    action: str,
    input_value: str | None,
    assertion: str,
    expected: str,
) -> None:
    steps = CURRENT_STEPS[case_id]

    assert len(steps) == 16
    assert tuple(step[1] for step in steps[-6:]) == CUSTOMER_MUTATION_TAIL_ACTIONS[
        case_id
    ]
    assert steps[10][1] == "scroll_profile_remark"
    assert steps[11][1] == action
    assert steps[11][2].endswith("customer_edit_remark_et")
    assert steps[11][3] == input_value
    assert steps[14][1] == "scroll_profile_remark"
    assert steps[15][4:6] == (assertion, expected)
    assert "八项资料" in str(steps[15][10])


@pytest.mark.parametrize(
    ("case_id", "input_data", "assertion", "expected"),
    [
        ("TC-DETAIL-013", None, "attribute_equals", "text|"),
        ("TC-DETAIL-014", "……&&*……*&", "text_equals", "……&&*……*&"),
    ],
)
def test_single_sheet_profile_remark_steps_match_canonical_contract(
    case_id: str,
    input_data: str | None,
    assertion: str,
    expected: str,
) -> None:
    steps = _detail_validation_steps(
        row=55,
        case_id=case_id,
        input_data=input_data,
        timeout=10,
    )

    assert len(steps) == 16
    assert tuple(step.action for step in steps[-6:]) == CUSTOMER_MUTATION_TAIL_ACTIONS[
        case_id
    ]
    assert steps[-1].assertion == assertion
    assert steps[-1].expected == expected


def test_single_sheet_special_remark_requires_contract_input() -> None:
    with pytest.raises(SingleSheetMigrationError, match="缺少个人备注特殊字符输入"):
        _detail_validation_steps(
            row=56,
            case_id="TC-DETAIL-014",
            input_data=None,
            timeout=10,
        )


def test_customer_edit_screen_scrolls_native_profile_form_to_remark() -> None:
    driver = Mock()
    scroll_view = Mock()
    scroll_view.id = "profile-scroll"
    field = Mock()
    field.is_displayed.return_value = True
    screen = CustomerEditScreen(driver, timeout=1)
    screen.find_all = Mock(side_effect=[[], [field]])
    screen.find = Mock(return_value=scroll_view)

    assert screen.scroll_to_remark() is field

    assert driver.execute_script.call_args_list == [
        call(
            "mobile: scrollGesture",
            {
                "elementId": "profile-scroll",
                "direction": "down",
                "percent": 0.75,
            },
        )
    ]


def test_customer_edit_screen_clears_remark_without_sending_empty_text() -> None:
    field = Mock()
    screen = CustomerEditScreen(Mock(), timeout=1)
    screen.scroll_to_remark = Mock(return_value=field)

    screen.set_remark("")

    field.clear.assert_called_once_with()
    field.send_keys.assert_not_called()
