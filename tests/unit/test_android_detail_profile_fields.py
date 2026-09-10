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


def test_birthday_steps_reconcile_source_contract_and_use_native_wheels() -> None:
    steps = CURRENT_STEPS["TC-DETAIL-007"]

    assert len(steps) == 14
    assert CUSTOMER_MUTATION_TAIL_ACTIONS["TC-DETAIL-007"] == (
        "select_birthday",
        "click",
        "click",
        "assert",
    )
    assert steps[10][1] == "select_birthday"
    assert steps[10][3] == "2022-08-04"
    assert steps[-1][2].endswith("customer_edit_birthday_tv")
    assert steps[-1][4:6] == ("text_equals", "2022-08-04")
    assert "03日与输入/验证点冲突" in str(steps[10][10])


@pytest.mark.parametrize(
    ("case_id", "action", "input_value", "assertion", "expected"),
    [
        ("TC-DETAIL-008", "clear", None, "attribute_equals", "text|"),
        ("TC-DETAIL-009", "input", "……&&*……*&", "text_equals", "……&&*……*&"),
        ("TC-DETAIL-010", "input", "${SPACE}", "attribute_equals", "text|"),
    ],
)
def test_address_steps_use_guarded_write_reopen_and_restore_contract(
    case_id: str,
    action: str,
    input_value: str | None,
    assertion: str,
    expected: str,
) -> None:
    steps = CURRENT_STEPS[case_id]

    assert len(steps) == 14
    assert case_id in CUSTOMER_MUTATION_CASE_IDS
    assert tuple(step[1] for step in steps[-4:]) == CUSTOMER_MUTATION_TAIL_ACTIONS[
        case_id
    ]
    assert steps[10][1] == action
    assert steps[10][2].endswith("customer_edit_address_et")
    assert steps[10][3] == input_value
    assert steps[-1][4:6] == (assertion, expected)
    assert "恢复运行前姓名、手机号、性别、生日、邮箱、婚姻状态、地址和个人备注" in str(
        steps[-1][10]
    )


@pytest.mark.parametrize(
    ("case_id", "input_data", "tail_actions", "assertion", "expected"),
    [
        (
            "TC-DETAIL-007",
            "2022-08-04",
            ("select_birthday", "click", "click", "assert"),
            "text_equals",
            "2022-08-04",
        ),
        (
            "TC-DETAIL-008",
            None,
            ("clear", "click", "click", "assert"),
            "attribute_equals",
            "text|",
        ),
        (
            "TC-DETAIL-009",
            "……&&*……*&",
            ("input", "click", "click", "assert"),
            "text_equals",
            "……&&*……*&",
        ),
        (
            "TC-DETAIL-010",
            None,
            ("input", "click", "click", "assert"),
            "attribute_equals",
            "text|",
        ),
    ],
)
def test_single_sheet_profile_steps_match_canonical_tail(
    case_id: str,
    input_data: str | None,
    tail_actions: tuple[str, ...],
    assertion: str,
    expected: str,
) -> None:
    steps = _detail_validation_steps(
        row=49,
        case_id=case_id,
        input_data=input_data,
        timeout=10,
    )

    assert len(steps) == 14
    assert tuple(step.action for step in steps[-4:]) == tail_actions
    assert steps[-1].assertion == assertion
    assert steps[-1].expected == expected


def test_single_sheet_birthday_rejects_the_conflicting_old_step_date() -> None:
    with pytest.raises(SingleSheetMigrationError, match="统一为 2022-08-04"):
        _detail_validation_steps(
            row=49,
            case_id="TC-DETAIL-007",
            input_data="2022-08-03",
            timeout=10,
        )


def test_customer_edit_screen_selects_birthday_with_single_native_wheel() -> None:
    driver = Mock()
    wheel = Mock()
    wheel.rect = {"x": 10, "y": 20, "width": 100, "height": 400}
    driver.find_elements.return_value = [wheel]
    screen = CustomerEditScreen(driver, timeout=1)
    screen.click = Mock()
    screen.birthday_value = Mock(side_effect=["2022-08-02", "2022-08-04"])

    screen.select_birthday("2022-08-04")

    assert screen.click.call_args_list == [
        call(screen.birthday),
        call(screen.birthday_confirm),
    ]
    driver.find_elements.assert_called_once_with(*screen.birthday_day_wheels)
    driver.execute.assert_called_once()
    payload = driver.execute.call_args.args[1]
    pointer_moves = payload["actions"][0]["actions"]
    end_moves = [
        action
        for action in pointer_moves
        if action["type"] == "pointerMove" and action["duration"] > 0
    ]
    assert len(end_moves) == 2


def test_customer_edit_screen_recalculates_day_after_shorter_month() -> None:
    driver = Mock()
    screen = CustomerEditScreen(driver, timeout=1)
    screen.click = Mock()
    screen.birthday_value = Mock(side_effect=["2024-01-31", "2024-02-29"])
    scroll = Mock()
    screen._scroll_date_wheel = scroll

    screen.select_birthday("2024-02-29")

    assert [item.args for item in scroll.call_args_list] == [
        (screen.birthday_year_wheels, 2024, 2024),
        (screen.birthday_month_wheels, 1, 2),
        (screen.birthday_day_wheels, 29, 29),
    ]


def test_customer_edit_screen_sets_address_without_sending_empty_text() -> None:
    driver = Mock()
    field = Mock()
    screen = CustomerEditScreen(driver, timeout=1)
    screen.find = Mock(return_value=field)

    screen.set_address("")

    field.clear.assert_called_once_with()
    field.send_keys.assert_not_called()
