from __future__ import annotations

from datetime import date
from unittest.mock import Mock

from selenium.webdriver.remote.command import Command

from yanjia_automation.date_picker import (
    WHEEL_BATCH_SIZE,
    WHEEL_STEP_DURATION_MS,
    swipe_number_picker_steps,
)
from yanjia_automation.excel.runner import ExcelCaseRunner

WHEEL_RECT = {"x": 10, "y": 20, "width": 100, "height": 400}


def _performed_actions(driver: Mock) -> list[list[dict[str, object]]]:
    return [call.args[1]["actions"][0]["actions"] for call in driver.execute.call_args_list]


def _end_moves(actions: list[dict[str, object]]) -> list[dict[str, object]]:
    return [
        action
        for action in actions
        if action["type"] == "pointerMove" and action["duration"] == WHEEL_STEP_DURATION_MS
    ]


def test_number_picker_skips_driver_calls_when_already_at_target() -> None:
    driver = Mock()

    swipe_number_picker_steps(driver, WHEEL_RECT, 0)

    driver.execute.assert_not_called()


def test_number_picker_batches_large_forward_delta_and_swipes_up() -> None:
    driver = Mock()
    steps = WHEEL_BATCH_SIZE * 2 + 1

    swipe_number_picker_steps(driver, WHEEL_RECT, steps)

    assert driver.execute.call_count == 3
    assert all(call.args[0] == Command.W3C_ACTIONS for call in driver.execute.call_args_list)
    batches = _performed_actions(driver)
    assert [len(_end_moves(actions)) for actions in batches] == [
        WHEEL_BATCH_SIZE,
        WHEEL_BATCH_SIZE,
        1,
    ]
    assert {move["x"] for actions in batches for move in _end_moves(actions)} == {60}
    assert {move["y"] for actions in batches for move in _end_moves(actions)} == {128}


def test_number_picker_swipes_down_for_backward_delta() -> None:
    driver = Mock()

    swipe_number_picker_steps(driver, WHEEL_RECT, -2)

    actions = _performed_actions(driver)
    assert len(actions) == 1
    assert [move["y"] for move in _end_moves(actions[0])] == [312, 312]


def test_excel_runner_locates_wheel_once_for_eleven_year_delta() -> None:
    driver = Mock()
    start_wheel = Mock()
    end_wheel = Mock()
    end_wheel.rect = WHEEL_RECT
    driver.find_elements.return_value = [start_wheel, end_wheel]
    runner = ExcelCaseRunner(driver, Mock(), Mock())

    runner._scroll_date_wheel("app:id/year", 1, 2026, 2015)

    driver.find_elements.assert_called_once_with("id", "app:id/year")
    driver.execute.assert_called_once()
    actions = _performed_actions(driver)
    assert len(_end_moves(actions[0])) == 11


def test_excel_runner_does_not_locate_unchanged_wheel() -> None:
    driver = Mock()
    runner = ExcelCaseRunner(driver, Mock(), Mock())

    runner._scroll_date_wheel("app:id/month", 0, 2, 2)

    driver.find_elements.assert_not_called()
    driver.execute.assert_not_called()


def test_date_endpoint_recalculates_day_after_range_constraint_clamps_month() -> None:
    runner = ExcelCaseRunner(Mock(), Mock(), Mock())
    scroll = Mock()
    runner._scroll_date_wheel = scroll

    runner._select_date_endpoint(
        ("year", "month", "day"),
        1,
        date(2026, 9, 2),
        date(2015, 2, 13),
        minimum=date(2015, 2, 4),
    )

    assert scroll.call_args_list == [
        (("year", 1, 2026, 2015),),
        (("month", 1, 9, 2),),
        (("day", 1, 4, 13),),
    ]


def test_date_endpoint_recalculates_day_for_shorter_target_month() -> None:
    runner = ExcelCaseRunner(Mock(), Mock(), Mock())
    scroll = Mock()
    runner._scroll_date_wheel = scroll

    runner._select_date_endpoint(
        ("year", "month", "day"),
        0,
        date(2024, 1, 31),
        date(2024, 2, 29),
    )

    assert scroll.call_args_list[-1].args == ("day", 0, 29, 29)
