from __future__ import annotations

from calendar import monthrange
from collections.abc import Mapping
from datetime import date
from typing import Any

from selenium.webdriver.common.actions import interaction
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.common.actions.pointer_input import PointerInput

WHEEL_BATCH_SIZE = 12
WHEEL_STEP_DURATION_MS = 400
# The custom wheel starts its snap animation after pointer-up. Keeping that
# animation inside the same W3C batch prevents adjacent steps from merging.
WHEEL_STEP_PAUSE_SECONDS = 0.35


def date_after_picker_change(
    current: date,
    *,
    year: int | None = None,
    month: int | None = None,
    minimum: date | None = None,
    maximum: date | None = None,
) -> date:
    updated_year = year if year is not None else current.year
    updated_month = month if month is not None else current.month
    updated_day = min(current.day, monthrange(updated_year, updated_month)[1])
    updated = date(updated_year, updated_month, updated_day)
    if minimum is not None and updated < minimum:
        updated = minimum
    if maximum is not None and updated > maximum:
        updated = maximum
    return updated


def swipe_number_picker_steps(
    driver: Any,
    rect: Mapping[str, int],
    steps: int,
) -> None:
    """Move a native number picker by exact rows with batched W3C touch actions."""
    if steps == 0:
        return

    center_x = round(rect["x"] + rect["width"] / 2)
    center_y = round(rect["y"] + rect["height"] / 2)
    distance = max(80, round(rect["height"] * 0.23))
    end_y = center_y - distance if steps > 0 else center_y + distance

    remaining = abs(steps)
    while remaining:
        batch_size = min(remaining, WHEEL_BATCH_SIZE)
        touch = PointerInput(interaction.POINTER_TOUCH, "date-picker-wheel")
        actions = ActionBuilder(driver, mouse=touch)
        source = actions.pointer_action.source
        for _ in range(batch_size):
            source.create_pointer_move(duration=0, x=center_x, y=center_y)
            source.create_pointer_down(button=0)
            source.create_pointer_move(
                duration=WHEEL_STEP_DURATION_MS,
                x=center_x,
                y=end_y,
            )
            source.create_pointer_up(button=0)
            source.create_pause(WHEEL_STEP_PAUSE_SECONDS)
        actions.perform()
        remaining -= batch_size
