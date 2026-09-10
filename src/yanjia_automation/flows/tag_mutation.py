from __future__ import annotations

from collections import Counter
from collections.abc import Callable
from dataclasses import dataclass, field
from math import hypot
from types import TracebackType
from typing import Literal, Self

from appium.webdriver.webdriver import WebDriver
from appium.webdriver.webelement import WebElement
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait

from yanjia_automation.config import Settings
from yanjia_automation.flows.customer_mutation import (
    require_dedicated_customer_target,
    require_unique_customer_match,
)
from yanjia_automation.flows.recovery import restart_to_home
from yanjia_automation.screens.base import resource_id
from yanjia_automation.screens.customer import CustomerDetailScreen, CustomerListScreen

TagScope = Literal["image", "consultation"]

_TAG_CONTAINER = resource_id("a_records_remark_tag_cl")
_TAG_TEXT = resource_id("a_records_remark_tag_tv")
_TAG_DELETE = resource_id("a_records_remark_tag_delete_ifv")
_TAG_DIALOG = resource_id("records_remark_tag_cl")
_TAG_CANCEL = resource_id("records_remark_tag_cancel_tv")
_DELETE_CONFIRM = resource_id("cover_prompt_right_tv")
_REMARK_CLOSE = resource_id("records_remark_close_ifv")


class TagMutationSafetyError(RuntimeError):
    """Raised before a tag write or when cleanup cannot be proven safe."""


class TagRestoreError(RuntimeError):
    """Raised when the original tag collection cannot be proven restored."""


@dataclass(frozen=True, repr=False)
class TagSnapshot:
    """Tag state held in memory only; existing tag text must never be logged."""

    texts: tuple[str, ...] = field(repr=False)
    container_count: int
    delete_count: int

    def __repr__(self) -> str:
        return (
            "TagSnapshot(texts=<redacted>, "
            f"container_count={self.container_count}, delete_count={self.delete_count})"
        )


class TagMutationSession:
    """Run one tag write and force exact collection restoration on exit."""

    def __init__(
        self,
        snapshot: TagSnapshot,
        *,
        restore: Callable[[TagSnapshot], None],
        verify: Callable[[TagSnapshot], bool],
    ) -> None:
        self.snapshot = snapshot
        self._restore = restore
        self._verify = verify
        self._entered = False
        self._mutation_started = False
        self._closed = False

    def __enter__(self) -> Self:
        if self._entered or self._closed:
            raise TagMutationSafetyError("Tag mutation session cannot be reused.")
        self._entered = True
        return self

    def run[ResultT](self, mutation: Callable[[], ResultT]) -> ResultT:
        if not self._entered or self._closed:
            raise TagMutationSafetyError(
                "Tag mutation must run inside its restoration context."
            )
        if self._mutation_started:
            raise TagMutationSafetyError(
                "A tag mutation session permits exactly one mutation callback."
            )
        self._mutation_started = True
        return mutation()

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> bool:
        del exc_type, exc, traceback
        self._closed = True
        if not self._mutation_started:
            return False
        try:
            self._restore(self.snapshot)
            restored = self._verify(self.snapshot)
        except BaseException as error:
            raise TagRestoreError(
                "Tag restoration failed; stop further mutation tests."
            ) from error
        if not restored:
            raise TagRestoreError(
                "Tag restoration could not be verified; stop further mutation tests."
            )
        return False


class DedicatedTagMutationFlow:
    """Open one dedicated customer's remark tags and restore by collection diff."""

    def __init__(
        self,
        driver: WebDriver,
        settings: Settings,
        *,
        timeout: float = 30,
    ) -> None:
        self.driver = driver
        self.settings = settings
        self.timeout = timeout
        self.customer_list = CustomerListScreen(driver, timeout)
        self.customer_detail = CustomerDetailScreen(driver, timeout)

    def restoration_session(self, scope: TagScope) -> TagMutationSession:
        snapshot = self.open_and_snapshot(scope)
        return TagMutationSession(
            snapshot,
            restore=self._restore_tags,
            verify=self._verify_tags,
        )

    def open_and_snapshot(self, scope: TagScope) -> TagSnapshot:
        if not self.settings.customer_preflight_verified:
            raise TagMutationSafetyError(
                "Dedicated tag mutation requires customer preflight verification."
            )
        target = require_dedicated_customer_target(self.settings)
        home = restart_to_home(self.driver, self.settings, timeout=self.timeout)
        home.open_customer_records()
        self.customer_list.wait_loaded(timeout=self.timeout)
        self.customer_list.search_customers(
            target.query,
            timeout=self.timeout,
            expected_count=1,
        )
        try:
            WebDriverWait(self.driver, self.timeout).until(
                lambda _: self._has_exact_customer_identifier(target.query)
            )
        except TimeoutException as error:
            raise TagMutationSafetyError(
                "Dedicated customer search did not resolve the exact configured identifier."
            ) from error
        matches = self.customer_list.find_all(self.customer_list.cards)
        require_unique_customer_match(matches).click()
        self.customer_detail.wait_loaded(timeout=self.timeout)

        entry = (
            resource_id("a_records_detail_all_remark_ifv")
            if scope == "image"
            else resource_id("a_consultation_result_remark_ifv")
        )
        try:
            entries = WebDriverWait(self.driver, self.timeout).until(
                lambda _: self.customer_detail.find_all(entry) or False
            )
        except TimeoutException as error:
            raise TagMutationSafetyError(
                f"Dedicated customer requires at least one {scope} remark entry."
            ) from error
        entries[0].click()
        WebDriverWait(self.driver, self.timeout).until(
            lambda current: current.current_activity.endswith(".RecordsRemarkActivity")
        )
        return self._snapshot()

    def _has_exact_customer_identifier(self, expected: str) -> bool:
        identifiers = self.customer_list.find_all(self.customer_list.card_identifier)
        return len(identifiers) == 1 and str(
            identifiers[0].get_attribute("text") or ""
        ) == expected

    def _snapshot(self) -> TagSnapshot:
        return TagSnapshot(
            texts=tuple(
                str(element.get_attribute("text") or "")
                for element in self.driver.find_elements(*_TAG_TEXT)
            ),
            container_count=len(self.driver.find_elements(*_TAG_CONTAINER)),
            delete_count=len(self.driver.find_elements(*_TAG_DELETE)),
        )

    def _restore_tags(self, snapshot: TagSnapshot) -> None:
        self._dismiss_add_dialog()
        current = self._snapshot()
        if current == snapshot:
            self._close_remark()
            return
        created_display_text = self._created_display_text(snapshot, current)
        self._delete_exact_display_tag(created_display_text)
        WebDriverWait(self.driver, self.timeout).until(
            lambda _: self._snapshot() == snapshot
        )
        self._close_remark()

    def delete_created_tag(self, snapshot: TagSnapshot) -> None:
        """Delete only the unique tag added after ``snapshot`` and verify restore."""

        self._dismiss_add_dialog()
        current = self._snapshot()
        if current == snapshot:
            raise TagMutationSafetyError("No generated tag is available to delete.")
        created_display_text = self._created_display_text(snapshot, current)
        self._delete_exact_display_tag(created_display_text)
        WebDriverWait(self.driver, self.timeout).until(
            lambda _: self._snapshot() == snapshot
        )

    @staticmethod
    def _created_display_text(snapshot: TagSnapshot, current: TagSnapshot) -> str:
        original = Counter(snapshot.texts)
        observed = Counter(current.texts)
        removed = original - observed
        additions = observed - original
        if removed:
            raise TagMutationSafetyError(
                "Refusing cleanup because an original tag is missing."
            )
        if (
            sum(additions.values()) != 1
            or current.container_count != snapshot.container_count + 1
            or current.delete_count != snapshot.delete_count + 1
        ):
            raise TagMutationSafetyError(
                "Refusing cleanup because the tag delta is not exactly one addition."
            )
        return next(additions.elements())

    def _verify_tags(self, snapshot: TagSnapshot) -> bool:
        if self.driver.current_activity.endswith(".RecordsRemarkActivity"):
            return self._snapshot() == snapshot
        return self.driver.current_activity.endswith(".CustomerDetailActivity")

    def _dismiss_add_dialog(self) -> None:
        dialogs = self.driver.find_elements(*_TAG_DIALOG)
        if not dialogs:
            return
        cancel = self.driver.find_elements(*_TAG_CANCEL)
        if len(cancel) != 1:
            raise TagMutationSafetyError(
                "Tag dialog is open but its cancel control is not unique."
            )
        cancel[0].click()
        WebDriverWait(self.driver, self.timeout).until(
            lambda current: not current.find_elements(*_TAG_DIALOG)
        )

    def _delete_exact_display_tag(self, display_text: str) -> None:
        matching_texts = [
            element
            for element in self.driver.find_elements(*_TAG_TEXT)
            if str(element.get_attribute("text") or "") == display_text
        ]
        if len(matching_texts) != 1:
            raise TagMutationSafetyError(
                "Generated tag display text is not uniquely identifiable."
            )
        delete_buttons = self.driver.find_elements(*_TAG_DELETE)
        if not delete_buttons:
            raise TagMutationSafetyError("Generated tag has no delete control.")
        text_x, text_y = _center(matching_texts[0])
        delete_button = min(
            delete_buttons,
            key=lambda element: hypot(
                _center(element)[0] - text_x,
                _center(element)[1] - text_y,
            ),
        )
        delete_button.click()
        WebDriverWait(self.driver, self.timeout).until(
            lambda current: current.find_elements(*_DELETE_CONFIRM)
        )
        confirms = self.driver.find_elements(*_DELETE_CONFIRM)
        if len(confirms) != 1:
            raise TagMutationSafetyError("Tag delete confirmation is not unique.")
        confirms[0].click()

    def _close_remark(self) -> None:
        close_buttons = self.driver.find_elements(*_REMARK_CLOSE)
        if len(close_buttons) != 1:
            raise TagMutationSafetyError("Remark close control is not unique.")
        close_buttons[0].click()
        WebDriverWait(self.driver, self.timeout).until(
            lambda current: current.current_activity.endswith(".CustomerDetailActivity")
        )


def _center(element: WebElement) -> tuple[float, float]:
    rect = element.rect
    return (
        float(rect["x"]) + float(rect["width"]) / 2,
        float(rect["y"]) + float(rect["height"]) / 2,
    )
