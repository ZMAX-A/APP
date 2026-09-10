from __future__ import annotations

from collections import Counter
from collections.abc import Callable
from dataclasses import dataclass, field
from types import TracebackType
from typing import Literal, Self

from appium.webdriver.common.appiumby import AppiumBy
from appium.webdriver.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait

from yanjia_automation.config import Settings
from yanjia_automation.flows.customer_mutation import (
    require_dedicated_customer_target,
    require_unique_customer_match,
)
from yanjia_automation.flows.recovery import restart_to_home
from yanjia_automation.screens.base import resource_id
from yanjia_automation.screens.customer import CustomerDetailScreen, CustomerListScreen

RemarkScope = Literal["image", "consultation"]

_REMARK_CONTENT = resource_id("a_records_remark_list_content_tv")
_REMARK_DELETE = resource_id("a_records_remark_list_delete_ifv")
_DELETE_CONFIRM = resource_id("cover_prompt_right_tv")
_REMARK_CLOSE = resource_id("records_remark_close_ifv")


class RemarkMutationSafetyError(RuntimeError):
    """Raised before a remark write or when cleanup cannot be proven safe."""


class RemarkRestoreError(RuntimeError):
    """Raised when the original remark collection cannot be proven restored."""


@dataclass(frozen=True, repr=False)
class RemarkSnapshot:
    """Remark state held in memory only; content must never be logged."""

    texts: tuple[str, ...] = field(repr=False)
    record_count: int
    delete_count: int

    def __repr__(self) -> str:
        return (
            "RemarkSnapshot(texts=<redacted>, "
            f"record_count={self.record_count}, delete_count={self.delete_count})"
        )


class RemarkMutationSession:
    """Run one remark write and force exact collection restoration on exit."""

    def __init__(
        self,
        snapshot: RemarkSnapshot,
        *,
        restore: Callable[[RemarkSnapshot], None],
        verify: Callable[[RemarkSnapshot], bool],
    ) -> None:
        self.snapshot = snapshot
        self._restore = restore
        self._verify = verify
        self._entered = False
        self._mutation_started = False
        self._closed = False

    def __enter__(self) -> Self:
        if self._entered or self._closed:
            raise RemarkMutationSafetyError("Remark mutation session cannot be reused.")
        self._entered = True
        return self

    def run[ResultT](self, mutation: Callable[[], ResultT]) -> ResultT:
        if not self._entered or self._closed:
            raise RemarkMutationSafetyError(
                "Remark mutation must run inside its restoration context."
            )
        if self._mutation_started:
            raise RemarkMutationSafetyError(
                "A remark mutation session permits exactly one mutation callback."
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
            raise RemarkRestoreError(
                "Remark restoration failed; stop further mutation tests."
            ) from error
        if not restored:
            raise RemarkRestoreError(
                "Remark restoration could not be verified; stop further mutation tests."
            )
        return False


class DedicatedRemarkMutationFlow:
    """Open one dedicated image/consultation remark and restore by collection diff."""

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

    def restoration_session(self, scope: RemarkScope) -> RemarkMutationSession:
        snapshot = self.open_and_snapshot(scope)
        return RemarkMutationSession(
            snapshot,
            restore=self._restore_remarks,
            verify=self._verify_remarks,
        )

    def open_and_snapshot(self, scope: RemarkScope) -> RemarkSnapshot:
        if not self.settings.customer_preflight_verified:
            raise RemarkMutationSafetyError(
                "Dedicated remark mutation requires customer preflight verification."
            )
        target = require_dedicated_customer_target(self.settings)
        home = restart_to_home(self.driver, self.settings, timeout=self.timeout)
        home.open_customer_records()
        self.customer_list.wait_loaded(timeout=self.timeout)
        matches = self.customer_list.search_customers(
            target.query,
            timeout=self.timeout,
            expected_count=1,
        )
        require_unique_customer_match(matches).click()
        self.customer_detail.wait_loaded(timeout=self.timeout)
        entry = (
            resource_id("a_records_detail_all_remark_ifv")
            if scope == "image"
            else resource_id("a_consultation_result_remark_ifv")
        )
        entries = self.customer_detail.find_all(entry)
        if not entries:
            raise RemarkMutationSafetyError(
                f"Dedicated customer requires at least one {scope} remark entry."
            )
        entries[0].click()
        WebDriverWait(self.driver, self.timeout).until(
            lambda current: current.current_activity.endswith(".RecordsRemarkActivity")
        )
        return self._snapshot()

    def _snapshot(self) -> RemarkSnapshot:
        texts = tuple(
            str(element.get_attribute("text") or "")
            for element in self.driver.find_elements(*_REMARK_CONTENT)
        )
        return RemarkSnapshot(
            texts=texts,
            record_count=len(texts),
            delete_count=len(self.driver.find_elements(*_REMARK_DELETE)),
        )

    def delete_created_remark(self, snapshot: RemarkSnapshot) -> None:
        current = self._snapshot()
        if current == snapshot:
            raise RemarkMutationSafetyError("No generated remark is available to delete.")
        created_display_text = self._created_display_text(snapshot, current)
        self._delete_exact_display_remark(created_display_text)
        WebDriverWait(self.driver, self.timeout).until(
            lambda _: self._snapshot() == snapshot
        )

    def _restore_remarks(self, snapshot: RemarkSnapshot) -> None:
        current = self._snapshot()
        if current != snapshot:
            created_display_text = self._created_display_text(snapshot, current)
            self._delete_exact_display_remark(created_display_text)
            WebDriverWait(self.driver, self.timeout).until(
                lambda _: self._snapshot() == snapshot
            )
        self._close_remark()

    @staticmethod
    def _created_display_text(
        snapshot: RemarkSnapshot,
        current: RemarkSnapshot,
    ) -> str:
        original = Counter(snapshot.texts)
        observed = Counter(current.texts)
        removed = original - observed
        additions = observed - original
        if removed:
            raise RemarkMutationSafetyError(
                "Refusing cleanup because an original remark is missing."
            )
        if (
            sum(additions.values()) != 1
            or current.record_count != snapshot.record_count + 1
            or current.delete_count != snapshot.delete_count + 1
        ):
            raise RemarkMutationSafetyError(
                "Refusing cleanup because the remark delta is not exactly one addition."
            )
        return next(additions.elements())

    def _delete_exact_display_remark(self, display_text: str) -> None:
        matching_texts = [
            element
            for element in self.driver.find_elements(*_REMARK_CONTENT)
            if str(element.get_attribute("text") or "") == display_text
        ]
        if len(matching_texts) != 1:
            raise RemarkMutationSafetyError(
                "Generated remark display text is not uniquely identifiable."
            )
        row_delete_xpath = (
            f"//*[@resource-id={_xpath_literal(_REMARK_CONTENT[1])} "
            f"and @text={_xpath_literal(display_text)}]"
            f"/..//*[@resource-id={_xpath_literal(_REMARK_DELETE[1])}]"
        )
        row_delete_buttons = self.driver.find_elements(
            AppiumBy.XPATH,
            row_delete_xpath,
        )
        if len(row_delete_buttons) != 1:
            raise RemarkMutationSafetyError(
                "Generated remark row does not contain exactly one delete control."
            )
        row_delete_buttons[0].click()
        WebDriverWait(self.driver, self.timeout).until(
            lambda current: current.find_elements(*_DELETE_CONFIRM)
        )
        confirms = self.driver.find_elements(*_DELETE_CONFIRM)
        if len(confirms) != 1:
            raise RemarkMutationSafetyError(
                "Remark delete confirmation is not unique."
            )
        confirms[0].click()

    def _verify_remarks(self, snapshot: RemarkSnapshot) -> bool:
        if self.driver.current_activity.endswith(".RecordsRemarkActivity"):
            return self._snapshot() == snapshot
        return self.driver.current_activity.endswith(".CustomerDetailActivity")

    def _close_remark(self) -> None:
        close_buttons = self.driver.find_elements(*_REMARK_CLOSE)
        if len(close_buttons) != 1:
            raise RemarkMutationSafetyError("Remark close control is not unique.")
        close_buttons[0].click()
        WebDriverWait(self.driver, self.timeout).until(
            lambda current: current.current_activity.endswith(".CustomerDetailActivity")
        )


def _xpath_literal(value: str) -> str:
    if '"' not in value:
        return f'"{value}"'
    if "'" not in value:
        return f"'{value}'"
    pieces = value.split('"')
    return "concat(" + ", '\"', ".join(f'"{piece}"' for piece in pieces) + ")"
