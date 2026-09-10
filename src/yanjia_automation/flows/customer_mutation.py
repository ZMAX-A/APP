from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from datetime import date
from types import TracebackType
from typing import Self

from appium.webdriver.webdriver import WebDriver

from yanjia_automation.config import Settings
from yanjia_automation.flows.recovery import restart_to_home
from yanjia_automation.screens.customer import (
    CustomerDetailScreen,
    CustomerEditScreen,
    CustomerListScreen,
)


class CustomerMutationSafetyError(RuntimeError):
    """Raised before a customer write when a safety precondition is missing."""


class CustomerRestoreError(RuntimeError):
    """Raised when the original customer state cannot be proven restored."""


@dataclass(frozen=True, repr=False)
class DedicatedCustomerTarget:
    """An isolated customer search value that must never be logged verbatim."""

    query: str = field(repr=False)

    def __repr__(self) -> str:
        return "DedicatedCustomerTarget(query=<redacted>)"


@dataclass(frozen=True, repr=False)
class CustomerProfileSnapshot:
    """Sensitive customer fields held in memory only for restoration."""

    name: str = field(repr=False)
    phone: str = field(repr=False)
    gender: str = field(repr=False)
    birthday: str = field(repr=False)
    email: str = field(repr=False)
    marital: str = field(repr=False)
    address: str = field(repr=False)
    remark: str = field(repr=False)

    def __repr__(self) -> str:
        return (
            "CustomerProfileSnapshot(name=<redacted>, phone=<redacted>, "
            "gender=<redacted>, birthday=<redacted>, email=<redacted>, "
            "marital=<redacted>, address=<redacted>, remark=<redacted>)"
        )


def require_dedicated_customer_target(settings: Settings) -> DedicatedCustomerTarget:
    """Fail closed unless all non-destructive mutation prerequisites are present."""

    if not settings.run_seeded:
        raise CustomerMutationSafetyError(
            "Dedicated customer mutation requires YANJIA_RUN_SEEDED=true."
        )
    if not settings.allow_mutation:
        raise CustomerMutationSafetyError(
            "Dedicated customer mutation requires YANJIA_ALLOW_MUTATION=true."
        )
    query = settings.mutation_customer_query
    if not query or not query.strip():
        raise CustomerMutationSafetyError(
            "Set YANJIA_MUTATION_CUSTOMER_QUERY to an isolated, restorable customer."
        )
    return DedicatedCustomerTarget(query=query.strip())


def require_unique_customer_match[MatchT](matches: Sequence[MatchT]) -> MatchT:
    """Return the only search result without exposing any result content in errors."""

    match_count = len(matches)
    if match_count != 1:
        raise CustomerMutationSafetyError(
            "Dedicated customer search must return exactly one match; "
            f"received {match_count}."
        )
    return matches[0]


class CustomerMutationSession[SnapshotT]:
    """Run one mutation and force a restore plus verification on context exit.

    The mutation is marked as started before its callback is invoked, so a callback
    that fails after a partial write still triggers restoration. If restoration or
    verification fails, CustomerRestoreError takes precedence over the mutation
    result because the customer is no longer known to be in its original state.
    """

    def __init__(
        self,
        snapshot: SnapshotT,
        *,
        restore: Callable[[SnapshotT], None],
        verify: Callable[[SnapshotT], bool],
    ) -> None:
        self.snapshot = snapshot
        self._restore = restore
        self._verify = verify
        self._entered = False
        self._mutation_started = False
        self._closed = False

    def __enter__(self) -> Self:
        if self._entered or self._closed:
            raise CustomerMutationSafetyError("Customer mutation session cannot be reused.")
        self._entered = True
        return self

    def run[ResultT](self, mutation: Callable[[], ResultT]) -> ResultT:
        if not self._entered or self._closed:
            raise CustomerMutationSafetyError(
                "Customer mutation must run inside its restoration context."
            )
        if self._mutation_started:
            raise CustomerMutationSafetyError(
                "A customer mutation session permits exactly one mutation callback."
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
            raise CustomerRestoreError(
                "Customer restoration failed; stop further mutation tests."
            ) from error

        if not restored:
            raise CustomerRestoreError(
                "Customer restoration could not be verified; stop further mutation tests."
            )
        return False


class DedicatedCustomerMutationFlow:
    """Open one dedicated customer and provide a verified restoration session."""

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
        self.customer_edit = CustomerEditScreen(driver, timeout)

    def open_and_snapshot(self) -> CustomerProfileSnapshot:
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
        self.customer_detail.open_editor()
        self.customer_edit.wait_loaded(timeout=self.timeout)

        name, phone = self.customer_edit.profile_values()
        gender = self.customer_edit.gender_value()
        birthday = self.customer_edit.birthday_value()
        email = self.customer_edit.email_value()
        marital = self.customer_edit.marital_value()
        address = self.customer_edit.address_value()
        remark = self.customer_edit.remark_value()
        # Reading the bottom-of-form remark leaves the native ScrollView at the
        # end. Reset it before handing control to a guarded mutation tail so
        # name/phone fields are immediately visible.
        self.customer_edit.scroll_to_top()
        if not name.strip() or not phone.strip() or gender not in {"男", "女"}:
            raise CustomerMutationSafetyError(
                "Dedicated customer snapshot requires name, phone, and supported gender."
            )
        try:
            date.fromisoformat(birthday)
        except ValueError as error:
            raise CustomerMutationSafetyError(
                "Dedicated customer snapshot requires an ISO birthday."
            ) from error
        if marital not in {"保密", "未婚", "已婚"}:
            raise CustomerMutationSafetyError(
                "Dedicated customer snapshot requires a supported marital value."
            )
        return CustomerProfileSnapshot(
            name=name,
            phone=phone,
            gender=gender,
            birthday=birthday,
            email=email,
            marital=marital,
            address=address,
            remark=remark,
        )

    def restoration_session(
        self,
    ) -> CustomerMutationSession[CustomerProfileSnapshot]:
        snapshot = self.open_and_snapshot()
        return CustomerMutationSession(
            snapshot,
            restore=self._restore_profile,
            verify=self._verify_profile,
        )

    def _restore_profile(self, snapshot: CustomerProfileSnapshot) -> None:
        self._ensure_editor_open(snapshot)
        self.customer_edit.set_profile(
            name=snapshot.name,
            phone=snapshot.phone,
            gender=snapshot.gender,
            birthday=snapshot.birthday,
            email=snapshot.email,
            marital=snapshot.marital,
            address=snapshot.address,
            remark=snapshot.remark,
        )
        self.customer_edit.save()
        self.customer_detail.wait_loaded(timeout=self.timeout)

    def _verify_profile(self, snapshot: CustomerProfileSnapshot) -> bool:
        self.customer_detail.wait_loaded(timeout=self.timeout)
        self.customer_detail.open_editor()
        self.customer_edit.wait_loaded(timeout=self.timeout)
        current_name, current_phone = self.customer_edit.profile_values()
        current_gender = self.customer_edit.gender_value()
        current_birthday = self.customer_edit.birthday_value()
        current_email = self.customer_edit.email_value()
        current_marital = self.customer_edit.marital_value()
        current_address = self.customer_edit.address_value()
        current_remark = self.customer_edit.remark_value()
        try:
            return (
                current_name == snapshot.name
                and current_phone == snapshot.phone
                and current_gender == snapshot.gender
                and current_birthday == snapshot.birthday
                and current_email == snapshot.email
                and current_marital == snapshot.marital
                and current_address == snapshot.address
                and current_remark == snapshot.remark
            )
        finally:
            self.customer_edit.close()
            self.customer_detail.wait_loaded(timeout=self.timeout)

    def _ensure_editor_open(self, snapshot: CustomerProfileSnapshot) -> None:
        if self.customer_edit.is_visible(self.customer_edit.root, timeout=0.5):
            return
        if not self.customer_detail.is_visible(self.customer_detail.root, timeout=0.5):
            if not self.customer_list.is_visible(self.customer_list.root, timeout=0.5):
                home = restart_to_home(self.driver, self.settings, timeout=self.timeout)
                home.open_customer_records()
                self.customer_list.wait_loaded(timeout=self.timeout)
            self._open_snapshot_customer_from_list(snapshot)
        else:
            self.customer_detail.wait_loaded(timeout=self.timeout)
        self.customer_detail.open_editor()
        self.customer_edit.wait_loaded(timeout=self.timeout)

    def _open_snapshot_customer_from_list(
        self, snapshot: CustomerProfileSnapshot
    ) -> None:
        visible_cards = self.customer_list.find_all(self.customer_list.cards)
        if len(visible_cards) == 1:
            visible_cards[0].click()
            self.customer_detail.wait_loaded(timeout=self.timeout)
            return

        for stable_query in (snapshot.phone, snapshot.name):
            matches = self.customer_list.search_customers(
                stable_query,
                timeout=self.timeout,
                expected_count=1,
            )
            if len(matches) == 1:
                matches[0].click()
                self.customer_detail.wait_loaded(timeout=self.timeout)
                return
        raise CustomerMutationSafetyError(
            "Unable to relocate the dedicated customer for restoration."
        )
