from __future__ import annotations

from dataclasses import replace
from unittest.mock import Mock, call

import pytest

from yanjia_automation.config import load_settings
from yanjia_automation.excel import runner as excel_runner_module
from yanjia_automation.excel.models import ExcelCase, ExcelStep
from yanjia_automation.excel.runner import ExcelCaseRunner
from yanjia_automation.excel.safety import execution_skip_reason
from yanjia_automation.excel.variables import VariableResolver
from yanjia_automation.flows import customer_mutation as customer_mutation_module
from yanjia_automation.flows.customer_mutation import (
    CustomerMutationSafetyError,
    CustomerMutationSession,
    CustomerProfileSnapshot,
    CustomerRestoreError,
    DedicatedCustomerMutationFlow,
    require_dedicated_customer_target,
    require_unique_customer_match,
)
from yanjia_automation.screens.customer import (
    CustomerDetailScreen,
    CustomerEditScreen,
    CustomerListScreen,
)
from yanjia_automation.screens.home import HomeScreen


def _customer_mutation_case() -> ExcelCase:
    return ExcelCase(
        source_row=1,
        case_id="TC-DETAIL-001",
        module="顾客详情",
        scenario="姓名必填",
        test_point="保存校验",
        priority="P0",
        precondition=None,
        input_data=None,
        expected_result=None,
        timeout=10,
        tags=("requires_seed", "mutating", "customer"),
        enabled=True,
        automation_status="AUTOMATED",
        steps=(),
    )


def _excel_step(order: int, action: str) -> ExcelStep:
    return ExcelStep(
        source_row=order,
        case_id="TC-DETAIL-001",
        order=order,
        name=action,
        action=action,
        locator=None,
        input_value=None,
        assertion=None,
        expected=None,
        timeout=10,
        element_index=0,
        continue_on_failure=False,
        enabled=True,
        note=None,
    )


def _authorized_settings():
    return replace(
        load_settings(),
        run_seeded=True,
        allow_mutation=True,
        mutation_customer_query="private-customer-query",
    )


@pytest.mark.parametrize(
    ("overrides", "expected_key"),
    [
        ({"run_seeded": False}, "YANJIA_RUN_SEEDED"),
        ({"allow_mutation": False}, "YANJIA_ALLOW_MUTATION"),
        ({"mutation_customer_query": None}, "YANJIA_MUTATION_CUSTOMER_QUERY"),
        ({"mutation_customer_query": "   "}, "YANJIA_MUTATION_CUSTOMER_QUERY"),
    ],
)
def test_dedicated_customer_target_requires_every_safety_precondition(
    overrides: dict[str, object], expected_key: str
) -> None:
    settings = replace(_authorized_settings(), **overrides)

    with pytest.raises(CustomerMutationSafetyError, match=expected_key):
        require_dedicated_customer_target(settings)


def test_dedicated_customer_query_is_redacted_from_representations() -> None:
    settings = _authorized_settings()

    target = require_dedicated_customer_target(settings)

    assert target.query == "private-customer-query"
    assert "private-customer-query" not in repr(settings)
    assert not settings.udid or settings.udid not in repr(settings)
    assert not settings.store_name or settings.store_name not in repr(settings)
    assert "private-customer-query" not in repr(target)
    assert "<redacted>" in repr(target)


def test_excel_customer_mutation_remains_locked_until_workflow_is_connected() -> None:
    case = _customer_mutation_case()

    assert execution_skip_reason(
        case,
        run_seeded=True,
        allow_mutation=True,
        allow_destructive=False,
        mutation_customer_query=None,
        customer_preflight_verified=False,
    ) == "专用顾客用例需要配置 YANJIA_MUTATION_CUSTOMER_QUERY"
    assert execution_skip_reason(
        case,
        run_seeded=True,
        allow_mutation=True,
        allow_destructive=False,
        mutation_customer_query="private-customer-query",
        customer_preflight_verified=False,
    ) == "专用顾客真机只读预检尚未授权本次运行，写入用例保持闭锁"
    assert (
        execution_skip_reason(
            case,
            run_seeded=True,
            allow_mutation=True,
            allow_destructive=False,
            mutation_customer_query="private-customer-query",
            customer_preflight_verified=True,
        )
        is None
    )


def test_unique_customer_match_rejects_zero_or_multiple_results_without_values() -> None:
    sensitive_match = "private-name-and-phone"

    with pytest.raises(CustomerMutationSafetyError, match="received 0"):
        require_unique_customer_match([])
    with pytest.raises(CustomerMutationSafetyError, match="received 2") as error:
        require_unique_customer_match([sensitive_match, sensitive_match])

    assert sensitive_match not in str(error.value)
    assert require_unique_customer_match([sensitive_match]) == sensitive_match


def test_customer_search_waits_for_the_expected_unique_result() -> None:
    field = Mock()
    stale_cards = [Mock(), Mock()]
    unique_card = Mock()
    screen = CustomerListScreen(Mock(), timeout=1)
    screen.find = Mock(return_value=field)
    screen.click = Mock()
    screen.find_all = Mock(side_effect=[stale_cards, [unique_card], [unique_card]])

    results = screen.search_customers(
        "private-query",
        timeout=1,
        expected_count=1,
    )

    assert results == [unique_card]
    field.clear.assert_called_once_with()
    field.send_keys.assert_called_once_with("private-query")


def test_dedicated_customer_flow_searches_unique_target_and_snapshots_in_memory(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    home = Mock(spec=HomeScreen)
    customer_list = Mock(spec=CustomerListScreen)
    customer_detail = Mock(spec=CustomerDetailScreen)
    customer_edit = Mock(spec=CustomerEditScreen)
    card = Mock()
    customer_list.search_customers.return_value = [card]
    customer_edit.profile_values.return_value = ("private-name", "private-phone")
    customer_edit.gender_value.return_value = "男"
    customer_edit.birthday_value.return_value = "2000-01-02"
    customer_edit.email_value.return_value = "private-email"
    customer_edit.marital_value.return_value = "保密"
    customer_edit.address_value.return_value = "private-address"
    customer_edit.remark_value.return_value = "private-remark"
    restart = Mock(return_value=home)
    monkeypatch.setattr(customer_mutation_module, "restart_to_home", restart)
    monkeypatch.setattr(
        customer_mutation_module,
        "CustomerListScreen",
        Mock(return_value=customer_list),
    )
    monkeypatch.setattr(
        customer_mutation_module,
        "CustomerDetailScreen",
        Mock(return_value=customer_detail),
    )
    monkeypatch.setattr(
        customer_mutation_module,
        "CustomerEditScreen",
        Mock(return_value=customer_edit),
    )
    driver = Mock()
    settings = _authorized_settings()

    snapshot = DedicatedCustomerMutationFlow(driver, settings).open_and_snapshot()

    restart.assert_called_once_with(driver, settings, timeout=30)
    home.open_customer_records.assert_called_once_with()
    customer_list.search_customers.assert_called_once_with(
        "private-customer-query",
        timeout=30,
        expected_count=1,
    )
    card.click.assert_called_once_with()
    customer_detail.open_editor.assert_called_once_with()
    customer_edit.scroll_to_top.assert_called_once_with()
    assert snapshot == CustomerProfileSnapshot(
        name="private-name",
        phone="private-phone",
        gender="男",
        birthday="2000-01-02",
        email="private-email",
        marital="保密",
        address="private-address",
        remark="private-remark",
    )


def test_dedicated_customer_flow_restores_saves_and_reopens_to_verify(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    home = Mock(spec=HomeScreen)
    customer_list = Mock(spec=CustomerListScreen)
    customer_detail = Mock(spec=CustomerDetailScreen)
    customer_edit = Mock(spec=CustomerEditScreen)
    card = Mock()
    customer_list.search_customers.return_value = [card]
    customer_edit.profile_values.side_effect = [
        ("original-name", "original-phone"),
        ("original-name", "original-phone"),
    ]
    customer_edit.gender_value.side_effect = ["男", "男"]
    customer_edit.birthday_value.side_effect = ["2000-01-02", "2000-01-02"]
    customer_edit.email_value.side_effect = ["original-email", "original-email"]
    customer_edit.marital_value.side_effect = ["保密", "保密"]
    customer_edit.address_value.side_effect = ["original-address", "original-address"]
    customer_edit.remark_value.side_effect = ["original-remark", "original-remark"]
    customer_edit.is_visible.return_value = True
    monkeypatch.setattr(
        customer_mutation_module,
        "restart_to_home",
        Mock(return_value=home),
    )
    monkeypatch.setattr(
        customer_mutation_module,
        "CustomerListScreen",
        Mock(return_value=customer_list),
    )
    monkeypatch.setattr(
        customer_mutation_module,
        "CustomerDetailScreen",
        Mock(return_value=customer_detail),
    )
    monkeypatch.setattr(
        customer_mutation_module,
        "CustomerEditScreen",
        Mock(return_value=customer_edit),
    )
    events: list[str] = []
    flow = DedicatedCustomerMutationFlow(Mock(), _authorized_settings())

    with flow.restoration_session() as session:
        session.run(lambda: events.append("mutation"))

    assert events == ["mutation"]
    customer_edit.set_profile.assert_called_once_with(
        name="original-name",
        phone="original-phone",
        gender="男",
        birthday="2000-01-02",
        email="original-email",
        marital="保密",
        address="original-address",
        remark="original-remark",
    )
    customer_edit.save.assert_called_once_with()
    assert customer_detail.open_editor.call_count == 2
    customer_edit.close.assert_called_once_with()


def test_restore_reopens_the_single_filtered_card_from_customer_list(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    customer_list = Mock(spec=CustomerListScreen)
    customer_detail = Mock(spec=CustomerDetailScreen)
    customer_edit = Mock(spec=CustomerEditScreen)
    card = Mock()
    customer_edit.is_visible.return_value = False
    customer_detail.is_visible.return_value = False
    customer_list.is_visible.return_value = True
    customer_list.find_all.return_value = [card]
    monkeypatch.setattr(
        customer_mutation_module,
        "CustomerListScreen",
        Mock(return_value=customer_list),
    )
    monkeypatch.setattr(
        customer_mutation_module,
        "CustomerDetailScreen",
        Mock(return_value=customer_detail),
    )
    monkeypatch.setattr(
        customer_mutation_module,
        "CustomerEditScreen",
        Mock(return_value=customer_edit),
    )
    snapshot = CustomerProfileSnapshot(
        name="original-name",
        phone="original-phone",
        gender="男",
        birthday="2000-01-02",
        email="original-email",
        marital="保密",
        address="original-address",
        remark="original-remark",
    )
    flow = DedicatedCustomerMutationFlow(Mock(), _authorized_settings())

    flow._ensure_editor_open(snapshot)

    card.click.assert_called_once_with()
    customer_list.search_customers.assert_not_called()
    customer_detail.open_editor.assert_called_once_with()
    customer_edit.wait_loaded.assert_called_once_with(timeout=30)


def test_restore_falls_back_to_unchanged_profile_key_without_logging_values(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    home = Mock(spec=HomeScreen)
    customer_list = Mock(spec=CustomerListScreen)
    customer_detail = Mock(spec=CustomerDetailScreen)
    customer_edit = Mock(spec=CustomerEditScreen)
    card = Mock()
    customer_edit.is_visible.return_value = False
    customer_detail.is_visible.return_value = False
    customer_list.is_visible.return_value = False
    customer_list.find_all.return_value = []
    customer_list.search_customers.side_effect = [[], [card]]
    restart = Mock(return_value=home)
    monkeypatch.setattr(customer_mutation_module, "restart_to_home", restart)
    monkeypatch.setattr(
        customer_mutation_module,
        "CustomerListScreen",
        Mock(return_value=customer_list),
    )
    monkeypatch.setattr(
        customer_mutation_module,
        "CustomerDetailScreen",
        Mock(return_value=customer_detail),
    )
    monkeypatch.setattr(
        customer_mutation_module,
        "CustomerEditScreen",
        Mock(return_value=customer_edit),
    )
    snapshot = CustomerProfileSnapshot(
        name="original-name",
        phone="original-phone",
        gender="男",
        birthday="2000-01-02",
        email="original-email",
        marital="保密",
        address="original-address",
        remark="original-remark",
    )
    flow = DedicatedCustomerMutationFlow(Mock(), _authorized_settings())

    flow._ensure_editor_open(snapshot)

    assert customer_list.search_customers.call_args_list == [
        call("original-phone", timeout=30, expected_count=1),
        call("original-name", timeout=30, expected_count=1),
    ]
    card.click.assert_called_once_with()
    customer_detail.open_editor.assert_called_once_with()


def test_excel_runner_delegates_mutation_tail_to_restoration_session(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    case = replace(
        _customer_mutation_case(),
        steps=(
            _excel_step(1, "restart_to_home"),
            _excel_step(2, "clear"),
            _excel_step(3, "click"),
            _excel_step(4, "assert"),
        ),
    )
    restore_events: list[str] = []
    session = CustomerMutationSession(
        CustomerProfileSnapshot(
            name="original-name",
            phone="original-phone",
            gender="男",
            birthday="2000-01-02",
            email="original-email",
            marital="保密",
            address="original-address",
            remark="original-remark",
        ),
        restore=lambda _: restore_events.append("restore"),
        verify=lambda _: restore_events.append("verify") or True,
    )
    flow = Mock()
    flow.restoration_session.return_value = session
    flow_factory = Mock(return_value=flow)
    monkeypatch.setattr(
        excel_runner_module,
        "DedicatedCustomerMutationFlow",
        flow_factory,
    )
    runner = ExcelCaseRunner(
        Mock(),
        _authorized_settings(),
        VariableResolver({}, sensitive_values=set()),
    )
    run_steps = Mock()
    monkeypatch.setattr(runner, "_run_steps", run_steps)

    runner.run(case)

    run_steps.assert_called_once_with(case, case.steps[-3:])
    assert restore_events == ["restore", "verify"]


def test_excel_runner_accepts_guarded_input_mutation_tail(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    case = replace(
        _customer_mutation_case(),
        case_id="TC-DETAIL-003",
        steps=(
            _excel_step(1, "restart_to_home"),
            _excel_step(2, "input"),
            _excel_step(3, "click"),
            _excel_step(4, "assert"),
        ),
    )
    session = CustomerMutationSession(
        CustomerProfileSnapshot(
            name="original-name",
            phone="original-phone",
            gender="男",
            birthday="2000-01-02",
            email="original-email",
            marital="保密",
            address="original-address",
            remark="original-remark",
        ),
        restore=lambda _: None,
        verify=lambda _: True,
    )
    flow = Mock()
    flow.restoration_session.return_value = session
    monkeypatch.setattr(
        excel_runner_module,
        "DedicatedCustomerMutationFlow",
        Mock(return_value=flow),
    )
    runner = ExcelCaseRunner(
        Mock(),
        _authorized_settings(),
        VariableResolver({}, sensitive_values=set()),
    )
    run_steps = Mock()
    monkeypatch.setattr(runner, "_run_steps", run_steps)

    runner.run(case)

    run_steps.assert_called_once_with(case, case.steps[-3:])


def test_profile_snapshot_repr_never_exposes_sensitive_fields() -> None:
    snapshot = CustomerProfileSnapshot(
        name="private-name",
        phone="private-phone",
        gender="男",
        birthday="2000-01-02",
        email="private-email",
        marital="保密",
        address="private-address",
        remark="private-remark",
    )

    assert "private-name" not in repr(snapshot)
    assert "private-phone" not in repr(snapshot)
    assert "男" not in repr(snapshot)
    assert "2000-01-02" not in repr(snapshot)
    assert "private-email" not in repr(snapshot)
    assert "保密" not in repr(snapshot)
    assert "private-address" not in repr(snapshot)
    assert "private-remark" not in repr(snapshot)
    assert repr(snapshot) == (
        "CustomerProfileSnapshot(name=<redacted>, phone=<redacted>, "
        "gender=<redacted>, birthday=<redacted>, email=<redacted>, "
        "marital=<redacted>, address=<redacted>, remark=<redacted>)"
    )


def test_mutation_session_restores_and_verifies_after_success() -> None:
    snapshot = CustomerProfileSnapshot(
        name="original-name",
        phone="original-phone",
        gender="男",
        birthday="2000-01-02",
        email="original-email",
        marital="保密",
        address="original-address",
        remark="original-remark",
    )
    events: list[str] = []

    with CustomerMutationSession(
        snapshot,
        restore=lambda current: events.append(f"restore:{current.name}"),
        verify=lambda current: events.append(f"verify:{current.phone}") or True,
    ) as session:
        result = session.run(lambda: events.append("mutate") or "result")

    assert result == "result"
    assert events == ["mutate", "restore:original-name", "verify:original-phone"]


def test_mutation_session_restores_when_mutation_raises() -> None:
    snapshot = CustomerProfileSnapshot(
        name="original-name",
        phone="original-phone",
        gender="男",
        birthday="2000-01-02",
        email="original-email",
        marital="保密",
        address="original-address",
        remark="original-remark",
    )
    events: list[str] = []

    def fail_after_partial_write() -> None:
        events.append("partial-write")
        raise ValueError("mutation failed")

    with pytest.raises(ValueError, match="mutation failed"):
        with CustomerMutationSession(
            snapshot,
            restore=lambda _: events.append("restore"),
            verify=lambda _: events.append("verify") or True,
        ) as session:
            session.run(fail_after_partial_write)

    assert events == ["partial-write", "restore", "verify"]


def test_mutation_session_fails_closed_when_restore_cannot_be_verified() -> None:
    snapshot = CustomerProfileSnapshot(
        name="original-name",
        phone="original-phone",
        gender="男",
        birthday="2000-01-02",
        email="original-email",
        marital="保密",
        address="original-address",
        remark="original-remark",
    )

    with pytest.raises(CustomerRestoreError, match="could not be verified"):
        with CustomerMutationSession(
            snapshot,
            restore=lambda _: None,
            verify=lambda _: False,
        ) as session:
            session.run(lambda: None)


def test_mutation_session_wraps_restore_errors_and_disallows_reuse() -> None:
    snapshot = CustomerProfileSnapshot(
        name="original-name",
        phone="original-phone",
        gender="男",
        birthday="2000-01-02",
        email="original-email",
        marital="保密",
        address="original-address",
        remark="original-remark",
    )

    def fail_restore(_: CustomerProfileSnapshot) -> None:
        raise RuntimeError("restore backend failed")

    session = CustomerMutationSession(snapshot, restore=fail_restore, verify=lambda _: True)
    with pytest.raises(CustomerRestoreError, match="stop further mutation tests") as error:
        with session:
            session.run(lambda: None)

    assert isinstance(error.value.__cause__, RuntimeError)
    with pytest.raises(CustomerMutationSafetyError, match="cannot be reused"):
        with session:
            pass


def test_mutation_session_requires_context_and_one_callback() -> None:
    snapshot = CustomerProfileSnapshot(
        name="original-name",
        phone="original-phone",
        gender="男",
        birthday="2000-01-02",
        email="original-email",
        marital="保密",
        address="original-address",
        remark="original-remark",
    )
    session = CustomerMutationSession(snapshot, restore=lambda _: None, verify=lambda _: True)

    with pytest.raises(CustomerMutationSafetyError, match="inside"):
        session.run(lambda: None)

    with session:
        assert session.run(lambda: "first") == "first"
        with pytest.raises(CustomerMutationSafetyError, match="exactly one"):
            session.run(lambda: "second")
