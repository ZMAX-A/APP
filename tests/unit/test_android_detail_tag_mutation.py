from __future__ import annotations

from dataclasses import replace
from unittest.mock import Mock

import pytest

from scripts.prepare_excel import CURRENT_STEPS
from yanjia_automation.config import load_settings
from yanjia_automation.excel import runner as runner_module
from yanjia_automation.excel.models import ExcelCase, ExcelStep
from yanjia_automation.excel.runner import ExcelCaseRunner
from yanjia_automation.excel.safety import execution_skip_reason
from yanjia_automation.excel.single_sheet import _detail_validation_steps
from yanjia_automation.excel.variables import VariableResolver
from yanjia_automation.flows import tag_mutation as tag_mutation_module
from yanjia_automation.flows.tag_mutation import (
    DedicatedTagMutationFlow,
    TagMutationSafetyError,
    TagMutationSession,
    TagRestoreError,
    TagSnapshot,
)
from yanjia_automation.screens.base import resource_id

CASES = {
    "TC-DETAIL-016": ("a_records_detail_all_remark_ifv", "自动化标签${RUN_TOKEN}"),
    "TC-DETAIL-018": ("a_records_detail_all_remark_ifv", "%……&*${RUN_TOKEN}"),
    "TC-DETAIL-031": ("a_consultation_result_remark_ifv", "咨询单标签${RUN_TOKEN}"),
    "TC-DETAIL-033": ("a_consultation_result_remark_ifv", "%……&*${RUN_TOKEN}"),
}


@pytest.mark.parametrize(("case_id", "contract"), CASES.items())
def test_tag_mutation_steps_create_unique_tag_inside_restoration_contract(
    case_id: str,
    contract: tuple[str, str],
) -> None:
    entry_locator, label = contract
    canonical = CURRENT_STEPS[case_id]
    migrated = _detail_validation_steps(
        row=58,
        case_id=case_id,
        input_data=None,
        timeout=10,
    )

    assert len(canonical) == 16
    assert len(migrated) == 16
    for steps in (canonical, migrated):
        actions = [step[1] if isinstance(step, tuple) else step.action for step in steps]
        locators = [step[2] if isinstance(step, tuple) else step.locator for step in steps]
        inputs = [step[3] if isinstance(step, tuple) else step.input_value for step in steps]
        assertions = [step[4] if isinstance(step, tuple) else step.assertion for step in steps]
        expected = [step[5] if isinstance(step, tuple) else step.expected for step in steps]
        notes = [step[10] if isinstance(step, tuple) else step.note for step in steps]

        assert any(str(locator or "").endswith(entry_locator) for locator in locators)
        assert actions[-5:] == ["click", "assert", "input", "click", "assert"]
        assert inputs[-3] == label
        assert assertions[-1] == "text_contains"
        assert expected[-1] == label
        assert any("差集新增标签" in str(note or "") for note in notes)
        assert not any("delete" in str(locator or "") for locator in locators)


def _tag_case(case_id: str = "TC-DETAIL-016") -> ExcelCase:
    return ExcelCase(
        source_row=58,
        case_id=case_id,
        module="顾客详情",
        scenario="标签创建",
        test_point="创建后恢复",
        priority="P1",
        precondition=None,
        input_data=None,
        expected_result=None,
        timeout=10,
        tags=("mutating", "requires_seed", "serial", "restorable"),
        enabled=True,
        automation_status="AUTOMATED",
        steps=(),
    )


def _step(order: int, action: str, case_id: str = "TC-DETAIL-016") -> ExcelStep:
    return ExcelStep(
        source_row=order,
        case_id=case_id,
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


def test_tag_mutation_case_is_fail_closed_until_all_gates_are_enabled() -> None:
    case = _tag_case()

    assert execution_skip_reason(
        case,
        run_seeded=True,
        allow_mutation=False,
        allow_destructive=False,
        mutation_customer_query="private-query",
        customer_preflight_verified=True,
    ) == "写入用例需要 --allow-mutation 或 YANJIA_ALLOW_MUTATION=true"
    assert execution_skip_reason(
        case,
        run_seeded=True,
        allow_mutation=True,
        allow_destructive=False,
        mutation_customer_query="private-query",
        customer_preflight_verified=False,
    ) == "专用顾客真机只读预检尚未授权本次运行，写入用例保持闭锁"
    assert execution_skip_reason(
        case,
        run_seeded=True,
        allow_mutation=True,
        allow_destructive=False,
        mutation_customer_query="private-query",
        customer_preflight_verified=True,
    ) is None


def test_tag_snapshot_repr_redacts_all_tag_text() -> None:
    snapshot = TagSnapshot(("private-tag", "%private%"), 2, 1)

    assert "private-tag" not in repr(snapshot)
    assert "%private%" not in repr(snapshot)
    assert repr(snapshot) == (
        "TagSnapshot(texts=<redacted>, container_count=2, delete_count=1)"
    )


def test_tag_mutation_session_restores_on_success_and_failure() -> None:
    snapshot = TagSnapshot(("original",), 1, 1)
    events: list[str] = []

    with TagMutationSession(
        snapshot,
        restore=lambda _: events.append("restore"),
        verify=lambda _: events.append("verify") or True,
    ) as session:
        session.run(lambda: events.append("mutate"))
    assert events == ["mutate", "restore", "verify"]

    events.clear()
    with pytest.raises(ValueError, match="failed"):
        with TagMutationSession(
            snapshot,
            restore=lambda _: events.append("restore"),
            verify=lambda _: events.append("verify") or True,
        ) as session:
            session.run(lambda: (_ for _ in ()).throw(ValueError("failed")))
    assert events == ["restore", "verify"]


def test_tag_mutation_session_fails_closed_when_restore_is_unverified() -> None:
    snapshot = TagSnapshot(("original",), 1, 1)

    with pytest.raises(TagRestoreError, match="could not be verified"):
        with TagMutationSession(
            snapshot,
            restore=lambda _: None,
            verify=lambda _: False,
        ) as session:
            session.run(lambda: None)


def test_tag_restore_uses_collection_diff_instead_of_original_input_text(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    before = TagSnapshot(("existing-a", "existing-b"), 2, 1)
    displayed_by_app = "·自动化标签123456"
    after = TagSnapshot((displayed_by_app, "existing-a", "existing-b"), 3, 2)
    flow = object.__new__(DedicatedTagMutationFlow)
    flow.driver = Mock()
    flow.timeout = 1
    flow._dismiss_add_dialog = Mock()
    flow._delete_exact_display_tag = Mock()
    flow._close_remark = Mock()
    flow._snapshot = Mock(side_effect=[after, before])

    class ImmediateWait:
        def __init__(self, driver, timeout) -> None:
            del driver, timeout

        def until(self, condition):
            assert condition(None)
            return True

    monkeypatch.setattr(tag_mutation_module, "WebDriverWait", ImmediateWait)

    flow._restore_tags(before)

    flow._delete_exact_display_tag.assert_called_once_with(displayed_by_app)
    flow._close_remark.assert_called_once_with()


def test_consultation_tag_flow_searches_exact_phone_and_waits_for_remark_entry(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    query = "private-phone"
    settings = replace(
        load_settings(),
        run_seeded=True,
        allow_mutation=True,
        mutation_customer_query=query,
        customer_preflight_verified=True,
    )
    driver = Mock()
    driver.current_activity = ".RecordsRemarkActivity"
    home = Mock()
    card = Mock()
    identifier = Mock()
    identifier.get_attribute.return_value = query
    entry = Mock()
    customer_list = Mock()
    customer_list.cards = resource_id("a_records_cl")
    customer_list.card_identifier = resource_id("a_records_unique_tv")
    customer_list.find_all.side_effect = lambda locator: (
        [identifier] if locator == customer_list.card_identifier else [card]
    )
    customer_detail = Mock()
    entry_calls = iter(([], [entry]))
    customer_detail.find_all.side_effect = lambda _locator: next(entry_calls)

    class PollingWait:
        def __init__(self, current_driver, timeout) -> None:
            self.driver = current_driver
            del timeout

        def until(self, condition):
            for _ in range(3):
                result = condition(self.driver)
                if result:
                    return result
            raise AssertionError("condition did not become true")

    monkeypatch.setattr(tag_mutation_module, "WebDriverWait", PollingWait)
    restart = Mock(return_value=home)
    monkeypatch.setattr(tag_mutation_module, "restart_to_home", restart)
    flow = DedicatedTagMutationFlow(driver, settings)
    flow.customer_list = customer_list
    flow.customer_detail = customer_detail
    expected_snapshot = TagSnapshot(("existing",), 1, 1)
    flow._snapshot = Mock(return_value=expected_snapshot)

    snapshot = flow.open_and_snapshot("consultation")

    assert snapshot == expected_snapshot
    customer_list.search_customers.assert_called_once_with(
        query, timeout=30, expected_count=1
    )
    card.click.assert_called_once_with()
    entry.click.assert_called_once_with()


def test_tag_restore_refuses_any_delta_larger_than_one_addition() -> None:
    before = TagSnapshot(("existing",), 1, 1)
    unsafe = TagSnapshot(("new-a", "new-b", "existing"), 3, 3)
    flow = object.__new__(DedicatedTagMutationFlow)
    flow._dismiss_add_dialog = Mock()
    flow._snapshot = Mock(return_value=unsafe)

    with pytest.raises(TagMutationSafetyError, match="not exactly one addition"):
        flow._restore_tags(before)


def test_runner_delegates_only_guarded_tail_to_tag_restoration(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    tail = ("click", "assert", "input", "click", "assert")
    case = replace(
        _tag_case(),
        steps=(_step(1, "restart_to_home"),)
        + tuple(_step(index + 2, action) for index, action in enumerate(tail)),
    )
    events: list[str] = []
    session = TagMutationSession(
        TagSnapshot(("original",), 1, 1),
        restore=lambda _: events.append("restore"),
        verify=lambda _: events.append("verify") or True,
    )
    flow = Mock()
    flow.restoration_session.return_value = session
    monkeypatch.setattr(runner_module, "DedicatedTagMutationFlow", Mock(return_value=flow))
    settings = replace(
        load_settings(),
        run_seeded=True,
        allow_mutation=True,
        mutation_customer_query="private-query",
        customer_preflight_verified=True,
    )
    runner = ExcelCaseRunner(
        Mock(),
        settings,
        VariableResolver({}, sensitive_values=set()),
    )
    run_steps = Mock()
    monkeypatch.setattr(runner, "_run_steps", run_steps)

    runner.run(case)

    flow.restoration_session.assert_called_once_with("image")
    run_steps.assert_called_once_with(case, case.steps[-5:])
    assert events == ["restore", "verify"]
