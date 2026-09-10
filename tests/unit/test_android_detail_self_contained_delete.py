from __future__ import annotations

from dataclasses import replace
from unittest.mock import Mock

import pytest
from appium.webdriver.common.appiumby import AppiumBy

from scripts.prepare_excel import CURRENT_STEPS
from yanjia_automation.config import load_settings
from yanjia_automation.excel import runner as runner_module
from yanjia_automation.excel.android_catalog import (
    SINGLE_SHEET_SOURCE,
    STEP_SHEET_SOURCE,
    get_android_case_definition,
)
from yanjia_automation.excel.models import ExcelCase, ExcelStep
from yanjia_automation.excel.runner import ExcelCaseRunner
from yanjia_automation.excel.safety import execution_skip_reason
from yanjia_automation.excel.single_sheet import _detail_validation_steps
from yanjia_automation.excel.variables import VariableResolver
from yanjia_automation.flows import remark_mutation as remark_mutation_module
from yanjia_automation.flows.remark_mutation import (
    DedicatedRemarkMutationFlow,
    RemarkMutationSafetyError,
    RemarkMutationSession,
    RemarkSnapshot,
)
from yanjia_automation.flows.tag_mutation import (
    DedicatedTagMutationFlow,
    TagSnapshot,
)
from yanjia_automation.screens.base import resource_id


@pytest.mark.parametrize(
    ("case_id", "tail", "label", "delete_action"),
    [
        (
            "TC-DETAIL-020",
            [
                "click",
                "assert",
                "input",
                "click",
                "assert",
                "delete_current_run_tag",
                "assert",
            ],
            "自动化删除标签${RUN_TOKEN}",
            "delete_current_run_tag",
        ),
        (
            "TC-DETAIL-037",
            [
                "assert",
                "input",
                "assert",
                "click",
                "assert",
                "delete_current_run_remark",
                "assert",
            ],
            "自动化删除备注${RUN_TOKEN}",
            "delete_current_run_remark",
        ),
    ],
)
def test_self_contained_delete_steps_match_both_execution_sources(
    case_id: str,
    tail: list[str],
    label: str,
    delete_action: str,
) -> None:
    canonical = CURRENT_STEPS[case_id]
    migrated = _detail_validation_steps(
        row=62 if case_id.endswith("020") else 79,
        case_id=case_id,
        input_data=None,
        timeout=10,
    )

    assert len(canonical) == 18
    assert len(migrated) == 18
    for steps in (canonical, migrated):
        actions = [step[1] if isinstance(step, tuple) else step.action for step in steps]
        locators = [step[2] if isinstance(step, tuple) else step.locator for step in steps]
        inputs = [step[3] if isinstance(step, tuple) else step.input_value for step in steps]
        expected = [step[5] if isinstance(step, tuple) else step.expected for step in steps]
        notes = [step[10] if isinstance(step, tuple) else step.note for step in steps]

        assert actions[-7:] == tail
        assert label in inputs
        assert label in expected
        assert locators[actions.index(delete_action)] is None
        assert any("仅删除本轮" in str(note or "") for note in notes)


def _destructive_case(case_id: str) -> ExcelCase:
    return ExcelCase(
        source_row=1,
        case_id=case_id,
        module="顾客详情",
        scenario="自包含删除",
        test_point="仅删除本轮数据",
        priority="P1",
        precondition=None,
        input_data=None,
        expected_result=None,
        timeout=10,
        tags=("destructive", "requires_seed", "serial", "restorable"),
        enabled=True,
        automation_status="AUTOMATED",
        steps=(),
    )


@pytest.mark.parametrize("case_id", ["TC-DETAIL-020", "TC-DETAIL-037"])
def test_self_contained_delete_cases_require_all_safety_gates(case_id: str) -> None:
    case = _destructive_case(case_id)

    assert execution_skip_reason(
        case,
        run_seeded=True,
        allow_mutation=False,
        allow_destructive=True,
        mutation_customer_query="private-query",
        customer_preflight_verified=True,
    ) == "破坏性用例需要同时授权 --allow-mutation --allow-destructive"
    assert execution_skip_reason(
        case,
        run_seeded=True,
        allow_mutation=True,
        allow_destructive=True,
        mutation_customer_query="private-query",
        customer_preflight_verified=False,
    ) == "专用顾客真机只读预检尚未授权本次运行，写入用例保持闭锁"
    assert execution_skip_reason(
        case,
        run_seeded=True,
        allow_mutation=True,
        allow_destructive=True,
        mutation_customer_query="private-query",
        customer_preflight_verified=True,
    ) is None


@pytest.mark.parametrize(
    ("case_id", "flow", "value", "locator"),
    [
        (
            "TC-DETAIL-020",
            "detail_image_tag_self_contained_delete",
            "自动化删除标签${RUN_TOKEN}",
            "records_remark_tag_et",
        ),
        (
            "TC-DETAIL-037",
            "detail_consultation_remark_self_contained_delete",
            "自动化删除备注${RUN_TOKEN}",
            "records_remark_remark_et",
        ),
    ],
)
def test_catalog_registers_restorable_destructive_contract(
    case_id: str,
    flow: str,
    value: str,
    locator: str,
) -> None:
    definition = get_android_case_definition(case_id)

    assert definition is not None
    assert definition.supports(SINGLE_SHEET_SOURCE)
    assert definition.supports(STEP_SHEET_SOURCE)
    assert definition.mutability == "destructive"
    assert definition.requires_seed is True
    assert "restorable" in definition.tags
    assert definition.flow == flow
    assert definition.filter_value == value
    assert definition.filter_locator == locator


def test_remark_snapshot_repr_redacts_all_content() -> None:
    snapshot = RemarkSnapshot(("private-remark",), 1, 1)

    assert "private-remark" not in repr(snapshot)
    assert repr(snapshot) == (
        "RemarkSnapshot(texts=<redacted>, record_count=1, delete_count=1)"
    )


def test_remark_mutation_session_restores_after_success_and_failure() -> None:
    snapshot = RemarkSnapshot(("original",), 1, 1)
    events: list[str] = []

    with RemarkMutationSession(
        snapshot,
        restore=lambda _: events.append("restore"),
        verify=lambda _: events.append("verify") or True,
    ) as session:
        session.run(lambda: events.append("mutate"))
    assert events == ["mutate", "restore", "verify"]

    events.clear()
    with pytest.raises(ValueError, match="failed"):
        with RemarkMutationSession(
            snapshot,
            restore=lambda _: events.append("restore"),
            verify=lambda _: events.append("verify") or True,
        ) as session:
            session.run(lambda: (_ for _ in ()).throw(ValueError("failed")))
    assert events == ["restore", "verify"]


def test_remark_collection_diff_accepts_only_one_addition() -> None:
    before = RemarkSnapshot(("existing",), 1, 1)
    after = RemarkSnapshot(("new", "existing"), 2, 2)
    unsafe = RemarkSnapshot(("new-a", "new-b", "existing"), 3, 3)

    assert DedicatedRemarkMutationFlow._created_display_text(before, after) == "new"
    with pytest.raises(RemarkMutationSafetyError, match="not exactly one addition"):
        DedicatedRemarkMutationFlow._created_display_text(before, unsafe)


def test_remark_delete_targets_the_generated_text_parent_row(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    display_text = 'generated "quoted" remark'
    content = Mock()
    content.get_attribute.return_value = display_text
    row_delete = Mock()
    confirm = Mock()
    driver = Mock()

    def find_elements(by: str, value: str):
        if (by, value) == resource_id("a_records_remark_list_content_tv"):
            return [content]
        if by == AppiumBy.XPATH:
            assert "/..//*" in value
            assert "a_records_remark_list_delete_ifv" in value
            assert "generated" in value
            return [row_delete]
        if (by, value) == resource_id("cover_prompt_right_tv"):
            return [confirm]
        return []

    driver.find_elements.side_effect = find_elements
    flow = object.__new__(DedicatedRemarkMutationFlow)
    flow.driver = driver
    flow.timeout = 1

    class ImmediateWait:
        def __init__(self, current_driver, timeout) -> None:
            del current_driver, timeout

        def until(self, condition):
            assert condition(driver)
            return True

    monkeypatch.setattr(remark_mutation_module, "WebDriverWait", ImmediateWait)

    flow._delete_exact_display_remark(display_text)

    row_delete.click.assert_called_once_with()
    confirm.click.assert_called_once_with()


def _step(order: int, action: str, case_id: str) -> ExcelStep:
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


def test_runner_delegates_guarded_tail_to_remark_restoration(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    tail = (
        "assert",
        "input",
        "assert",
        "click",
        "assert",
        "delete_current_run_remark",
        "assert",
    )
    case = replace(
        _destructive_case("TC-DETAIL-037"),
        steps=(_step(1, "restart_to_home", "TC-DETAIL-037"),)
        + tuple(
            _step(index + 2, action, "TC-DETAIL-037")
            for index, action in enumerate(tail)
        ),
    )
    events: list[str] = []
    session = RemarkMutationSession(
        RemarkSnapshot((), 0, 0),
        restore=lambda _: events.append("restore"),
        verify=lambda _: events.append("verify") or True,
    )
    flow = Mock()
    flow.restoration_session.return_value = session
    monkeypatch.setattr(
        runner_module,
        "DedicatedRemarkMutationFlow",
        Mock(return_value=flow),
    )
    runner = ExcelCaseRunner(
        Mock(),
        replace(
            load_settings(),
            run_seeded=True,
            allow_mutation=True,
            allow_destructive=True,
            mutation_customer_query="private-query",
            customer_preflight_verified=True,
        ),
        VariableResolver({}, sensitive_values=set()),
    )
    run_steps = Mock()
    monkeypatch.setattr(runner, "_run_steps", run_steps)

    runner.run(case)

    flow.restoration_session.assert_called_once_with("consultation")
    run_steps.assert_called_once_with(case, case.steps[-7:])
    assert events == ["restore", "verify"]


def test_runner_custom_delete_actions_require_and_use_active_session() -> None:
    runner = ExcelCaseRunner(
        Mock(),
        load_settings(),
        VariableResolver({}, sensitive_values=set()),
    )
    tag_flow = Mock(spec=DedicatedTagMutationFlow)
    tag_snapshot = TagSnapshot((), 0, 0)
    remark_flow = Mock(spec=DedicatedRemarkMutationFlow)
    remark_snapshot = RemarkSnapshot((), 0, 0)

    with pytest.raises(RuntimeError, match="受控标签恢复会话"):
        runner._delete_current_run_tag(
            _destructive_case("TC-DETAIL-020"),
            _step(1, "delete_current_run_tag", "TC-DETAIL-020"),
        )
    runner._active_tag_mutation = (tag_flow, tag_snapshot)
    runner._delete_current_run_tag(
        _destructive_case("TC-DETAIL-020"),
        _step(1, "delete_current_run_tag", "TC-DETAIL-020"),
    )
    tag_flow.delete_created_tag.assert_called_once_with(tag_snapshot)

    with pytest.raises(RuntimeError, match="受控备注恢复会话"):
        runner._delete_current_run_remark(
            _destructive_case("TC-DETAIL-037"),
            _step(1, "delete_current_run_remark", "TC-DETAIL-037"),
        )
    runner._active_remark_mutation = (remark_flow, remark_snapshot)
    runner._delete_current_run_remark(
        _destructive_case("TC-DETAIL-037"),
        _step(1, "delete_current_run_remark", "TC-DETAIL-037"),
    )
    remark_flow.delete_created_remark.assert_called_once_with(remark_snapshot)
