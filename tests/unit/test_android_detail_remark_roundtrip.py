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
from yanjia_automation.flows.remark_mutation import RemarkMutationSession, RemarkSnapshot

CASES = {
    "TC-DETAIL-012": (
        "image",
        "a_records_detail_all_remark_ifv",
        "测试备注${RUN_TOKEN}",
    ),
    "TC-DETAIL-021": (
        "image",
        "a_records_detail_all_remark_ifv",
        "影像备注${RUN_TOKEN}",
    ),
    "TC-DETAIL-022": (
        "image",
        "a_records_detail_all_remark_ifv",
        "……&&*……*&${RUN_TOKEN}",
    ),
    "TC-DETAIL-035": (
        "consultation",
        "a_consultation_result_remark_ifv",
        "咨询单备注${RUN_TOKEN}",
    ),
}


@pytest.mark.parametrize(("case_id", "contract"), CASES.items())
def test_remark_roundtrip_steps_match_both_execution_sources(
    case_id: str,
    contract: tuple[str, str, str],
) -> None:
    _, entry_locator, label = contract
    canonical = CURRENT_STEPS[case_id]
    migrated = _detail_validation_steps(
        row=63,
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
        assert actions[-5:] == ["assert", "input", "assert", "click", "assert"]
        assert inputs[-4] == label
        assert assertions[-3] == "element_enabled"
        assert assertions[-1] == "text_contains"
        assert expected[-1] == label
        assert any("本轮唯一新增备注" in str(note or "") for note in notes)
        assert "delete_current_run_remark" not in actions


def _remark_case(case_id: str) -> ExcelCase:
    return ExcelCase(
        source_row=63,
        case_id=case_id,
        module="顾客详情",
        scenario="备注创建",
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


@pytest.mark.parametrize("case_id", CASES)
def test_remark_roundtrip_cases_are_fail_closed_until_all_gates_are_enabled(
    case_id: str,
) -> None:
    case = _remark_case(case_id)

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


@pytest.mark.parametrize(
    ("case_id", "scope"),
    [
        ("TC-DETAIL-012", "image"),
        ("TC-DETAIL-021", "image"),
        ("TC-DETAIL-035", "consultation"),
    ],
)
def test_runner_delegates_guarded_tail_to_scoped_remark_restoration(
    monkeypatch: pytest.MonkeyPatch,
    case_id: str,
    scope: str,
) -> None:
    tail = ("assert", "input", "assert", "click", "assert")
    case = replace(
        _remark_case(case_id),
        steps=(_step(1, "restart_to_home", case_id),)
        + tuple(
            _step(index + 2, action, case_id)
            for index, action in enumerate(tail)
        ),
    )
    events: list[str] = []
    session = RemarkMutationSession(
        RemarkSnapshot(("original",), 1, 1),
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

    flow.restoration_session.assert_called_once_with(scope)
    run_steps.assert_called_once_with(case, case.steps[-5:])
    assert events == ["restore", "verify"]
