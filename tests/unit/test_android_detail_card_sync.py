from __future__ import annotations

from dataclasses import replace

import pytest

from scripts.prepare_excel import CURRENT_STEPS
from yanjia_automation.config import load_settings
from yanjia_automation.excel.models import ExcelCase
from yanjia_automation.excel.safety import (
    CUSTOMER_MUTATION_CASE_IDS,
    CUSTOMER_MUTATION_TAIL_ACTIONS,
    execution_skip_reason,
)
from yanjia_automation.excel.single_sheet import _detail_validation_steps

CONTRACTS = {
    "TC-DETAIL-026": (
        "customer_edit_username_et",
        "自动化${RUN_TOKEN}",
        "a_records_name_tv",
    ),
    "TC-DETAIL-027": (
        "customer_edit_unique_et",
        "${RUN_PHONE}",
        "a_records_unique_tv",
    ),
}


@pytest.mark.parametrize(("case_id", "contract"), CONTRACTS.items())
def test_card_sync_steps_match_both_execution_sources(
    case_id: str,
    contract: tuple[str, str, str],
) -> None:
    field_id, target_value, card_id = contract
    canonical = CURRENT_STEPS[case_id]
    migrated = _detail_validation_steps(
        row=68,
        case_id=case_id,
        input_data=None,
        timeout=10,
    )

    assert len(canonical) == 15
    assert len(migrated) == 15
    assert case_id in CUSTOMER_MUTATION_CASE_IDS
    for steps in (canonical, migrated):
        actions = tuple(
            step[1] if isinstance(step, tuple) else step.action for step in steps
        )
        locators = tuple(
            step[2] if isinstance(step, tuple) else step.locator for step in steps
        )
        inputs = tuple(
            step[3] if isinstance(step, tuple) else step.input_value for step in steps
        )
        assertions = tuple(
            step[4] if isinstance(step, tuple) else step.assertion for step in steps
        )
        expected = tuple(
            step[5] if isinstance(step, tuple) else step.expected for step in steps
        )
        notes = tuple(
            step[10] if isinstance(step, tuple) else step.note for step in steps
        )

        assert actions[-5:] == CUSTOMER_MUTATION_TAIL_ACTIONS[case_id]
        assert str(locators[-5]).endswith(field_id)
        assert inputs[-5] == target_value
        assert actions[-3] == "back"
        assert assertions[-2:] == ("activity_endswith", "text_equals")
        assert expected[-2:] == (".CustomerRecordsActivity", target_value)
        assert str(locators[-1]).endswith(card_id)
        assert any("恢复运行前姓名、手机号" in str(note or "") for note in notes)


@pytest.mark.parametrize("case_id", CONTRACTS)
def test_card_sync_cases_remain_fail_closed_until_preflight(case_id: str) -> None:
    case = ExcelCase(
        source_row=68,
        case_id=case_id,
        module="顾客详情",
        scenario="卡片同步",
        test_point="资料保存后卡片同步",
        priority="P1",
        precondition=None,
        input_data=None,
        expected_result=None,
        timeout=15,
        tags=("mutating", "requires_seed", "serial", "restorable"),
        enabled=True,
        automation_status="AUTOMATED",
        steps=(),
    )
    settings = replace(
        load_settings(),
        run_seeded=True,
        allow_mutation=True,
        mutation_customer_query="private-query",
    )
    assert execution_skip_reason(
        case,
        run_seeded=settings.run_seeded,
        allow_mutation=settings.allow_mutation,
        allow_destructive=settings.allow_destructive,
        mutation_customer_query=settings.mutation_customer_query,
        customer_preflight_verified=False,
    ) == "专用顾客真机只读预检尚未授权本次运行，写入用例保持闭锁"
