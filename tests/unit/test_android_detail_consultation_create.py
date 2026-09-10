from __future__ import annotations

from scripts.prepare_excel import CURRENT_STEPS
from yanjia_automation.excel.models import ExcelCase
from yanjia_automation.excel.runner import _first_unlinked_image_index
from yanjia_automation.excel.safety import execution_skip_reason


def test_consultation_create_keeps_new_record_and_uses_dynamic_image() -> None:
    steps = CURRENT_STEPS["TC-DETAIL-030"]

    assert len(steps) == 18
    assert steps[3][3] == "咨询单特定"
    assert steps[5][4:6] == ("element_count_equals", "1")
    assert [step[1] for step in steps].count("open_first_unlinked_image") == 1
    assert [step[1] for step in steps].count("verify_selected_image_consultation") == 1
    assert steps[14][2].endswith(":id/cover_prompt_v2_tv")
    assert steps[14][4:6] == ("text_equals", "保存并退出")
    assert steps[-1][0] == "校验新增咨询单已保留"
    assert not any("delete" in step[1] for step in steps)


def test_first_unlinked_image_index_normalizes_whitespace() -> None:
    assert (
        _first_unlinked_image_index(
            ["2026-08-25 10:00", "2026-08-26\n11:30", "2026-08-27 09:15"],
            {"2026-08-2510:00", "2026-08-27 09:15"},
        )
        == 1
    )
    assert _first_unlinked_image_index(["2026-08-25"], {"2026-08-25"}) is None
    assert _first_unlinked_image_index(["  "], set()) is None


def test_persistent_consultation_requires_seed_and_mutation_only() -> None:
    case = ExcelCase(
        source_row=72,
        case_id="TC-DETAIL-030",
        module="顾客详情",
        scenario="咨询单-添加咨询单",
        test_point="验证咨询单添加功能",
        priority="P1",
        precondition=None,
        input_data="咨询单特定",
        expected_result=None,
        timeout=15,
        tags=("mutating", "requires_seed", "serial", "persistent"),
        enabled=True,
        automation_status="AUTOMATED",
        steps=(),
    )

    assert "--run-seeded" in (
        execution_skip_reason(
            case,
            run_seeded=False,
            allow_mutation=False,
            allow_destructive=False,
            mutation_customer_query=None,
            customer_preflight_verified=False,
        )
        or ""
    )
    assert "--allow-mutation" in (
        execution_skip_reason(
            case,
            run_seeded=True,
            allow_mutation=False,
            allow_destructive=False,
            mutation_customer_query=None,
            customer_preflight_verified=False,
        )
        or ""
    )
    assert (
        execution_skip_reason(
            case,
            run_seeded=True,
            allow_mutation=True,
            allow_destructive=False,
            mutation_customer_query=None,
            customer_preflight_verified=False,
        )
        is None
    )
