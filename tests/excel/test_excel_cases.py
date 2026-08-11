from __future__ import annotations

import json
import time
from datetime import datetime

import allure
import pytest

from yanjia_automation.config import Settings
from yanjia_automation.driver import DriverManager, is_recoverable_driver_error
from yanjia_automation.excel.models import ExcelCase
from yanjia_automation.excel.results import ExcelResult, ExcelResultWriter
from yanjia_automation.excel.runner import ExcelCaseRunner
from yanjia_automation.excel.variables import VariableResolver


def test_excel_case(
    excel_case: ExcelCase,
    request: pytest.FixtureRequest,
    pytestconfig: pytest.Config,
    settings: Settings,
    excel_variables: VariableResolver,
    excel_result_writer: ExcelResultWriter,
) -> None:
    _configure_allure(excel_case)
    started_at = datetime.now()
    started_clock = time.perf_counter()

    skip_reason = _safety_skip_reason(excel_case, pytestconfig, settings)
    if skip_reason:
        finished_at = datetime.now()
        excel_result_writer.record(
            ExcelResult(
                case_id=excel_case.case_id,
                status="SKIP",
                started_at=started_at,
                finished_at=finished_at,
                duration_seconds=time.perf_counter() - started_clock,
                error_message=skip_reason,
            )
        )
        pytest.skip(skip_reason)

    driver_manager: DriverManager = request.getfixturevalue("driver_manager")
    driver = driver_manager.get()
    runner = ExcelCaseRunner(driver, settings, excel_variables)
    try:
        try:
            runner.run(excel_case)
        except Exception as first_error:
            if not _can_retry_infrastructure(excel_case, first_error):
                raise
            allure.attach(
                excel_variables.redact(first_error),
                name="基础设施异常-自动重试一次",
                attachment_type=allure.attachment_type.TEXT,
            )
            allure.dynamic.label("infrastructure_retry", "1")
            driver = driver_manager.restart()
            runner = ExcelCaseRunner(driver, settings, excel_variables)
            runner.run(excel_case)
    except Exception as error:
        finished_at = datetime.now()
        safe_error = excel_variables.redact(error)
        try:
            runner.attach_failure_evidence()
        except Exception as evidence_error:
            allure.attach(
                excel_variables.redact(evidence_error),
                name="失败证据采集异常",
                attachment_type=allure.attachment_type.TEXT,
            )
        excel_result_writer.record(
            ExcelResult(
                case_id=excel_case.case_id,
                status="FAIL" if isinstance(error, AssertionError) else "ERROR",
                started_at=started_at,
                finished_at=finished_at,
                duration_seconds=time.perf_counter() - started_clock,
                error_message=f"{type(error).__name__}: {safe_error}",
            )
        )
        raise
    else:
        finished_at = datetime.now()
        excel_result_writer.record(
            ExcelResult(
                case_id=excel_case.case_id,
                status="PASS",
                started_at=started_at,
                finished_at=finished_at,
                duration_seconds=time.perf_counter() - started_clock,
            )
        )


def _configure_allure(case: ExcelCase) -> None:
    severity_by_priority = {
        "P0": allure.severity_level.BLOCKER,
        "P1": allure.severity_level.CRITICAL,
        "P2": allure.severity_level.NORMAL,
        "P3": allure.severity_level.MINOR,
    }
    allure.dynamic.title(case.title)
    allure.dynamic.feature(case.module or "未分类模块")
    allure.dynamic.story(case.scenario or case.test_point or case.case_id)
    allure.dynamic.severity(severity_by_priority.get(case.priority, allure.severity_level.NORMAL))
    allure.dynamic.label("case_id", case.case_id)
    if case.tags:
        allure.dynamic.tag(*case.tags)
    description = {
        "用例ID": case.case_id,
        "测试点": case.test_point,
        "优先级": case.priority,
        "前置条件": case.precondition,
        "Excel行号": case.source_row,
        "步骤数": len(case.steps),
    }
    allure.attach(
        json.dumps(description, ensure_ascii=False, indent=2),
        name="Excel用例信息",
        attachment_type=allure.attachment_type.JSON,
    )


def _safety_skip_reason(
    case: ExcelCase,
    config: pytest.Config,
    settings: Settings,
) -> str | None:
    tags = set(case.tags)
    run_seeded = bool(config.getoption("--run-seeded")) or settings.run_seeded
    allow_mutation = bool(config.getoption("--allow-mutation")) or settings.allow_mutation
    allow_destructive = (
        bool(config.getoption("--allow-destructive")) or settings.allow_destructive
    )
    if "requires_seed" in tags and not run_seeded:
        return "需要 --run-seeded 或 YANJIA_RUN_SEEDED=true"
    if "destructive" in tags and not (allow_mutation and allow_destructive):
        return "破坏性用例需要同时授权 --allow-mutation --allow-destructive"
    if "mutating" in tags and not allow_mutation:
        return "写入用例需要 --allow-mutation 或 YANJIA_ALLOW_MUTATION=true"
    return None


def _can_retry_infrastructure(case: ExcelCase, error: BaseException) -> bool:
    tags = set(case.tags)
    safe_to_repeat = (
        "readonly" in tags
        and "mutating" not in tags
        and "destructive" not in tags
        and "no_retry" not in tags
    )
    return safe_to_repeat and is_recoverable_driver_error(error)
