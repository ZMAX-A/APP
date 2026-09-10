from __future__ import annotations

import json
import time
from dataclasses import replace
from datetime import datetime

import allure
import pytest
from appium.webdriver.webdriver import WebDriver

from yanjia_automation.config import Settings
from yanjia_automation.driver import DriverManager, is_recoverable_driver_error
from yanjia_automation.excel.models import ExcelCase
from yanjia_automation.excel.results import ExcelResult, ExcelResultWriter
from yanjia_automation.excel.runner import ExcelCaseRunner
from yanjia_automation.excel.safety import (
    execution_skip_reason,
    is_readonly,
    is_safe_to_retry,
)
from yanjia_automation.excel.variables import VariableResolver
from yanjia_automation.flows.customer_mutation import CustomerRestoreError
from yanjia_automation.flows.recovery import recover_login_if_needed
from yanjia_automation.flows.remark_mutation import RemarkRestoreError
from yanjia_automation.flows.tag_mutation import TagRestoreError

MUTATION_BLOCKED_KEY: pytest.StashKey[str] = pytest.StashKey()
RESTORE_ERRORS = (CustomerRestoreError, TagRestoreError, RemarkRestoreError)


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
        if pytestconfig.stash.get(MUTATION_BLOCKED_KEY, None) and not is_readonly(excel_case.tags):
            allure.dynamic.label("mutation_blocked", "true")
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

    runner_settings = replace(
        settings,
        run_seeded=bool(pytestconfig.getoption("--run-seeded")) or settings.run_seeded,
        allow_mutation=(
            bool(pytestconfig.getoption("--allow-mutation")) or settings.allow_mutation
        ),
        allow_destructive=(
            bool(pytestconfig.getoption("--allow-destructive"))
            or settings.allow_destructive
        ),
        customer_preflight_verified=(
            bool(pytestconfig.getoption("--customer-preflight-verified"))
            or settings.customer_preflight_verified
        ),
    )
    driver: WebDriver | None = None
    runner: ExcelCaseRunner | None = None
    try:
        driver_manager: DriverManager = request.getfixturevalue("driver_manager")
        try:
            driver = driver_manager.get()
            runner = ExcelCaseRunner(driver, runner_settings, excel_variables)
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
            # Do not collect evidence or recover authentication using a stale driver
            # if creating the replacement session itself fails.
            driver = None
            runner = None
            driver = driver_manager.restart()
            runner = ExcelCaseRunner(driver, runner_settings, excel_variables)
            runner.run(excel_case)
    except Exception as error:
        if isinstance(error, RESTORE_ERRORS):
            pytestconfig.stash[MUTATION_BLOCKED_KEY] = (
                f"{excel_case.case_id} 数据恢复失败；本轮后续写入已停止，请先核实并恢复测试数据"
            )
            allure.dynamic.label("restoration_failed", "true")
        finished_at = datetime.now()
        safe_error = excel_variables.redact(error)
        try:
            if runner is not None:
                runner.attach_failure_evidence()
        except Exception as evidence_error:
            allure.attach(
                excel_variables.redact(evidence_error),
                name="失败证据采集异常",
                attachment_type=allure.attachment_type.TEXT,
            )
        if driver is not None and not isinstance(error, RESTORE_ERRORS):
            _recover_authentication_for_next_case(
                excel_case,
                driver,
                runner_settings,
                excel_variables,
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
    blocked = config.stash.get(MUTATION_BLOCKED_KEY, None)
    if blocked and not is_readonly(case.tags):
        return blocked
    if config.getoption("--readonly-retry") and not is_safe_to_retry(case.tags):
        return "自动复跑仅允许可重复执行的只读用例；写入、删除、持久化和 no_retry 用例已阻止"
    run_seeded = bool(config.getoption("--run-seeded")) or settings.run_seeded
    allow_mutation = bool(config.getoption("--allow-mutation")) or settings.allow_mutation
    allow_destructive = (
        bool(config.getoption("--allow-destructive")) or settings.allow_destructive
    )
    customer_preflight_verified = (
        bool(config.getoption("--customer-preflight-verified"))
        or settings.customer_preflight_verified
    )
    return execution_skip_reason(
        case,
        run_seeded=run_seeded,
        allow_mutation=allow_mutation,
        allow_destructive=allow_destructive,
        mutation_customer_query=settings.mutation_customer_query,
        customer_preflight_verified=customer_preflight_verified,
    )


def _can_retry_infrastructure(case: ExcelCase, error: BaseException) -> bool:
    return is_safe_to_retry(case.tags) and is_recoverable_driver_error(error)


def _recover_authentication_for_next_case(
    case: ExcelCase,
    driver: WebDriver,
    settings: Settings,
    variables: VariableResolver,
) -> bool:
    tags = set(case.tags)
    if "requires_auth" not in tags or case.case_id.startswith("TC-LOGIN-"):
        return False
    try:
        recovered = recover_login_if_needed(driver, settings)
    except Exception as recovery_error:
        allure.attach(
            variables.redact(recovery_error),
            name="登录态恢复失败",
            attachment_type=allure.attachment_type.TEXT,
        )
        allure.dynamic.label("authentication_recovery", "failed")
        return False
    if not recovered:
        return False
    allure.attach(
        "检测到应用意外返回登录页，已重新登录并恢复到首页；本用例保留原始失败，复跑仍受安全门禁约束。",
        name="登录态已自动恢复",
        attachment_type=allure.attachment_type.TEXT,
    )
    allure.dynamic.label("authentication_recovery", "succeeded")
    return True
