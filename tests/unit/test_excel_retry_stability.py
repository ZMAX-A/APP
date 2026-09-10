from __future__ import annotations

from dataclasses import replace
from unittest.mock import ANY, Mock

import pytest
from selenium.common.exceptions import WebDriverException

from tests.excel import test_excel_cases as excel_test_module
from yanjia_automation import driver as driver_module
from yanjia_automation.config import Settings
from yanjia_automation.driver import DriverManager
from yanjia_automation.excel.models import ExcelCase
from yanjia_automation.excel.variables import VariableResolver
from yanjia_automation.flows.customer_mutation import CustomerRestoreError
from yanjia_automation.flows.remark_mutation import RemarkRestoreError
from yanjia_automation.flows.tag_mutation import TagRestoreError


def _settings() -> Settings:
    return Settings(
        appium_server_url="http://127.0.0.1:4723",
        app_package="com.xiaofutech.yanjia_ai",
        app_activity=".activity.SplashActivity",
        udid="device",
        store_name=None,
        no_reset=True,
        run_seeded=False,
        allow_mutation=False,
        allow_destructive=False,
        capture_sensitive_artifacts=False,
        skip_device_initialization=True,
        skip_server_installation=True,
    )


def _readonly_case() -> ExcelCase:
    return ExcelCase(
        source_row=1,
        case_id="TC-RETRY-001",
        module="稳定性",
        scenario="基础设施自动重试",
        test_point="保留运行门禁",
        priority="P1",
        precondition=None,
        input_data=None,
        expected_result=None,
        timeout=10,
        tags=("readonly",),
        enabled=True,
        automation_status="AUTOMATED",
        steps=(),
    )


def _authenticated_case() -> ExcelCase:
    case = _readonly_case()
    return ExcelCase(
        source_row=case.source_row,
        case_id="TC-DETAIL-RETRY-001",
        module="顾客详情",
        scenario="登录态恢复",
        test_point="失败后重新登录",
        priority=case.priority,
        precondition=case.precondition,
        input_data=case.input_data,
        expected_result=case.expected_result,
        timeout=case.timeout,
        tags=("readonly", "requires_auth"),
        enabled=case.enabled,
        automation_status=case.automation_status,
        steps=case.steps,
    )


def test_driver_restart_redeploys_uiautomator2_server(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    created_with: list[Settings] = []
    created_drivers: list[Mock] = []

    def fake_create_driver(settings: Settings) -> Mock:
        created_with.append(settings)
        driver = Mock()
        created_drivers.append(driver)
        return driver

    monkeypatch.setattr(driver_module, "create_driver", fake_create_driver)
    manager = DriverManager(_settings())

    original = manager.get()
    repaired = manager.restart()

    created_drivers[0].quit.assert_called_once_with()
    assert repaired is not original
    assert created_with[0].skip_server_installation is True
    assert created_with[1].skip_server_installation is False
    assert created_with[1].skip_device_initialization is True


def test_driver_restart_can_preserve_skip_server_installation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    created_with: list[Settings] = []

    def fake_create_driver(settings: Settings) -> Mock:
        created_with.append(settings)
        return Mock()

    monkeypatch.setattr(driver_module, "create_driver", fake_create_driver)
    manager = DriverManager(_settings())
    manager.get()

    manager.restart(repair_uiautomator2=False)

    assert created_with[-1].skip_server_installation is True


def test_infrastructure_retry_preserves_effective_runner_settings(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    observed_settings: list[Settings] = []

    class FakeRunner:
        def __init__(self, _driver: object, settings: Settings, _resolver: object) -> None:
            observed_settings.append(settings)

        def run(self, _case: ExcelCase) -> None:
            if len(observed_settings) == 1:
                raise WebDriverException("socket hang up")

        def attach_failure_evidence(self) -> None:
            raise AssertionError("failure evidence should not run after a successful retry")

    manager = Mock()
    manager.get.return_value = object()
    manager.restart.return_value = object()
    request = Mock()
    request.getfixturevalue.return_value = manager
    config = Mock()
    config.stash = pytest.Stash()
    config.getoption.side_effect = lambda name: name in {
        "--run-seeded",
        "--allow-mutation",
        "--allow-destructive",
        "--customer-preflight-verified",
    }
    resolver = Mock()
    resolver.redact.side_effect = str
    writer = Mock()

    monkeypatch.setattr(excel_test_module, "ExcelCaseRunner", FakeRunner)
    monkeypatch.setattr(excel_test_module, "_configure_allure", lambda _case: None)
    monkeypatch.setattr(excel_test_module.allure, "attach", Mock())
    monkeypatch.setattr(excel_test_module.allure.dynamic, "label", Mock())

    excel_test_module.test_excel_case(
        _readonly_case(),
        request,
        config,
        _settings(),
        resolver,
        writer,
    )

    assert len(observed_settings) == 2
    for effective in observed_settings:
        assert effective.run_seeded is True
        assert effective.allow_mutation is True
        assert effective.allow_destructive is True
        assert effective.customer_preflight_verified is True
    manager.restart.assert_called_once_with()
    writer.record.assert_called_once()


def test_failed_authenticated_case_recovers_login_for_following_case(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    failure = AssertionError("customer result was not visible")

    class FakeRunner:
        def __init__(self, _driver: object, _settings: Settings, _resolver: object) -> None:
            pass

        def run(self, _case: ExcelCase) -> None:
            raise failure

        def attach_failure_evidence(self) -> None:
            pass

    driver = object()
    manager = Mock()
    manager.get.return_value = driver
    request = Mock()
    request.getfixturevalue.return_value = manager
    config = Mock()
    config.stash = pytest.Stash()
    config.getoption.return_value = False
    resolver = Mock()
    resolver.redact.side_effect = str
    writer = Mock()
    recover = Mock(return_value=True)

    monkeypatch.setattr(excel_test_module, "ExcelCaseRunner", FakeRunner)
    monkeypatch.setattr(excel_test_module, "_configure_allure", lambda _case: None)
    monkeypatch.setattr(excel_test_module, "recover_login_if_needed", recover)
    monkeypatch.setattr(excel_test_module.allure, "attach", Mock())
    monkeypatch.setattr(excel_test_module.allure.dynamic, "label", Mock())

    with pytest.raises(AssertionError, match="customer result was not visible"):
        excel_test_module.test_excel_case(
            _authenticated_case(),
            request,
            config,
            _settings(),
            resolver,
            writer,
        )

    recover.assert_called_once_with(driver, ANY)
    excel_test_module.allure.dynamic.label.assert_called_with(
        "authentication_recovery", "succeeded"
    )
    writer.record.assert_called_once()
    assert writer.record.call_args.args[0].status == "FAIL"


def _runtime(monkeypatch: pytest.MonkeyPatch):
    manager = Mock()
    request = Mock()
    request.getfixturevalue.return_value = manager
    config = Mock()
    config.stash = pytest.Stash()
    config.getoption.side_effect = lambda name: name != "--readonly-retry"
    writer = Mock()
    runner = Mock()
    recover = Mock()
    monkeypatch.setattr(excel_test_module, "ExcelCaseRunner", Mock(return_value=runner))
    monkeypatch.setattr(excel_test_module, "recover_login_if_needed", recover)
    monkeypatch.setattr(excel_test_module, "_configure_allure", lambda _case: None)
    monkeypatch.setattr(excel_test_module.allure, "attach", Mock())
    monkeypatch.setattr(excel_test_module.allure.dynamic, "label", Mock())
    variables = VariableResolver({}, sensitive_values={"fake-private-value"})

    def run(case: ExcelCase) -> None:
        excel_test_module.test_excel_case(
            case, request, config, _settings(), variables, writer
        )

    return run, manager, config, writer, runner, recover


@pytest.mark.parametrize("error_type", [CustomerRestoreError, TagRestoreError, RemarkRestoreError])
def test_restore_failure_blocks_later_writes_before_driver_creation(
    monkeypatch: pytest.MonkeyPatch, error_type: type[Exception]
) -> None:
    run, manager, config, writer, runner, recover = _runtime(monkeypatch)
    mutation = replace(_readonly_case(), tags=("mutating", "requires_auth"))
    runner.run.side_effect = [error_type("restore failed"), None]
    with pytest.raises(error_type):
        run(mutation)
    for tags in (("mutating",), ("destructive",), ("readonly", "persistent"), ()):
        with pytest.raises(pytest.skip.Exception, match="数据恢复失败"):
            run(replace(mutation, tags=tags))
    assert manager.get.call_count == 1
    recover.assert_not_called()
    run(_readonly_case())
    assert manager.get.call_count == 2
    assert [call.args[0].status for call in writer.record.call_args_list] == [
        "ERROR", "SKIP", "SKIP", "SKIP", "SKIP", "PASS"
    ]
    assert config.stash[excel_test_module.MUTATION_BLOCKED_KEY]
    assert any(
        call.args == ("mutation_blocked", "true")
        for call in excel_test_module.allure.dynamic.label.call_args_list
    )


def test_new_round_retries_mutation_after_previous_round_restore_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    run, _manager, config, writer, runner, _recover = _runtime(monkeypatch)
    mutation = replace(_readonly_case(), tags=("mutating",))
    runner.run.side_effect = [CustomerRestoreError("restore failed"), None]
    with pytest.raises(CustomerRestoreError):
        run(mutation)
    # Each launcher round starts a new pytest process and therefore a new stash.
    config.stash = pytest.Stash()
    run(mutation)
    assert [call.args[0].status for call in writer.record.call_args_list] == ["ERROR", "PASS"]


def test_initial_driver_failure_can_recover_once_for_readonly_case(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    run, manager, _config, writer, runner, _recover = _runtime(monkeypatch)
    manager.get.side_effect = ConnectionError("connection refused")
    run(_readonly_case())
    manager.restart.assert_called_once_with()
    runner.run.assert_called_once()
    assert writer.record.call_args.args[0].status == "PASS"


@pytest.mark.parametrize("tags", [("readonly",), ("mutating",), ("readonly", "persistent")])
def test_driver_failures_are_recorded_without_using_absent_or_stale_driver(
    monkeypatch: pytest.MonkeyPatch, tags: tuple[str, ...]
) -> None:
    run, manager, _config, writer, runner, recover = _runtime(monkeypatch)
    manager.get.side_effect = ConnectionError("connection refused fake-private-value")
    manager.restart.side_effect = ConnectionError("connection refused fake-private-value")
    with pytest.raises(ConnectionError):
        run(replace(_readonly_case(), tags=tags))
    assert manager.restart.call_count == (1 if tags == ("readonly",) else 0)
    runner.run.assert_not_called()
    runner.attach_failure_evidence.assert_not_called()
    recover.assert_not_called()
    result = writer.record.call_args.args[0]
    assert result.status == "ERROR"
    assert "fake-private-value" not in result.error_message
    assert "ConnectionError" in result.error_message


def test_replacement_driver_failure_does_not_use_old_runner_for_evidence(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    run, manager, _config, writer, runner, recover = _runtime(monkeypatch)
    runner.run.side_effect = WebDriverException("socket hang up")
    manager.restart.side_effect = ConnectionError("connection refused")
    with pytest.raises(ConnectionError):
        run(_authenticated_case())
    runner.attach_failure_evidence.assert_not_called()
    recover.assert_not_called()
    assert writer.record.call_args.args[0].status == "ERROR"


@pytest.mark.parametrize(
    "tags",
    [("mutating",), ("destructive",), ("readonly", "persistent"), ("readonly", "no_retry"), ()],
)
def test_readonly_retry_gate_overrides_all_write_authorizations(
    monkeypatch: pytest.MonkeyPatch, tags: tuple[str, ...]
) -> None:
    run, manager, config, writer, _runner, _recover = _runtime(monkeypatch)
    config.getoption.side_effect = lambda _name: True
    with pytest.raises(pytest.skip.Exception, match="自动复跑"):
        run(replace(_readonly_case(), tags=tags))
    manager.get.assert_not_called()
    assert writer.record.call_args.args[0].status == "SKIP"
