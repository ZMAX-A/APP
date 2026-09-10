from __future__ import annotations

from collections.abc import Generator
from typing import Any

import allure_commons
import pytest

from yanjia_automation.excel.variables import VariableResolver

REDACTOR_KEY: pytest.StashKey[ReportRedactor] = pytest.StashKey()


class ReportRedactor:
    """Redact at reporting boundaries without changing exception/retry semantics."""

    def __init__(self, variables: VariableResolver) -> None:
        self.variables = variables

    @pytest.hookimpl(hookwrapper=True, trylast=True)
    def pytest_runtest_makereport(self) -> Generator[None, Any, None]:
        # Exit before the Allure wrapper consumes captured output and the report.
        report = (yield).get_result()
        if report.longrepr is not None:
            original = report.longreprtext
            redacted = self.variables.redact(original)
            if redacted != original:
                report.longrepr = redacted
        report.sections = [
            (name, self.variables.redact(content)) for name, content in report.sections
        ]

    @allure_commons.hookimpl(tryfirst=True)
    def report_result(self, result: Any) -> None:
        self._redact_item(result)

    @allure_commons.hookimpl(tryfirst=True)
    def report_container(self, container: Any) -> None:
        self._redact_item(container)

    def _redact_item(self, item: Any) -> None:
        # Allure records step/fixture exceptions before pytest produces its report.
        # Scrub these objects before its file logger serializes them as JSON.
        # UUIDs, attachment sources, status and timing must retain their identity.
        for name in ("name", "value", "message", "trace", "description", "descriptionHtml"):
            value = getattr(item, name, None)
            if isinstance(value, str):
                setattr(item, name, self.variables.redact(value))
        details = getattr(item, "statusDetails", None)
        if details is not None:
            self._redact_item(details)
        for name in ("steps", "befores", "afters", "parameters", "labels", "attachments"):
            for child in getattr(item, name, ()):
                self._redact_item(child)


def install_report_redactor(config: pytest.Config, variables: VariableResolver) -> None:
    redactor = ReportRedactor(variables)
    config.stash[REDACTOR_KEY] = redactor
    config.pluginmanager.register(redactor, "yanjia-report-redactor")
    allure_commons.plugin_manager.register(redactor)
    config.add_cleanup(lambda: allure_commons.plugin_manager.unregister(redactor))
