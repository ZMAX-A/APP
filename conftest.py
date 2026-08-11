from __future__ import annotations

import pytest


def pytest_addoption(parser: pytest.Parser) -> None:
    safety = parser.getgroup("yanjia safety")
    safety.addoption(
        "--run-seeded",
        action="store_true",
        default=False,
        help="Run read-only tests that depend on controlled customer/case/image seed data.",
    )
    safety.addoption(
        "--allow-mutation",
        action="store_true",
        default=False,
        help="Allow tests that change server-side data.",
    )
    safety.addoption(
        "--allow-destructive",
        action="store_true",
        default=False,
        help="Allow tests that delete data. Mutation authorization is also required.",
    )

    excel = parser.getgroup("yanjia excel")
    excel.addoption(
        "--excel-file",
        default="test_case.xlsx",
        help="Excel workbook used as the test-case and step data source.",
    )
    excel.addoption(
        "--excel-output",
        default=None,
        help="Optional result workbook. By default results are written back in place.",
    )
    excel.addoption(
        "--excel-case-id",
        default=None,
        help="Comma-separated case IDs or wildcard patterns, for example TC-HOME-*.",
    )
    excel.addoption(
        "--excel-tags",
        default=None,
        help="Comma-separated tags; a case must contain every requested tag.",
    )
    excel.addoption(
        "--excel-run-id",
        default=None,
        help="Stable run identifier written to Excel and Allure metadata.",
    )
    excel.addoption(
        "--allure-report-dir",
        default=None,
        help="Generated Allure report directory written back to Excel.",
    )
    excel.addoption(
        "--no-excel-writeback",
        action="store_false",
        dest="excel_writeback",
        default=True,
        help="Run cases without changing the Excel workbook.",
    )
