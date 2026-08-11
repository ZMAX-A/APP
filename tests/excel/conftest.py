from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from yanjia_automation.config import Settings
from yanjia_automation.excel.models import ExcelCase
from yanjia_automation.excel.results import ExcelResultWriter
from yanjia_automation.excel.variables import VariableResolver
from yanjia_automation.excel.workbook import (
    ExcelCaseRepository,
    WorkbookFormatError,
    parse_csv_option,
)

EXCEL_CASES_KEY: pytest.StashKey[tuple[ExcelCase, ...]] = pytest.StashKey()


def pytest_generate_tests(metafunc: pytest.Metafunc) -> None:
    if "excel_case" not in metafunc.fixturenames:
        return

    workbook_path = _resolve_path(
        metafunc.config,
        _option_text(metafunc.config, "--excel-file", default="test_case.xlsx")
        or "test_case.xlsx",
    )
    repository = ExcelCaseRepository(workbook_path)
    try:
        cases = repository.load_cases(
            case_patterns=parse_csv_option(
                _option_text(metafunc.config, "--excel-case-id")
            ),
            required_tags=parse_csv_option(_option_text(metafunc.config, "--excel-tags")),
        )
    except WorkbookFormatError as error:
        raise pytest.UsageError(str(error)) from error
    if not cases:
        raise pytest.UsageError("Excel筛选结果为空：没有找到已启用且符合条件的用例")

    metafunc.config.stash[EXCEL_CASES_KEY] = tuple(cases)
    parameters = [
        pytest.param(case, id=case.case_id, marks=pytest.mark.excel_driven)
        for case in cases
    ]
    metafunc.parametrize("excel_case", parameters)


@pytest.fixture(scope="session")
def excel_run_id(pytestconfig: pytest.Config) -> str:
    configured = _option_text(pytestconfig, "--excel-run-id")
    return str(configured or datetime.now().strftime("%Y%m%d_%H%M%S"))


@pytest.fixture(scope="session")
def excel_variables(settings: Settings, excel_run_id: str) -> VariableResolver:
    return VariableResolver.from_settings(settings, run_id=excel_run_id)


@pytest.fixture(scope="session")
def excel_result_writer(
    pytestconfig: pytest.Config,
    excel_run_id: str,
) -> ExcelResultWriter:
    source = _resolve_path(
        pytestconfig,
        _option_text(pytestconfig, "--excel-file", default="test_case.xlsx")
        or "test_case.xlsx",
    )
    output_option = _option_text(pytestconfig, "--excel-output")
    output = _resolve_path(pytestconfig, output_option) if output_option else None
    report_option = _option_text(pytestconfig, "--allure-report-dir")
    report_dir = (
        _resolve_path(pytestconfig, report_option)
        if report_option
        else Path(pytestconfig.rootpath) / "reports" / "allure-report" / excel_run_id
    )
    writer = ExcelResultWriter(
        source_path=source,
        output_path=output,
        run_id=excel_run_id,
        allure_report_dir=report_dir,
        enabled=bool(pytestconfig.getoption("excel_writeback")),
    )
    cases = pytestconfig.stash.get(EXCEL_CASES_KEY, ())
    writer.prepare(tuple(case.case_id for case in cases))
    return writer


def _resolve_path(config: pytest.Config, value: str) -> Path:
    path = Path(value)
    return path.resolve() if path.is_absolute() else (Path(config.rootpath) / path).resolve()


def _option_text(
    config: pytest.Config,
    name: str,
    *,
    default: str | None = None,
) -> str | None:
    value: object = config.getoption(name)
    return value if isinstance(value, str) else default
