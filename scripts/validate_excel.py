from __future__ import annotations

import argparse
from pathlib import Path

from yanjia_automation.config import load_settings
from yanjia_automation.excel.validation import validate_cases
from yanjia_automation.excel.variables import VariableResolver
from yanjia_automation.excel.workbook import (
    ExcelCaseRepository,
    WorkbookFormatError,
    parse_csv_option,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate Excel cases without a device.")
    parser.add_argument("workbook", type=Path, nargs="?", default=Path("test_case.xlsx"))
    parser.add_argument(
        "--case-id",
        default=None,
        help="Comma-separated case IDs or wildcard patterns for targeted validation.",
    )
    args = parser.parse_args()
    workbook_path = args.workbook.resolve()

    try:
        cases = ExcelCaseRepository(workbook_path).load_cases(
            case_patterns=parse_csv_option(args.case_id)
        )
        settings = load_settings()
        variables = VariableResolver.from_settings(
            settings,
            run_id="VALIDATION",
            require_credentials=False,
        )
        report = validate_cases(cases, variables)
    except (WorkbookFormatError, OSError, RuntimeError, ValueError) as error:
        print(f"[ERROR] Excel校验无法完成：{error}")
        return 2

    for issue in report.issues:
        print(issue.format())
    print(
        "校验完成："
        f"{report.cases} 条启用用例，{report.steps} 个步骤，"
        f"{len(report.errors)} 个错误，{len(report.warnings)} 个警告"
    )
    return 0 if report.is_valid else 2


if __name__ == "__main__":
    raise SystemExit(main())
