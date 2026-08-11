from __future__ import annotations

from dataclasses import dataclass

from yanjia_automation.excel.locators import LocatorFormatError, parse_locator_candidates
from yanjia_automation.excel.models import ExcelCase, ExcelStep
from yanjia_automation.excel.runner import (
    ASSERTION_NAMES,
    LOCATOR_REQUIRED_ACTIONS,
    LOCATOR_REQUIRED_ASSERTIONS,
    SUPPORTED_ACTIONS,
)
from yanjia_automation.excel.variables import VARIABLE_PATTERN, VariableResolver


@dataclass(frozen=True)
class ValidationIssue:
    level: str
    location: str
    message: str

    def format(self) -> str:
        return f"[{self.level}] {self.location}: {self.message}"


@dataclass(frozen=True)
class ValidationReport:
    cases: int
    steps: int
    issues: tuple[ValidationIssue, ...]

    @property
    def errors(self) -> tuple[ValidationIssue, ...]:
        return tuple(issue for issue in self.issues if issue.level == "ERROR")

    @property
    def warnings(self) -> tuple[ValidationIssue, ...]:
        return tuple(issue for issue in self.issues if issue.level == "WARN")

    @property
    def is_valid(self) -> bool:
        return not self.errors


def validate_cases(
    cases: list[ExcelCase],
    variables: VariableResolver,
) -> ValidationReport:
    issues: list[ValidationIssue] = []
    for case in cases:
        case_location = f"主表第{case.source_row}行/{case.case_id}"
        if any(
            step.action == "input" and step.input_value is None
            for step in case.steps
        ):
            _validate_variables(case.input_data, case_location, variables, issues)
        if "destructive" in case.tags:
            issues.append(
                ValidationIssue(
                    "WARN",
                    case_location,
                    "用例带有 destructive 标签，默认不会执行",
                )
            )
        elif "mutating" in case.tags:
            issues.append(
                ValidationIssue(
                    "WARN",
                    case_location,
                    "用例带有 mutating 标签，需要显式授权才会执行",
                )
            )

        for step in case.steps:
            _validate_step(step, variables, issues)

    return ValidationReport(
        cases=len(cases),
        steps=sum(len(case.steps) for case in cases),
        issues=tuple(issues),
    )


def _validate_step(
    step: ExcelStep,
    variables: VariableResolver,
    issues: list[ValidationIssue],
) -> None:
    location = f"步骤表第{step.source_row}行/{step.case_id}/步骤{step.order}"
    action = step.action
    if action not in SUPPORTED_ACTIONS and action not in ASSERTION_NAMES:
        issues.append(
            ValidationIssue("ERROR", location, f"不支持的操作类型：{action or '<空>'}")
        )

    if action in {"assert", "verify"} and not step.assertion:
        issues.append(
            ValidationIssue("ERROR", location, "assert/verify 必须填写断言类型")
        )
    if step.assertion and step.assertion not in ASSERTION_NAMES:
        issues.append(
            ValidationIssue("ERROR", location, f"不支持的断言类型：{step.assertion}")
        )
    if action in ASSERTION_NAMES and step.assertion:
        issues.append(
            ValidationIssue(
                "ERROR",
                location,
                "操作类型已直接填写断言名时，断言类型列应留空",
            )
        )

    effective_assertion = (
        step.assertion
        if action in {"assert", "verify"}
        else action if action in ASSERTION_NAMES else None
    )
    locator_required = (
        action in LOCATOR_REQUIRED_ACTIONS
        or effective_assertion in LOCATOR_REQUIRED_ASSERTIONS
    )
    if locator_required and not step.locator:
        issues.append(ValidationIssue("ERROR", location, "当前关键字必须填写元素定位器"))
    elif step.locator:
        try:
            parse_locator_candidates(step.locator)
        except LocatorFormatError as error:
            issues.append(ValidationIssue("ERROR", location, str(error)))

    for value in (step.locator, step.input_value, step.expected):
        _validate_variables(value, location, variables, issues)

    if step.continue_on_failure:
        issues.append(
            ValidationIssue(
                "WARN",
                location,
                "已启用失败继续；最终仍会汇总失败，请确认这是预期行为",
            )
        )


def _validate_variables(
    value: str | None,
    location: str,
    variables: VariableResolver,
    issues: list[ValidationIssue],
) -> None:
    if not value:
        return
    missing = sorted(
        {
            match.group(1)
            for match in VARIABLE_PATTERN.finditer(value)
            if match.group(1) not in variables.available_names
        }
    )
    if missing:
        issues.append(
            ValidationIssue(
                "ERROR",
                location,
                f"引用了未配置的变量：{', '.join(missing)}",
            )
        )
