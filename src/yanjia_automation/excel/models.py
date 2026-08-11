from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ExcelStep:
    source_row: int
    case_id: str
    order: int
    name: str
    action: str
    locator: str | None
    input_value: str | None
    assertion: str | None
    expected: str | None
    timeout: float
    element_index: int
    continue_on_failure: bool
    enabled: bool
    note: str | None


@dataclass(frozen=True)
class ExcelCase:
    source_row: int
    case_id: str
    module: str
    scenario: str
    test_point: str
    priority: str
    precondition: str | None
    input_data: str | None
    expected_result: str | None
    timeout: float
    tags: tuple[str, ...]
    enabled: bool
    automation_status: str
    steps: tuple[ExcelStep, ...]

    @property
    def title(self) -> str:
        parts = [part for part in (self.case_id, self.module, self.scenario) if part]
        return " | ".join(parts)
