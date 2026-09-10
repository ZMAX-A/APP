from __future__ import annotations

import json
from pathlib import Path

import pytest

from scripts.select_retry_cases import build_retry_plan


def write_result(
    directory: Path,
    case_id: str,
    status: str,
    tags: tuple[str, ...] = ("readonly",),
    *,
    restoration_failed: bool = False,
    mutation_blocked: bool = False,
) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    labels = [{"name": "case_id", "value": case_id}]
    labels.extend({"name": "tag", "value": tag} for tag in tags)
    if restoration_failed:
        labels.append({"name": "restoration_failed", "value": "true"})
    if mutation_blocked:
        labels.append({"name": "mutation_blocked", "value": "true"})
    (directory / f"{case_id}-result.json").write_text(
        json.dumps({"name": case_id, "status": status, "labels": labels, "start": 1}),
        encoding="utf-8",
    )


def test_retry_selection_blocks_persistent_writes_and_uses_readonly_result(
    tmp_path: Path,
) -> None:
    first, retry = tmp_path / "first", tmp_path / "retry"
    write_result(first, "TC-HOME-007", "failed")
    write_result(first, "TC-DETAIL-030", "broken", ("mutating", "persistent"))
    write_result(retry, "TC-HOME-007", "passed")
    plan = build_retry_plan(first, retry)
    assert plan["retryCaseIds"] == ["TC-HOME-007"]
    assert plan["blockedRetryCaseIds"] == ["TC-DETAIL-030"]
    assert plan["remainingFailedCaseIds"] == ["TC-DETAIL-030"]


@pytest.mark.parametrize(
    "tags", [("destructive",), ("readonly", "persistent"), ("readonly", "no_retry"), ()]
)
def test_outer_second_round_blocks_non_repeatable_failure_categories(
    tmp_path: Path, tags: tuple[str, ...]
) -> None:
    write_result(tmp_path, "TC-TEST-001", "broken", tags)
    plan = build_retry_plan(tmp_path)
    assert plan["retryCaseIds"] == []
    assert plan["blockedRetryCaseIds"] == ["TC-TEST-001"]
    assert plan["remainingFailedCaseIds"] == ["TC-TEST-001"]


def test_restoration_failure_only_allows_independent_readonly_retry(tmp_path: Path) -> None:
    write_result(tmp_path, "TC-HOME-007", "failed")
    write_result(tmp_path, "TC-DETAIL-004", "broken", ("mutating",), restoration_failed=True)
    plan = build_retry_plan(tmp_path)
    assert plan["restorationFailed"] is True
    assert plan["retryCaseIds"] == ["TC-HOME-007"]
    assert plan["blockedRetryCaseIds"] == ["TC-DETAIL-004"]
    assert len(plan["remainingFailedCaseIds"]) == 2


def test_restoration_related_mutation_skip_is_not_retried(tmp_path: Path) -> None:
    write_result(tmp_path, "TC-DETAIL-005", "skipped", ("mutating",), mutation_blocked=True)
    write_result(tmp_path, "TC-DETAIL-006", "skipped", ("mutating",))
    write_result(tmp_path, "TC-HOME-001", "passed")
    plan = build_retry_plan(tmp_path)
    assert plan["retryCaseIds"] == []
    assert plan["blockedRetryCaseIds"] == ["TC-DETAIL-005"]
    assert plan["remainingFailedCaseIds"] == ["TC-DETAIL-005"]


@pytest.mark.parametrize("status", ["passed", "skipped", "broken", "failed"])
def test_retry_requires_actual_pass_for_every_selected_case(tmp_path: Path, status: str) -> None:
    first, retry = tmp_path / "first", tmp_path / "retry"
    write_result(first, "TC-HOME-007", "failed")
    write_result(retry, "TC-HOME-007", status)
    assert build_retry_plan(first, retry)["remainingFailedCaseIds"] == (
        [] if status == "passed" else ["TC-HOME-007"]
    )


def test_missing_retry_case_or_corrupt_evidence_fails_closed(tmp_path: Path) -> None:
    first, retry = tmp_path / "first", tmp_path / "retry"
    write_result(first, "TC-HOME-007", "failed")
    write_result(first, "TC-HOME-006", "failed")
    write_result(retry, "TC-HOME-007", "passed")
    with pytest.raises(ValueError, match="does not match"):
        build_retry_plan(first, retry)
    (first / "corrupt-result.json").write_text("broken-json", encoding="utf-8")
    with pytest.raises(ValueError, match="Invalid Allure evidence"):
        build_retry_plan(first)
