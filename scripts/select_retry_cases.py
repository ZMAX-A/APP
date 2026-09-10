from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class CaseResult:
    case_id: str
    status: str
    tags: frozenset[str]
    restoration_failed: bool
    mutation_blocked: bool
    started_at: int


def read_results(directory: Path) -> list[CaseResult]:
    files = sorted(directory.glob("*-result.json"))
    if not files:
        raise ValueError("No Allure result files; automatic retry cannot be verified.")
    results: list[CaseResult] = []
    for path in files:
        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
            labels = data.get("labels", [])
            case_ids = {str(item["value"]) for item in labels if item["name"] == "case_id"}
            if not case_ids:
                match = re.search(r"(?:^|\[)(TC-[A-Za-z0-9-]+)(?:[ |\]]|$)", data["name"])
                if match:
                    case_ids.add(match.group(1))
            if len(case_ids) != 1:
                raise ValueError("Missing or conflicting case ID")
            case_id = case_ids.pop()
            if not re.fullmatch(r"TC-[A-Za-z0-9-]+", case_id):
                raise ValueError("Invalid case ID")
            status = data["status"]
            if status not in {"passed", "failed", "broken", "skipped", "unknown"}:
                raise ValueError("Invalid status")
            tags = frozenset(
                str(item["value"])
                for item in labels
                if item["name"] == "tag"
            )
            results.append(
                CaseResult(
                    case_id=case_id,
                    status=status,
                    tags=tags,
                    restoration_failed=any(
                        item["name"] == "restoration_failed" and item["value"] == "true"
                        for item in labels
                    ),
                    mutation_blocked=any(
                        item["name"] == "mutation_blocked" and item["value"] == "true"
                        for item in labels
                    ),
                    started_at=int(data.get("start", 0)),
                )
            )
        except (OSError, ValueError, TypeError, KeyError, AttributeError) as error:
            # Do not include arbitrary JSON content or exception text in the launcher log.
            raise ValueError("Invalid Allure evidence; automatic retry is disabled.") from error
    return sorted(results, key=lambda result: result.started_at)


def build_retry_plan(results_dir: Path, retry_results_dir: Path | None = None) -> dict:
    results = read_results(results_dir)
    # The second round is deliberately narrower than the original authorized run.
    # A failed write may already have reached the server before its assertion failed,
    # so only repeatable read-only cases are eligible for an automatic retry.
    failures = [
        result for result in results
        if result.status in {"failed", "broken", "unknown"}
        or (result.status == "skipped" and result.mutation_blocked)
    ]
    failed_ids = list(dict.fromkeys(result.case_id for result in failures))
    restoration_failed = any(result.restoration_failed for result in results)
    retry_ids = list(
        dict.fromkeys(
            result.case_id
            for result in failures
            if _is_safe_to_retry(result.tags)
        )
    )
    blocked_ids = [case_id for case_id in failed_ids if case_id not in retry_ids]
    remaining_ids = failed_ids.copy()
    if retry_results_dir is not None:
        retry_results = read_results(retry_results_dir)
        if {result.case_id for result in retry_results} != set(retry_ids):
            raise ValueError("Retry evidence does not match the selected first-round cases.")
        passed_ids = {
            case_id
            for case_id in retry_ids
            if all(
                result.status == "passed"
                for result in retry_results
                if result.case_id == case_id
            )
        }
        remaining_ids = [case_id for case_id in failed_ids if case_id not in passed_ids]
    return {
        "retryCaseIds": retry_ids,
        "blockedRetryCaseIds": blocked_ids,
        "remainingFailedCaseIds": remaining_ids,
        "restorationFailed": restoration_failed,
    }


def _is_safe_to_retry(tags: frozenset[str]) -> bool:
    return "readonly" in tags and not tags.intersection(
        {"mutating", "destructive", "persistent", "no_retry"}
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Select first-round failures and verify retry results."
    )
    parser.add_argument("--results-dir", type=Path, required=True)
    parser.add_argument("--retry-results-dir", type=Path)
    args = parser.parse_args()
    try:
        plan = build_retry_plan(args.results_dir, args.retry_results_dir)
    except (OSError, ValueError) as error:
        parser.exit(2, f"{error}\n")
    print(json.dumps(plan))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
