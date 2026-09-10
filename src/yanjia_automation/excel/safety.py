from __future__ import annotations

from collections.abc import Iterable, Mapping
from types import MappingProxyType

from yanjia_automation.excel.models import ExcelCase

CUSTOMER_MUTATION_TAIL_ACTIONS: Mapping[str, tuple[str, ...]] = MappingProxyType(
    {
        "TC-DETAIL-001": ("clear", "click", "assert"),
        "TC-DETAIL-003": ("input", "click", "assert"),
        "TC-DETAIL-004": ("click", "click", "click", "click", "assert"),
        "TC-DETAIL-005": ("click", "click", "click", "click", "assert"),
        "TC-DETAIL-007": ("select_birthday", "click", "click", "assert"),
        "TC-DETAIL-008": ("clear", "click", "click", "assert"),
        "TC-DETAIL-009": ("input", "click", "click", "assert"),
        "TC-DETAIL-010": ("input", "click", "click", "assert"),
        "TC-DETAIL-011": ("clear", "click", "assert"),
        "TC-DETAIL-013": (
            "scroll_profile_remark",
            "clear",
            "click",
            "click",
            "scroll_profile_remark",
            "assert",
        ),
        "TC-DETAIL-014": (
            "scroll_profile_remark",
            "input",
            "click",
            "click",
            "scroll_profile_remark",
            "assert",
        ),
        "TC-DETAIL-015": ("click", "click", "click", "click", "assert"),
        "TC-DETAIL-026": ("input", "click", "back", "assert", "assert"),
        "TC-DETAIL-027": ("input", "click", "back", "assert", "assert"),
    }
)
CUSTOMER_MUTATION_CASE_IDS = frozenset(CUSTOMER_MUTATION_TAIL_ACTIONS)
TAG_MUTATION_TAIL_ACTIONS: Mapping[str, tuple[str, ...]] = MappingProxyType(
    {
        "TC-DETAIL-016": ("click", "assert", "input", "click", "assert"),
        "TC-DETAIL-018": ("click", "assert", "input", "click", "assert"),
        "TC-DETAIL-020": (
            "click",
            "assert",
            "input",
            "click",
            "assert",
            "delete_current_run_tag",
            "assert",
        ),
        "TC-DETAIL-031": ("click", "assert", "input", "click", "assert"),
        "TC-DETAIL-033": ("click", "assert", "input", "click", "assert"),
    }
)
TAG_MUTATION_CASE_IDS = frozenset(TAG_MUTATION_TAIL_ACTIONS)
TAG_MUTATION_SCOPES: Mapping[str, str] = MappingProxyType(
    {
        "TC-DETAIL-016": "image",
        "TC-DETAIL-018": "image",
        "TC-DETAIL-020": "image",
        "TC-DETAIL-031": "consultation",
        "TC-DETAIL-033": "consultation",
    }
)
REMARK_MUTATION_TAIL_ACTIONS: Mapping[str, tuple[str, ...]] = MappingProxyType(
    {
        "TC-DETAIL-012": ("assert", "input", "assert", "click", "assert"),
        "TC-DETAIL-021": ("assert", "input", "assert", "click", "assert"),
        "TC-DETAIL-022": ("assert", "input", "assert", "click", "assert"),
        "TC-DETAIL-035": ("assert", "input", "assert", "click", "assert"),
        "TC-DETAIL-037": (
            "assert",
            "input",
            "assert",
            "click",
            "assert",
            "delete_current_run_remark",
            "assert",
        ),
    }
)
REMARK_MUTATION_CASE_IDS = frozenset(REMARK_MUTATION_TAIL_ACTIONS)
REMARK_MUTATION_SCOPES: Mapping[str, str] = MappingProxyType(
    {
        "TC-DETAIL-012": "image",
        "TC-DETAIL-021": "image",
        "TC-DETAIL-022": "image",
        "TC-DETAIL-035": "consultation",
        "TC-DETAIL-037": "consultation",
    }
)
DEDICATED_CUSTOMER_TARGET_CASE_IDS = CUSTOMER_MUTATION_CASE_IDS | {
    "TC-DETAIL-017",
    "TC-DETAIL-019",
} | TAG_MUTATION_CASE_IDS | REMARK_MUTATION_CASE_IDS


def is_readonly(tags: Iterable[str]) -> bool:
    values = set(tags)
    return "readonly" in values and not values.intersection(
        {"mutating", "destructive", "persistent"}
    )


def is_safe_to_retry(tags: Iterable[str]) -> bool:
    values = set(tags)
    return is_readonly(values) and "no_retry" not in values


def execution_skip_reason(
    case: ExcelCase,
    *,
    run_seeded: bool,
    allow_mutation: bool,
    allow_destructive: bool,
    mutation_customer_query: str | None,
    customer_preflight_verified: bool,
) -> str | None:
    """Return a fail-closed safety reason before an Excel case gets a driver."""

    tags = set(case.tags)
    if "requires_seed" in tags and not run_seeded:
        return "需要 --run-seeded 或 YANJIA_RUN_SEEDED=true"
    if "destructive" in tags and not (allow_mutation and allow_destructive):
        return "破坏性用例需要同时授权 --allow-mutation --allow-destructive"
    if "mutating" in tags and not allow_mutation:
        return "写入用例需要 --allow-mutation 或 YANJIA_ALLOW_MUTATION=true"

    if case.case_id in DEDICATED_CUSTOMER_TARGET_CASE_IDS:
        if not mutation_customer_query or not mutation_customer_query.strip():
            return "专用顾客用例需要配置 YANJIA_MUTATION_CUSTOMER_QUERY"
    if case.case_id in (
        CUSTOMER_MUTATION_CASE_IDS | TAG_MUTATION_CASE_IDS | REMARK_MUTATION_CASE_IDS
    ):
        if not customer_preflight_verified:
            return "专用顾客真机只读预检尚未授权本次运行，写入用例保持闭锁"
    return None
