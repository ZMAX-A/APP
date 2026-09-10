"""Central Android case catalog and coverage helpers.

The workbook remains the source of truth for case steps and execution flags.
This module owns the smaller piece that must stay consistent across the
single-sheet migration, the canonical step-sheet normalizer, and coverage
reports: which Android cases are registered, which execution source supports
them, and their safety/category metadata.
"""

from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from types import MappingProxyType

SINGLE_SHEET_SOURCE = "single_sheet"
STEP_SHEET_SOURCE = "step_sheet"
CATALOG_SOURCES = frozenset({SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE})


@dataclass(frozen=True, slots=True)
class AndroidCaseDefinition:
    """Metadata shared by Android execution adapters."""

    case_id: str
    category: str
    sources: frozenset[str]
    tags: tuple[str, ...]
    mutability: str = "readonly"
    requires_seed: bool = False
    flow: str = "contract"
    result_state: str = "positive"
    filter_value: str | None = None
    filter_locator: str | None = None
    filter_options: tuple[tuple[str, str], ...] = ()
    blocked_reason: str | None = None

    def supports(self, source: str) -> bool:
        """Return whether this case is implemented by ``source``."""

        return source in self.sources

    def automation_status_for(self, source: str) -> str:
        """Return the status written by the selected execution adapter."""

        if source not in CATALOG_SOURCES:
            raise ValueError(f"Unsupported Android catalog source: {source}")
        if not self.supports(source):
            raise ValueError(f"{self.case_id} is not registered for {source}")
        if source == SINGLE_SHEET_SOURCE:
            return "AUTOMATED_ANDROID_MIGRATED"
        return "AUTOMATED"


def _sources(*names: str) -> frozenset[str]:
    return frozenset(names)


_DEFINITIONS = (
    # Login cases migrated to the canonical step sheet keep the single-sheet
    # source during the compatibility period.
    AndroidCaseDefinition(
        "TC-LOGIN-001",
        "login",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "login"),
    ),
    AndroidCaseDefinition(
        "TC-LOGIN-002",
        "login",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "login"),
    ),
    AndroidCaseDefinition(
        "TC-LOGIN-003",
        "login",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "login"),
    ),
    AndroidCaseDefinition(
        "TC-LOGIN-004",
        "login",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "login"),
    ),
    AndroidCaseDefinition(
        "TC-LOGIN-005",
        "login",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "login"),
    ),
    AndroidCaseDefinition(
        "TC-LOGIN-006",
        "login",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "login"),
    ),
    AndroidCaseDefinition(
        "TC-LOGIN-007",
        "login",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "login"),
    ),
    # Home and case-library cases.
    AndroidCaseDefinition(
        "TC-HOME-001",
        "home",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "home"),
    ),
    AndroidCaseDefinition(
        "TC-HOME-002",
        "home",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "home"),
        flow="home_search",
        result_state="empty",
    ),
    AndroidCaseDefinition(
        "TC-HOME-003",
        "home",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "home"),
        flow="home_search",
        result_state="empty",
    ),
    AndroidCaseDefinition(
        "TC-HOME-004",
        "home",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "home", "requires_seed"),
        requires_seed=True,
        flow="home_search",
    ),
    AndroidCaseDefinition(
        "TC-HOME-005",
        "home",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "home", "requires_seed"),
        requires_seed=True,
        flow="home_search",
    ),
    AndroidCaseDefinition(
        "TC-HOME-006",
        "home",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "home", "requires_seed"),
        requires_seed=True,
        flow="home_search",
        filter_locator="a_records_unique_tv",
    ),
    AndroidCaseDefinition(
        "TC-HOME-007",
        "home",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "home", "requires_seed"),
        requires_seed=True,
        flow="home_search",
        result_state="exact",
        filter_value="${SEEDED_CUSTOMER_PHONE}",
        filter_locator="a_records_unique_tv",
    ),
    AndroidCaseDefinition(
        "TC-HOME-008",
        "home",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "home"),
    ),
    AndroidCaseDefinition(
        "TC-HOME-009",
        "home",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "home"),
    ),
    AndroidCaseDefinition(
        "TC-HOME-010",
        "home",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "home"),
    ),
    AndroidCaseDefinition(
        "TC-HOME-011",
        "home",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "home"),
    ),
    AndroidCaseDefinition(
        "TC-HOME-012",
        "home",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "home"),
    ),
    AndroidCaseDefinition(
        "TC-CASE-001",
        "case-library",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "case-library"),
    ),
    AndroidCaseDefinition(
        "TC-CASE-002",
        "case-library",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "case-library"),
        flow="case_search",
        result_state="empty",
        filter_value="${NON_EXISTENT_CASE_TAG}",
    ),
    AndroidCaseDefinition(
        "TC-CASE-003",
        "case-library",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "case-library", "requires_seed"),
        requires_seed=True,
        flow="case_search",
        filter_value="火",
    ),
    AndroidCaseDefinition(
        "TC-CASE-004",
        "case-library",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "case-library", "requires_seed"),
        requires_seed=True,
        flow="case_filter",
        filter_value="火",
    ),
    # Customer list/detail slices currently implemented by the single-sheet
    # loader, plus the older canonical list slice retained by the normalizer.
    AndroidCaseDefinition(
        "TC-CUSTOMER-001",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer"),
        flow="customer_search",
        result_state="empty",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-002",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_search",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-003",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_search",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-004",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer"),
        flow="customer_search",
        result_state="empty",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-005",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_search",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-006",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_search",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-007",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_date_filter",
        result_state="empty",
        filter_value="2015-02-04~2015-02-13",
        filter_locator="customer_records_date_ll",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-008",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_date_filter",
        filter_value="2026-01-01~2026-06-28",
        filter_locator="customer_records_date_ll",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-009",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_filter",
        filter_value="男",
        filter_locator="customer_records_other_sex_man_tv",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-010",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_filter",
        filter_value="女",
        filter_locator="customer_records_other_sex_female_tv",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-011",
        "customer",
        _sources(),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_filter",
        filter_value="未知",
        blocked_reason="真实 Android 顾客筛选页没有“未知”性别选项",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-012",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_filter",
        filter_value="18-25岁",
        filter_locator="customer_records_other_age_1_tv",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-013",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_filter",
        filter_value="26-35岁",
        filter_locator="customer_records_other_age_2_tv",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-014",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_filter",
        filter_value="36-45岁",
        filter_locator="customer_records_other_age_3_tv",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-015",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_filter",
        result_state="empty",
        filter_value="45岁以上",
        filter_locator="customer_records_other_age_4_tv",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-016",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_filter",
        filter_value="18岁以下",
        filter_locator="customer_records_other_age_0_tv",
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-017",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_date_filter",
        result_state="empty",
        filter_value="2015-02-04~2015-02-28",
        filter_locator="customer_records_date_ll",
        filter_options=(
            ("customer_records_other_sex_man_tv", "男"),
            ("customer_records_other_age_1_tv", "18-25岁"),
        ),
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-018",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_date_filter",
        filter_value="2026-01-01~2026-07-30",
        filter_locator="customer_records_date_ll",
        filter_options=(
            ("customer_records_other_sex_man_tv", "男"),
            ("customer_records_other_age_1_tv", "18-25岁"),
        ),
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-020",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-021",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
    ),
    AndroidCaseDefinition(
        "TC-CUSTOMER-022",
        "customer",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="customer_filter_reset",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-001",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-003",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
        flow="detail_name_length_validation",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-004",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
        flow="detail_gender_validation",
        filter_value="男",
        filter_locator="customer_edit_sex_man_tv",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-005",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
        flow="detail_gender_validation",
        filter_value="女",
        filter_locator="customer_edit_sex_female_tv",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-006",
        "detail",
        _sources(),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
        flow="detail_gender_validation",
        filter_value="保密",
        blocked_reason="真实 Android 顾客编辑页只有男/女选项，没有保密性别选项",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-007",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
        flow="detail_birthday_validation",
        filter_value="2022-08-04",
        filter_locator="customer_edit_birthday_tv",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-008",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
        flow="detail_address_validation",
        filter_value="",
        filter_locator="customer_edit_address_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-009",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
        flow="detail_address_validation",
        filter_value="……&&*……*&",
        filter_locator="customer_edit_address_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-010",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
        flow="detail_address_validation",
        filter_value="${SPACE}",
        filter_locator="customer_edit_address_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-011",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-012",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "mutating",
            "customer",
            "requires_seed",
            "serial",
            "restorable",
        ),
        mutability="mutating",
        requires_seed=True,
        flow="detail_image_remark_roundtrip",
        filter_value="测试备注${RUN_TOKEN}",
        filter_locator="records_remark_remark_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-013",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
        flow="detail_profile_remark_validation",
        filter_value="",
        filter_locator="customer_edit_remark_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-014",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
        flow="detail_profile_remark_validation",
        filter_value="……&&*……*&",
        filter_locator="customer_edit_remark_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-015",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "mutating", "customer", "requires_seed", "serial"),
        mutability="mutating",
        requires_seed=True,
        flow="detail_marital_validation",
        filter_value="保密",
        filter_locator="customer_edit_marital_secret_tv",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-016",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "mutating",
            "customer",
            "requires_seed",
            "serial",
            "restorable",
        ),
        mutability="mutating",
        requires_seed=True,
        flow="detail_image_tag_roundtrip",
        filter_value="自动化标签${RUN_TOKEN}",
        filter_locator="records_remark_tag_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-017",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="detail_empty_tag_validation",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-018",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "mutating",
            "customer",
            "requires_seed",
            "serial",
            "restorable",
        ),
        mutability="mutating",
        requires_seed=True,
        flow="detail_image_tag_special_roundtrip",
        filter_value="%……&*${RUN_TOKEN}",
        filter_locator="records_remark_tag_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-019",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="detail_image_tag_whitespace_validation",
        filter_value="${SPACE}",
        filter_locator="records_remark_tag_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-020",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "destructive",
            "customer",
            "requires_seed",
            "serial",
            "restorable",
        ),
        mutability="destructive",
        requires_seed=True,
        flow="detail_image_tag_self_contained_delete",
        filter_value="自动化删除标签${RUN_TOKEN}",
        filter_locator="records_remark_tag_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-021",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "mutating",
            "customer",
            "requires_seed",
            "serial",
            "restorable",
        ),
        mutability="mutating",
        requires_seed=True,
        flow="detail_image_remark_roundtrip",
        filter_value="影像备注${RUN_TOKEN}",
        filter_locator="records_remark_remark_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-022",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "mutating",
            "customer",
            "requires_seed",
            "serial",
            "restorable",
        ),
        mutability="mutating",
        requires_seed=True,
        flow="detail_image_remark_special_roundtrip",
        filter_value="……&&*……*&${RUN_TOKEN}",
        filter_locator="records_remark_remark_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-023",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="detail_image_remark_validation",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-024",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="detail_management",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-026",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "mutating",
            "customer",
            "requires_seed",
            "serial",
            "restorable",
        ),
        mutability="mutating",
        requires_seed=True,
        flow="detail_card_name_sync",
        filter_value="自动化${RUN_TOKEN}",
        filter_locator="a_records_name_tv",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-027",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "mutating",
            "customer",
            "requires_seed",
            "serial",
            "restorable",
        ),
        mutability="mutating",
        requires_seed=True,
        flow="detail_card_phone_sync",
        filter_value="${RUN_PHONE}",
        filter_locator="a_records_unique_tv",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-029",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="detail_unselected_delete_prompt",
        filter_value="咨询单特定",
        filter_locator="customer_records_search_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-030",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "mutating",
            "customer",
            "requires_seed",
            "serial",
            "persistent",
        ),
        mutability="mutating",
        requires_seed=True,
        flow="detail_consultation_create",
        filter_value="咨询单特定",
        filter_locator="customer_records_search_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-031",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "mutating",
            "customer",
            "requires_seed",
            "serial",
            "restorable",
        ),
        mutability="mutating",
        requires_seed=True,
        flow="detail_consultation_tag_roundtrip",
        filter_value="咨询单标签${RUN_TOKEN}",
        filter_locator="records_remark_tag_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-032",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="detail_consultation_tag_whitespace_validation",
        filter_value="${SPACE}",
        filter_locator="records_remark_tag_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-033",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "mutating",
            "customer",
            "requires_seed",
            "serial",
            "restorable",
        ),
        mutability="mutating",
        requires_seed=True,
        flow="detail_consultation_tag_special_roundtrip",
        filter_value="%……&*${RUN_TOKEN}",
        filter_locator="records_remark_tag_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-034",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="detail_consultation_empty_tag_validation",
        filter_value="",
        filter_locator="records_remark_tag_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-035",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "mutating",
            "customer",
            "requires_seed",
            "serial",
            "restorable",
        ),
        mutability="mutating",
        requires_seed=True,
        flow="detail_consultation_remark_roundtrip",
        filter_value="咨询单备注${RUN_TOKEN}",
        filter_locator="records_remark_remark_et",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-036",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        ("tablet", "readonly", "customer", "requires_seed"),
        requires_seed=True,
        flow="detail_consultation_remark_validation",
    ),
    AndroidCaseDefinition(
        "TC-DETAIL-037",
        "detail",
        _sources(SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE),
        (
            "tablet",
            "destructive",
            "customer",
            "requires_seed",
            "serial",
            "restorable",
        ),
        mutability="destructive",
        requires_seed=True,
        flow="detail_consultation_remark_self_contained_delete",
        filter_value="自动化删除备注${RUN_TOKEN}",
        filter_locator="records_remark_remark_et",
    ),
    # Canonical profile/image slices. These are intentionally registered as
    # step-sheet cases until their read-only flows are migrated to the loader.
    AndroidCaseDefinition(
        "TC-IMAGE-001",
        "image",
        _sources(STEP_SHEET_SOURCE),
        ("tablet", "readonly", "image"),
    ),
    AndroidCaseDefinition(
        "TC-IMAGE-030",
        "image",
        _sources(STEP_SHEET_SOURCE),
        ("tablet", "readonly", "image", "requires_seed"),
        requires_seed=True,
        flow="image_score_display",
    ),
    AndroidCaseDefinition(
        "TC-IMAGE-032",
        "image",
        _sources(STEP_SHEET_SOURCE),
        ("tablet", "readonly", "image", "requires_seed"),
        requires_seed=True,
        flow="image_view_case",
    ),
    AndroidCaseDefinition(
        "TC-IMAGE-033",
        "image",
        _sources(STEP_SHEET_SOURCE),
        ("tablet", "readonly", "image", "requires_seed"),
        requires_seed=True,
        flow="image_view_report",
    ),
    AndroidCaseDefinition(
        "TC-IMAGE-034",
        "image",
        _sources(STEP_SHEET_SOURCE),
        ("tablet", "readonly", "image", "requires_seed"),
        requires_seed=True,
        flow="image_discard_exit",
    ),
    AndroidCaseDefinition(
        "TC-PROFILE-001",
        "profile",
        _sources(STEP_SHEET_SOURCE),
        ("tablet", "readonly", "profile"),
    ),
    AndroidCaseDefinition(
        "TC-PROFILE-002",
        "profile",
        _sources(STEP_SHEET_SOURCE),
        ("tablet", "readonly", "profile"),
    ),
    AndroidCaseDefinition(
        "TC-PROFILE-003",
        "profile",
        _sources(STEP_SHEET_SOURCE),
        ("tablet", "readonly", "profile"),
    ),
    AndroidCaseDefinition(
        "TC-PROFILE-004",
        "profile",
        _sources(STEP_SHEET_SOURCE),
        ("tablet", "readonly", "profile"),
    ),
)

if len({definition.case_id for definition in _DEFINITIONS}) != len(_DEFINITIONS):
    raise RuntimeError("Android case catalog contains duplicate case IDs")

ANDROID_CASE_CATALOG: Mapping[str, AndroidCaseDefinition] = MappingProxyType(
    {definition.case_id: definition for definition in _DEFINITIONS}
)


def get_android_case_definition(case_id: str) -> AndroidCaseDefinition | None:
    """Look up a case by ID, tolerating surrounding whitespace."""

    normalized = case_id.strip()
    if not normalized:
        return None
    return ANDROID_CASE_CATALOG.get(normalized)


def android_case_ids(
    *,
    source: str | None = None,
    categories: Iterable[str] | None = None,
    flows: Iterable[str] | None = None,
) -> frozenset[str]:
    """Return registered IDs filtered by adapter source, category, and flow."""

    if source is not None and source not in CATALOG_SOURCES:
        raise ValueError(f"Unsupported Android catalog source: {source}")
    category_filter = frozenset(categories) if categories is not None else None
    flow_filter = frozenset(flows) if flows is not None else None
    return frozenset(
        definition.case_id
        for definition in ANDROID_CASE_CATALOG.values()
        if (source is None or definition.supports(source))
        and (category_filter is None or definition.category in category_filter)
        and (flow_filter is None or definition.flow in flow_filter)
    )


ANDROID_SINGLE_SHEET_CASE_IDS = android_case_ids(source=SINGLE_SHEET_SOURCE)
ANDROID_STEP_SHEET_CASE_IDS = android_case_ids(source=STEP_SHEET_SOURCE)


def unknown_android_case_ids(case_ids: Iterable[str], *, source: str) -> frozenset[str]:
    """Return IDs that are not registered for the requested source."""

    supported = android_case_ids(source=source)
    return frozenset(case_id.strip() for case_id in case_ids if case_id.strip()) - supported
