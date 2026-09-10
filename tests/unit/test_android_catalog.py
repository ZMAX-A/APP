from __future__ import annotations

from yanjia_automation.excel.android_catalog import (
    ANDROID_SINGLE_SHEET_CASE_IDS,
    ANDROID_STEP_SHEET_CASE_IDS,
    SINGLE_SHEET_SOURCE,
    STEP_SHEET_SOURCE,
    get_android_case_definition,
)
from yanjia_automation.excel.single_sheet import (
    ANDROID_CUSTOMER_CASES,
    ANDROID_DETAIL_CASES,
    ANDROID_HOME_CASES,
    ANDROID_LOGIN_CASES,
    ANDROID_SINGLE_SHEET_CASES,
)


def test_single_sheet_adapter_uses_catalog_case_groups() -> None:
    assert ANDROID_SINGLE_SHEET_CASES == ANDROID_SINGLE_SHEET_CASE_IDS
    assert ANDROID_LOGIN_CASES == {
        "TC-LOGIN-001",
        "TC-LOGIN-002",
        "TC-LOGIN-003",
        "TC-LOGIN-004",
        "TC-LOGIN-005",
        "TC-LOGIN-006",
        "TC-LOGIN-007",
    }
    assert ANDROID_HOME_CASES == {
        "TC-HOME-001",
        "TC-HOME-002",
        "TC-HOME-003",
        "TC-HOME-004",
        "TC-HOME-005",
        "TC-HOME-006",
        "TC-HOME-007",
        "TC-HOME-008",
        "TC-HOME-009",
        "TC-HOME-010",
        "TC-HOME-011",
        "TC-HOME-012",
        "TC-CASE-001",
        "TC-CASE-002",
        "TC-CASE-003",
        "TC-CASE-004",
    }
    assert ANDROID_CUSTOMER_CASES == {
        "TC-CUSTOMER-001",
        "TC-CUSTOMER-002",
        "TC-CUSTOMER-003",
        "TC-CUSTOMER-004",
        "TC-CUSTOMER-005",
        "TC-CUSTOMER-006",
        "TC-CUSTOMER-007",
        "TC-CUSTOMER-008",
        "TC-CUSTOMER-009",
        "TC-CUSTOMER-010",
        "TC-CUSTOMER-012",
        "TC-CUSTOMER-013",
        "TC-CUSTOMER-014",
        "TC-CUSTOMER-015",
        "TC-CUSTOMER-016",
        "TC-CUSTOMER-017",
        "TC-CUSTOMER-018",
        "TC-CUSTOMER-020",
        "TC-CUSTOMER-021",
        "TC-CUSTOMER-022",
    }
    assert ANDROID_DETAIL_CASES == {
        "TC-DETAIL-001",
        "TC-DETAIL-003",
        "TC-DETAIL-004",
        "TC-DETAIL-005",
        "TC-DETAIL-007",
        "TC-DETAIL-008",
        "TC-DETAIL-009",
        "TC-DETAIL-010",
        "TC-DETAIL-011",
        "TC-DETAIL-012",
        "TC-DETAIL-013",
        "TC-DETAIL-014",
        "TC-DETAIL-015",
        "TC-DETAIL-016",
        "TC-DETAIL-017",
        "TC-DETAIL-018",
        "TC-DETAIL-019",
        "TC-DETAIL-020",
        "TC-DETAIL-021",
        "TC-DETAIL-022",
        "TC-DETAIL-023",
        "TC-DETAIL-024",
        "TC-DETAIL-029",
        "TC-DETAIL-030",
        "TC-DETAIL-026",
        "TC-DETAIL-027",
        "TC-DETAIL-031",
        "TC-DETAIL-032",
        "TC-DETAIL-033",
        "TC-DETAIL-034",
        "TC-DETAIL-035",
        "TC-DETAIL-036",
        "TC-DETAIL-037",
    }


def test_catalog_registers_both_execution_sources() -> None:
    assert len(ANDROID_SINGLE_SHEET_CASE_IDS) == 76
    assert len(ANDROID_STEP_SHEET_CASE_IDS) == 85
    shared = ANDROID_SINGLE_SHEET_CASE_IDS & ANDROID_STEP_SHEET_CASE_IDS
    assert shared == {
        "TC-LOGIN-007",
        "TC-LOGIN-001",
        "TC-LOGIN-002",
        "TC-LOGIN-003",
        "TC-LOGIN-004",
        "TC-LOGIN-005",
        "TC-LOGIN-006",
        "TC-HOME-001",
        "TC-HOME-002",
        "TC-HOME-003",
        "TC-HOME-004",
        "TC-HOME-005",
        "TC-HOME-006",
        "TC-HOME-007",
        "TC-HOME-008",
        "TC-HOME-009",
        "TC-HOME-010",
        "TC-HOME-011",
        "TC-HOME-012",
        "TC-CASE-001",
        "TC-CASE-002",
        "TC-CASE-003",
        "TC-CASE-004",
        "TC-CUSTOMER-001",
        "TC-CUSTOMER-002",
        "TC-CUSTOMER-003",
        "TC-CUSTOMER-004",
        "TC-CUSTOMER-005",
        "TC-CUSTOMER-006",
        "TC-CUSTOMER-007",
        "TC-CUSTOMER-008",
        "TC-CUSTOMER-009",
        "TC-CUSTOMER-010",
        "TC-CUSTOMER-012",
        "TC-CUSTOMER-013",
        "TC-CUSTOMER-014",
        "TC-CUSTOMER-015",
        "TC-CUSTOMER-016",
        "TC-CUSTOMER-017",
        "TC-CUSTOMER-018",
        "TC-CUSTOMER-020",
        "TC-CUSTOMER-021",
        "TC-CUSTOMER-022",
        "TC-DETAIL-001",
        "TC-DETAIL-003",
        "TC-DETAIL-004",
        "TC-DETAIL-005",
        "TC-DETAIL-007",
        "TC-DETAIL-008",
        "TC-DETAIL-009",
        "TC-DETAIL-010",
        "TC-DETAIL-011",
        "TC-DETAIL-012",
        "TC-DETAIL-013",
        "TC-DETAIL-014",
        "TC-DETAIL-015",
        "TC-DETAIL-016",
        "TC-DETAIL-017",
        "TC-DETAIL-018",
        "TC-DETAIL-019",
        "TC-DETAIL-020",
        "TC-DETAIL-021",
        "TC-DETAIL-022",
        "TC-DETAIL-023",
        "TC-DETAIL-024",
        "TC-DETAIL-029",
        "TC-DETAIL-030",
        "TC-DETAIL-026",
        "TC-DETAIL-027",
        "TC-DETAIL-031",
        "TC-DETAIL-032",
        "TC-DETAIL-033",
        "TC-DETAIL-034",
        "TC-DETAIL-035",
        "TC-DETAIL-036",
        "TC-DETAIL-037",
    }
    detail = get_android_case_definition(" TC-DETAIL-011 ")
    profile = get_android_case_definition("TC-PROFILE-001")
    assert detail is not None and detail.supports(SINGLE_SHEET_SOURCE)
    assert profile is not None and profile.supports(STEP_SHEET_SOURCE)


def test_catalog_exposes_safety_metadata() -> None:
    detail = get_android_case_definition("TC-DETAIL-001")
    assert detail is not None
    assert detail.mutability == "mutating"
    assert detail.requires_seed is True
    assert detail.tags == ("tablet", "mutating", "customer", "requires_seed", "serial")
    assert detail.automation_status_for(SINGLE_SHEET_SOURCE) == "AUTOMATED_ANDROID_MIGRATED"

    assert detail.automation_status_for(STEP_SHEET_SOURCE) == "AUTOMATED"

    name_length = get_android_case_definition("TC-DETAIL-003")
    assert name_length is not None
    assert name_length.mutability == "mutating"
    assert name_length.requires_seed is True
    assert name_length.flow == "detail_name_length_validation"
    assert name_length.supports(SINGLE_SHEET_SOURCE)
    assert name_length.supports(STEP_SHEET_SOURCE)

    for case_id, gender, locator in (
        ("TC-DETAIL-004", "男", "customer_edit_sex_man_tv"),
        ("TC-DETAIL-005", "女", "customer_edit_sex_female_tv"),
    ):
        gender_case = get_android_case_definition(case_id)
        assert gender_case is not None
        assert gender_case.flow == "detail_gender_validation"
        assert gender_case.filter_value == gender
        assert gender_case.filter_locator == locator
        assert gender_case.supports(SINGLE_SHEET_SOURCE)
        assert gender_case.supports(STEP_SHEET_SOURCE)

    blocked_gender = get_android_case_definition("TC-DETAIL-006")
    assert blocked_gender is not None
    assert not blocked_gender.sources
    assert blocked_gender.blocked_reason == (
        "真实 Android 顾客编辑页只有男/女选项，没有保密性别选项"
    )

    birthday = get_android_case_definition("TC-DETAIL-007")
    assert birthday is not None
    assert birthday.flow == "detail_birthday_validation"
    assert birthday.filter_value == "2022-08-04"
    assert birthday.filter_locator == "customer_edit_birthday_tv"
    assert birthday.mutability == "mutating"

    for case_id in ("TC-DETAIL-008", "TC-DETAIL-009", "TC-DETAIL-010"):
        address = get_android_case_definition(case_id)
        assert address is not None
        assert address.flow == "detail_address_validation"
        assert address.filter_locator == "customer_edit_address_et"
        assert address.mutability == "mutating"

    for case_id, value in (
        ("TC-DETAIL-013", ""),
        ("TC-DETAIL-014", "……&&*……*&"),
    ):
        remark = get_android_case_definition(case_id)
        assert remark is not None
        assert remark.flow == "detail_profile_remark_validation"
        assert remark.filter_value == value
        assert remark.filter_locator == "customer_edit_remark_et"
        assert remark.mutability == "mutating"

    marital = get_android_case_definition("TC-DETAIL-015")
    assert marital is not None
    assert marital.flow == "detail_marital_validation"
    assert marital.filter_value == "保密"
    assert marital.filter_locator == "customer_edit_marital_secret_tv"
    assert marital.mutability == "mutating"

    for case_id, flow, value in (
        ("TC-DETAIL-016", "detail_image_tag_roundtrip", "自动化标签${RUN_TOKEN}"),
        (
            "TC-DETAIL-018",
            "detail_image_tag_special_roundtrip",
            "%……&*${RUN_TOKEN}",
        ),
        (
            "TC-DETAIL-031",
            "detail_consultation_tag_roundtrip",
            "咨询单标签${RUN_TOKEN}",
        ),
        (
            "TC-DETAIL-033",
            "detail_consultation_tag_special_roundtrip",
            "%……&*${RUN_TOKEN}",
        ),
    ):
        tag_roundtrip = get_android_case_definition(case_id)
        assert tag_roundtrip is not None
        assert tag_roundtrip.flow == flow
        assert tag_roundtrip.filter_value == value
        assert tag_roundtrip.filter_locator == "records_remark_tag_et"
        assert tag_roundtrip.mutability == "mutating"
        assert tag_roundtrip.requires_seed is True
        assert "restorable" in tag_roundtrip.tags

    for case_id, flow, value in (
        (
            "TC-DETAIL-012",
            "detail_image_remark_roundtrip",
            "测试备注${RUN_TOKEN}",
        ),
        (
            "TC-DETAIL-021",
            "detail_image_remark_roundtrip",
            "影像备注${RUN_TOKEN}",
        ),
        (
            "TC-DETAIL-022",
            "detail_image_remark_special_roundtrip",
            "……&&*……*&${RUN_TOKEN}",
        ),
        (
            "TC-DETAIL-035",
            "detail_consultation_remark_roundtrip",
            "咨询单备注${RUN_TOKEN}",
        ),
    ):
        remark_roundtrip = get_android_case_definition(case_id)
        assert remark_roundtrip is not None
        assert remark_roundtrip.flow == flow
        assert remark_roundtrip.filter_value == value
        assert remark_roundtrip.filter_locator == "records_remark_remark_et"
        assert remark_roundtrip.mutability == "mutating"
        assert remark_roundtrip.requires_seed is True
        assert "restorable" in remark_roundtrip.tags

    for case_id, flow, value, locator in (
        (
            "TC-DETAIL-026",
            "detail_card_name_sync",
            "自动化${RUN_TOKEN}",
            "a_records_name_tv",
        ),
        (
            "TC-DETAIL-027",
            "detail_card_phone_sync",
            "${RUN_PHONE}",
            "a_records_unique_tv",
        ),
    ):
        card_sync = get_android_case_definition(case_id)
        assert card_sync is not None
        assert card_sync.flow == flow
        assert card_sync.filter_value == value
        assert card_sync.filter_locator == locator
        assert card_sync.mutability == "mutating"
        assert card_sync.requires_seed is True
        assert "restorable" in card_sync.tags

    empty_tag = get_android_case_definition("TC-DETAIL-017")
    assert empty_tag is not None
    assert empty_tag.flow == "detail_empty_tag_validation"
    assert empty_tag.mutability == "readonly"
    assert empty_tag.requires_seed is True

    for case_id, flow, value in (
        ("TC-DETAIL-019", "detail_image_tag_whitespace_validation", "${SPACE}"),
        (
            "TC-DETAIL-032",
            "detail_consultation_tag_whitespace_validation",
            "${SPACE}",
        ),
        ("TC-DETAIL-034", "detail_consultation_empty_tag_validation", ""),
    ):
        tag_validation = get_android_case_definition(case_id)
        assert tag_validation is not None
        assert tag_validation.flow == flow
        assert tag_validation.filter_value == value
        assert tag_validation.filter_locator == "records_remark_tag_et"
        assert tag_validation.mutability == "readonly"
        assert tag_validation.requires_seed is True

    readonly_detail = get_android_case_definition("TC-DETAIL-024")
    assert readonly_detail is not None
    assert readonly_detail.mutability == "readonly"
    assert readonly_detail.requires_seed is True
    assert readonly_detail.flow == "detail_management"

    image_remark = get_android_case_definition("TC-DETAIL-023")
    consultation_remark = get_android_case_definition("TC-DETAIL-036")
    assert image_remark is not None
    assert image_remark.mutability == "readonly"
    assert image_remark.requires_seed is True
    assert image_remark.flow == "detail_image_remark_validation"
    assert consultation_remark is not None
    assert consultation_remark.mutability == "readonly"
    assert consultation_remark.requires_seed is True
    assert consultation_remark.flow == "detail_consultation_remark_validation"

    unselected_delete = get_android_case_definition("TC-DETAIL-029")
    assert unselected_delete is not None
    assert unselected_delete.sources == {SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE}
    assert unselected_delete.flow == "detail_unselected_delete_prompt"
    assert unselected_delete.filter_value == "咨询单特定"
    assert unselected_delete.blocked_reason is None

    consultation_create = get_android_case_definition("TC-DETAIL-030")
    assert consultation_create is not None
    assert consultation_create.sources == {SINGLE_SHEET_SOURCE, STEP_SHEET_SOURCE}
    assert consultation_create.flow == "detail_consultation_create"
    assert consultation_create.filter_value == "咨询单特定"
    assert consultation_create.mutability == "mutating"
    assert consultation_create.requires_seed is True
    assert "persistent" in consultation_create.tags
    assert "restorable" not in consultation_create.tags


def test_catalog_marks_search_result_contracts() -> None:
    empty_home = get_android_case_definition("TC-HOME-002")
    seeded_home = get_android_case_definition("TC-HOME-004")
    exact_home = get_android_case_definition("TC-HOME-007")
    empty_customer = get_android_case_definition("TC-CUSTOMER-001")
    seeded_customer = get_android_case_definition("TC-CUSTOMER-005")

    assert empty_home is not None
    assert empty_home.flow == "home_search"
    assert empty_home.result_state == "empty"
    assert seeded_home is not None and seeded_home.requires_seed is True
    assert exact_home is not None
    assert exact_home.result_state == "exact"
    assert exact_home.filter_value == "${SEEDED_CUSTOMER_PHONE}"
    assert exact_home.filter_locator == "a_records_unique_tv"
    assert exact_home.requires_seed is True
    assert empty_customer is not None and empty_customer.result_state == "empty"
    assert seeded_customer is not None and seeded_customer.supports(SINGLE_SHEET_SOURCE)

    empty_case_search = get_android_case_definition("TC-CASE-002")
    case_search = get_android_case_definition("TC-CASE-003")
    case_filter = get_android_case_definition("TC-CASE-004")
    assert empty_case_search is not None
    assert empty_case_search.flow == "case_search"
    assert empty_case_search.result_state == "empty"
    assert empty_case_search.filter_value == "${NON_EXISTENT_CASE_TAG}"
    assert empty_case_search.requires_seed is False
    assert case_search is not None
    assert case_search.flow == "case_search"
    assert case_search.filter_value == "火"
    assert case_search.requires_seed is True
    assert case_filter is not None
    assert case_filter.flow == "case_filter"
    assert case_filter.filter_value == "火"


def test_catalog_marks_readonly_image_navigation_slice() -> None:
    expected_flows = {
        "TC-IMAGE-030": "image_score_display",
        "TC-IMAGE-032": "image_view_case",
        "TC-IMAGE-033": "image_view_report",
        "TC-IMAGE-034": "image_discard_exit",
    }
    for case_id, flow in expected_flows.items():
        definition = get_android_case_definition(case_id)
        assert definition is not None
        assert definition.supports(STEP_SHEET_SOURCE)
        assert not definition.supports(SINGLE_SHEET_SOURCE)
        assert definition.mutability == "readonly"
        assert definition.requires_seed is True
        assert definition.flow == flow
        assert definition.tags == ("tablet", "readonly", "image", "requires_seed")


def test_catalog_marks_customer_filter_options() -> None:
    expected = {
        "TC-CUSTOMER-009": "男",
        "TC-CUSTOMER-010": "女",
        "TC-CUSTOMER-012": "18-25岁",
        "TC-CUSTOMER-013": "26-35岁",
        "TC-CUSTOMER-014": "36-45岁",
        "TC-CUSTOMER-015": "45岁以上",
        "TC-CUSTOMER-016": "18岁以下",
    }
    for case_id, filter_value in expected.items():
        definition = get_android_case_definition(case_id)
        assert definition is not None
        assert definition.flow == "customer_filter"
        assert definition.filter_value == filter_value
        assert definition.requires_seed is True

    empty_age = get_android_case_definition("TC-CUSTOMER-015")
    assert empty_age is not None and empty_age.result_state == "empty"

    blocked = get_android_case_definition("TC-CUSTOMER-011")
    assert blocked is not None
    assert not blocked.supports(SINGLE_SHEET_SOURCE)
    assert blocked.blocked_reason == "真实 Android 顾客筛选页没有“未知”性别选项"

    reset = get_android_case_definition("TC-CUSTOMER-022")
    assert reset is not None
    assert reset.flow == "customer_filter_reset"
    assert reset.mutability == "readonly"
    assert reset.requires_seed is True
    assert reset.supports(SINGLE_SHEET_SOURCE)
    assert reset.supports(STEP_SHEET_SOURCE)


def test_catalog_marks_customer_date_filter_contracts() -> None:
    early_empty = get_android_case_definition("TC-CUSTOMER-007")
    positive = get_android_case_definition("TC-CUSTOMER-008")
    empty = get_android_case_definition("TC-CUSTOMER-017")
    latest = get_android_case_definition("TC-CUSTOMER-018")

    assert early_empty is not None and early_empty.flow == "customer_date_filter"
    assert early_empty.result_state == "empty"
    assert early_empty.filter_value == "2015-02-04~2015-02-13"
    assert early_empty.filter_locator == "customer_records_date_ll"
    assert early_empty.filter_options == ()
    assert early_empty.requires_seed is True
    assert positive is not None and positive.flow == "customer_date_filter"
    assert positive.filter_value == "2026-01-01~2026-06-28"
    assert positive.filter_locator == "customer_records_date_ll"
    assert empty is not None and empty.result_state == "empty"
    assert empty.filter_value == "2015-02-04~2015-02-28"
    assert latest is not None and latest.filter_value == "2026-01-01~2026-07-30"
    assert latest.filter_options == (
        ("customer_records_other_sex_man_tv", "男"),
        ("customer_records_other_age_1_tv", "18-25岁"),
    )
