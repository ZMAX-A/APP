from __future__ import annotations

import pytest
from appium.webdriver.webdriver import WebDriver

from yanjia_automation.config import Settings
from yanjia_automation.flows.navigation import ensure_home
from yanjia_automation.screens.case_library import CaseLibraryScreen
from yanjia_automation.screens.customer import CustomerDetailScreen, CustomerListScreen
from yanjia_automation.screens.home import HomeScreen
from yanjia_automation.screens.settings import ProfileScreen, SettingsScreen

pytestmark = [pytest.mark.tablet, pytest.mark.readonly, pytest.mark.requires_auth]


@pytest.mark.p0
@pytest.mark.smoke
@pytest.mark.case_id("TC-LOGIN-007", "TC-HOME-001")
def test_login_and_home_contract(home: HomeScreen) -> None:
    assert home.driver.current_activity.endswith(".MainActivity")
    assert home.find(home.customer_records).is_displayed()
    assert home.find(home.academy).is_displayed()
    assert home.find(home.case_library).is_displayed()
    assert home.find(home.settings).is_displayed()


@pytest.mark.p0
@pytest.mark.smoke
@pytest.mark.requires_seed
@pytest.mark.case_id("TC-HOME-008", "TC-CUSTOMER-005")
def test_customer_list_contract(home: HomeScreen) -> None:
    home.open_customer_records()
    customer_list = CustomerListScreen(home.driver).wait_loaded(timeout=15)

    assert customer_list.find(customer_list.search_input).is_displayed()
    assert customer_list.find(customer_list.date_filter).is_displayed()
    assert customer_list.find(customer_list.other_filter).is_displayed()
    assert customer_list.find_all(customer_list.cards)


@pytest.mark.p0
@pytest.mark.smoke
@pytest.mark.requires_seed
@pytest.mark.case_id("TC-CUSTOMER-006")
def test_first_customer_detail_contract(home: HomeScreen) -> None:
    home.open_customer_records()
    customer_list = CustomerListScreen(home.driver).wait_loaded(timeout=15)
    customer_list.open_first_customer()

    detail = CustomerDetailScreen(home.driver).wait_loaded(timeout=15)
    assert home.driver.current_activity.endswith(".CustomerDetailActivity")
    assert detail.find(detail.customer_info).is_displayed()
    assert detail.find(detail.edit_button).is_displayed()


@pytest.mark.p0
@pytest.mark.smoke
@pytest.mark.case_id("TC-HOME-010", "TC-CASE-001")
def test_case_library_contract(home: HomeScreen) -> None:
    home.open_case_library()
    cases = CaseLibraryScreen(home.driver).wait_loaded(timeout=15)

    assert home.driver.current_activity.endswith(".CaseActivity")
    assert cases.find(cases.search_input).is_displayed()
    assert cases.find_all(cases.tags)
    assert cases.find_all(cases.categories)


@pytest.mark.p0
@pytest.mark.smoke
@pytest.mark.case_id(
    "TC-HOME-011",
    "TC-PROFILE-001",
    "TC-PROFILE-002",
    "TC-PROFILE-003",
    "TC-PROFILE-004",
)
def test_profile_contract(
    driver: WebDriver,
    settings: Settings,
    home: HomeScreen,
) -> None:
    home.open_settings()
    settings_screen = SettingsScreen(driver).wait_loaded(timeout=15)
    settings_screen.open_personal_center()
    profile = ProfileScreen(driver).wait_loaded(timeout=15)

    assert driver.current_activity.endswith(".PersonalActivity")
    assert profile.find(profile.account).text.strip()
    assert profile.find(profile.name).text.strip()
    assert profile.find(profile.registered_at).text.strip()

    # Phone is optional for this account; presence of the field itself is the contract.
    assert profile.find(profile.phone).is_displayed()

    ensure_home(driver, settings)
