from __future__ import annotations

from appium.webdriver.webdriver import WebDriver

from yanjia_automation.config import Settings
from yanjia_automation.screens.home import HomeScreen
from yanjia_automation.screens.login import LoginScreen


def ensure_home(driver: WebDriver, settings: Settings) -> HomeScreen:
    home = HomeScreen(driver)
    login = LoginScreen(driver)

    driver.activate_app(settings.app_package)
    if home.is_visible(home.root, timeout=4):
        return home

    if login.is_visible(login.root, timeout=2):
        login.sign_in(settings.credentials(), settings.store_name)
        return home.wait_loaded(timeout=30)

    for _ in range(6):
        driver.back()
        if home.is_visible(home.root, timeout=2):
            return home
        if login.is_visible(login.root, timeout=1):
            login.sign_in(settings.credentials(), settings.store_name)
            return home.wait_loaded(timeout=30)

    raise AssertionError(
        f"Could not navigate to the home screen from activity {driver.current_activity!r}"
    )
