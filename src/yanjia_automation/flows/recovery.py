from __future__ import annotations

from appium.webdriver.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait

from yanjia_automation.config import Settings
from yanjia_automation.screens.home import HomeScreen
from yanjia_automation.screens.login import LoginScreen
from yanjia_automation.screens.settings import SettingsScreen


def restart_to_home(driver: WebDriver, settings: Settings) -> HomeScreen:
    """Return to a deterministic home state regardless of the previous test.

    The image download overlay consumes Android back events. Terminating and
    activating the package relaunches the exported SplashActivity while keeping
    the authenticated session because the driver uses noReset=true.
    """

    home = HomeScreen(driver)
    login = LoginScreen(driver)

    driver.terminate_app(settings.app_package)
    driver.activate_app(settings.app_package)

    WebDriverWait(driver, 20).until(
        lambda _: home.is_visible(home.root, timeout=0.2)
        or login.is_visible(login.root, timeout=0.2)
    )
    if home.is_visible(home.root, timeout=0.5):
        return home

    login.sign_in(settings.credentials(), settings.store_name)
    home.wait_loaded(timeout=30)
    return home


def restart_to_login(driver: WebDriver, settings: Settings) -> LoginScreen:
    """Return to LoginActivity without requiring a manual logout."""

    home = HomeScreen(driver)
    login = LoginScreen(driver)
    preferences = SettingsScreen(driver)

    driver.terminate_app(settings.app_package)
    driver.activate_app(settings.app_package)

    WebDriverWait(driver, 20).until(
        lambda _: login.is_visible(login.root, timeout=0.2)
        or home.is_visible(home.root, timeout=0.2)
        or preferences.is_visible(preferences.root, timeout=0.2)
    )
    if login.is_visible(login.root, timeout=0.5):
        if login.is_visible(login.error_confirm, timeout=0.5):
            login.click(login.error_confirm)
        return login

    if home.is_visible(home.root, timeout=0.5):
        home.open_settings()
        preferences.wait_loaded(timeout=20)
    elif not preferences.is_visible(preferences.root, timeout=0.5):
        raise AssertionError("无法从当前页面进入设置页以退出登录")

    preferences.logout_and_confirm()
    WebDriverWait(driver, 20).until(
        lambda _: login.is_visible(login.root, timeout=0.2)
    )
    return login
