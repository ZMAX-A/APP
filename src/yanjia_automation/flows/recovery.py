from __future__ import annotations

from appium.webdriver.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait

from yanjia_automation.config import Settings
from yanjia_automation.screens.home import HomeScreen
from yanjia_automation.screens.login import LoginScreen


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
