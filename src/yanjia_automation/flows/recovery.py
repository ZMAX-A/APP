from __future__ import annotations

from appium.webdriver.common.appiumby import AppiumBy
from appium.webdriver.webdriver import WebDriver
from appium.webdriver.webelement import WebElement
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support.ui import WebDriverWait

from yanjia_automation.config import Settings
from yanjia_automation.screens.home import HomeScreen
from yanjia_automation.screens.login import LoginScreen
from yanjia_automation.screens.settings import SettingsScreen

_ANDROID_ALERT_TITLE = (AppiumBy.ID, "android:id/alertTitle")
_OEM_ALERT_TITLE = (AppiumBy.XPATH, "//*[contains(@resource-id, ':id/alertTitle')]")
_ANDROID_ALERT_CANCEL = (AppiumBy.ID, "android:id/button2")
_CREDENTIAL_SAVE_TITLE_TOKENS = ("账号", "密码")
_CREDENTIAL_SAVE_ACTION_TOKENS = ("保存", "存储")


def _element_text(element: WebElement) -> str:
    try:
        return str(element.get_attribute("text") or "")
    except WebDriverException:
        try:
            return str(element.text or "")
        except WebDriverException:
            return ""


def _dismiss_android_credential_save_prompt(driver: WebDriver) -> bool:
    """Cancel only the Android system prompt that offers to save credentials."""

    try:
        titles = driver.find_elements(*_ANDROID_ALERT_TITLE)
        if not titles:
            titles = driver.find_elements(*_OEM_ALERT_TITLE)
        if len(titles) != 1:
            return False

        normalized_title = "".join(_element_text(titles[0]).split())
        if not all(token in normalized_title for token in _CREDENTIAL_SAVE_TITLE_TOKENS):
            return False
        if not any(token in normalized_title for token in _CREDENTIAL_SAVE_ACTION_TOKENS):
            return False

        cancel_buttons = driver.find_elements(*_ANDROID_ALERT_CANCEL)
        if len(cancel_buttons) != 1:
            return False
        cancel = cancel_buttons[0]
        if not cancel.is_displayed() or not cancel.is_enabled():
            return False
        cancel.click()
        return True
    except WebDriverException:
        return False


def _wait_for_home(driver: WebDriver, home: HomeScreen, *, timeout: float) -> HomeScreen:
    def home_ready(_: WebDriver) -> bool:
        if home.is_visible(home.root, timeout=0.2):
            return True
        _dismiss_android_credential_save_prompt(driver)
        return home.is_visible(home.root, timeout=0.2)

    WebDriverWait(driver, timeout).until(home_ready)
    return home.wait_loaded(timeout=timeout)


def recover_login_if_needed(
    driver: WebDriver,
    settings: Settings,
    *,
    timeout: float = 30,
) -> bool:
    """Restore an authenticated home state after an unexpected logout."""

    login = LoginScreen(driver)
    try:
        on_login = driver.current_activity.endswith(".LoginActivity")
        if not on_login:
            on_login = login.is_visible(login.root, timeout=0.2)
    except WebDriverException:
        return False
    if not on_login:
        return False

    login.sign_in(settings.credentials(), settings.store_name)
    _wait_for_home(driver, HomeScreen(driver), timeout=timeout)
    return True


def restart_to_home(
    driver: WebDriver,
    settings: Settings,
    *,
    timeout: float = 30,
) -> HomeScreen:
    """Return to a deterministic home state regardless of the previous test.

    The image download overlay consumes Android back events. Terminating and
    activating the package relaunches the exported SplashActivity while keeping
    the authenticated session because the driver uses noReset=true. Android may
    then cover MainActivity with a credential-save prompt, which is cancelled
    only when its title contains the expected account/password/save semantics.
    """

    home = HomeScreen(driver)
    login = LoginScreen(driver)

    driver.terminate_app(settings.app_package)
    driver.activate_app(settings.app_package)

    WebDriverWait(driver, timeout).until(
        lambda _: home.is_visible(home.root, timeout=0.2)
        or login.is_visible(login.root, timeout=0.2)
        or driver.current_activity.endswith(".MainActivity")
    )
    if home.is_visible(home.root, timeout=0.5) or driver.current_activity.endswith(
        ".MainActivity"
    ):
        return _wait_for_home(driver, home, timeout=timeout)

    login.sign_in(settings.credentials(), settings.store_name)
    return _wait_for_home(driver, home, timeout=timeout)


def restart_to_login(
    driver: WebDriver,
    settings: Settings,
    *,
    timeout: float = 30,
) -> LoginScreen:
    """Return to LoginActivity without requiring a manual logout."""

    home = HomeScreen(driver)
    login = LoginScreen(driver)
    preferences = SettingsScreen(driver)

    driver.terminate_app(settings.app_package)
    driver.activate_app(settings.app_package)

    WebDriverWait(driver, timeout).until(
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
        preferences.wait_loaded(timeout=timeout)
    elif not preferences.is_visible(preferences.root, timeout=0.5):
        raise AssertionError("无法从当前页面进入设置页以退出登录")

    preferences.logout_and_confirm()
    WebDriverWait(driver, timeout).until(
        lambda _: login.is_visible(login.root, timeout=0.2)
    )
    return login
