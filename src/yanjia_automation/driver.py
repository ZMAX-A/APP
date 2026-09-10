from __future__ import annotations

from dataclasses import replace
from threading import Lock

from appium import webdriver
from appium.options.android import UiAutomator2Options
from appium.webdriver.webdriver import WebDriver
from selenium.common.exceptions import (
    InvalidSessionIdException,
    NoSuchDriverException,
    WebDriverException,
)

from yanjia_automation.config import Settings


def create_driver(settings: Settings) -> WebDriver:
    options = UiAutomator2Options()
    options.platform_name = "Android"
    options.automation_name = "UiAutomator2"
    options.device_name = settings.udid or "Android tablet"
    options.app_package = settings.app_package
    options.app_activity = settings.app_activity
    options.no_reset = settings.no_reset
    options.auto_grant_permissions = True
    options.new_command_timeout = 180
    options.set_capability("appium:adbExecTimeout", 60000)
    if settings.skip_device_initialization:
        options.set_capability("appium:skipDeviceInitialization", True)
    if settings.skip_server_installation:
        options.set_capability("appium:skipServerInstallation", True)
    options.set_capability("appium:appWaitActivity", "*")
    options.set_capability("appium:forceAppLaunch", True)
    options.set_capability("appium:disableWindowAnimation", True)
    if settings.udid:
        options.udid = settings.udid

    driver = webdriver.Remote(settings.appium_server_url, options=options)
    driver.implicitly_wait(0)
    return driver


class DriverManager:
    """Own the Appium session and allow a failed session to be recreated safely."""

    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._driver: WebDriver | None = None
        self._lock = Lock()

    def get(self) -> WebDriver:
        with self._lock:
            if self._driver is None:
                self._driver = create_driver(self._settings)
            return self._driver

    def restart(self, *, repair_uiautomator2: bool = True) -> WebDriver:
        """Recreate the session and optionally redeploy a broken UiAutomator2 server."""

        with self._lock:
            self._quit_current()
            restart_settings = self._settings
            if repair_uiautomator2 and self._settings.skip_server_installation:
                restart_settings = replace(
                    self._settings,
                    skip_server_installation=False,
                )
            self._driver = create_driver(restart_settings)
            return self._driver

    def quit(self) -> None:
        with self._lock:
            self._quit_current()

    def _quit_current(self) -> None:
        active_driver, self._driver = self._driver, None
        if active_driver is None:
            return
        try:
            active_driver.quit()
        except Exception:
            # A broken Appium session often cannot acknowledge quit.
            pass


_RECOVERABLE_MESSAGES = (
    "connection refused",
    "connection reset",
    "could not proxy command",
    "invalid session id",
    "instrumentation process is not running",
    "max retries exceeded",
    "no such driver",
    "session deleted",
    "socket hang up",
    "uiautomator2 server",
)


def is_recoverable_driver_error(error: BaseException) -> bool:
    """Identify transport/session failures without retrying business assertions."""
    if isinstance(error, AssertionError):
        return False

    current: BaseException | None = error
    visited: set[int] = set()
    while current is not None and id(current) not in visited:
        visited.add(id(current))
        if isinstance(current, (InvalidSessionIdException, NoSuchDriverException)):
            return True
        message = str(current).lower()
        if isinstance(current, (WebDriverException, ConnectionError, OSError)) and any(
            token in message for token in _RECOVERABLE_MESSAGES
        ):
            return True
        current = current.__cause__ or current.__context__
    return False
