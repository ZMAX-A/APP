from __future__ import annotations

from collections.abc import Callable
from typing import Self

from appium.webdriver.common.appiumby import AppiumBy
from appium.webdriver.webdriver import WebDriver
from appium.webdriver.webelement import WebElement
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait

PACKAGE = "com.xiaofutech.yanjia_ai"
Locator = tuple[str, str]


def resource_id(name: str) -> Locator:
    return AppiumBy.ID, f"{PACKAGE}:id/{name}"


class BaseScreen:
    root: Locator

    def __init__(self, driver: WebDriver, timeout: float = 10) -> None:
        self.driver = driver
        self.timeout = timeout

    def find(self, locator: Locator, timeout: float | None = None) -> WebElement:
        return WebDriverWait(self.driver, timeout or self.timeout).until(
            lambda current: current.find_element(*locator)
        )

    def find_all(self, locator: Locator) -> list[WebElement]:
        return self.driver.find_elements(*locator)

    def click(self, locator: Locator, timeout: float | None = None) -> None:
        element = WebDriverWait(self.driver, timeout or self.timeout).until(
            lambda current: current.find_element(*locator)
        )
        WebDriverWait(self.driver, timeout or self.timeout).until(
            lambda _: element.is_displayed() and element.is_enabled()
        )
        element.click()

    def is_visible(self, locator: Locator, timeout: float = 1.0) -> bool:
        try:
            element = WebDriverWait(self.driver, timeout).until(
                lambda current: current.find_element(*locator)
            )
            return element.is_displayed()
        except TimeoutException:
            return False

    def wait_until(self, condition: Callable[[WebDriver], object], timeout: float = 10) -> object:
        return WebDriverWait(self.driver, timeout).until(condition)

    def wait_loaded(self, timeout: float | None = None) -> Self:
        self.find(self.root, timeout)
        return self
