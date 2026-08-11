from __future__ import annotations

import pytest
from appium.webdriver.webdriver import WebDriver

from yanjia_automation.config import Settings
from yanjia_automation.flows.recovery import restart_to_home
from yanjia_automation.screens.home import HomeScreen


@pytest.fixture()
def home(driver: WebDriver, settings: Settings) -> HomeScreen:
    return restart_to_home(driver, settings)
