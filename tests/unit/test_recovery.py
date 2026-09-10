from __future__ import annotations

from unittest.mock import Mock, call

from appium.webdriver.webdriver import WebDriver
from appium.webdriver.webelement import WebElement

from yanjia_automation.flows import recovery
from yanjia_automation.screens.home import HomeScreen


def _visible_element(*, text: str = "") -> Mock:
    element = Mock(spec=WebElement)
    element.get_attribute.return_value = text
    element.is_displayed.return_value = True
    element.is_enabled.return_value = True
    return element


def test_credential_save_prompt_is_cancelled() -> None:
    driver = Mock(spec=WebDriver)
    title = _visible_element(text="是否保存“妍家AI”的账号和密码？")
    cancel = _visible_element(text="取消")
    driver.find_elements.side_effect = [[title], [cancel]]

    assert recovery._dismiss_android_credential_save_prompt(driver) is True
    cancel.click.assert_called_once_with()


def test_miui_credential_save_prompt_is_cancelled() -> None:
    driver = Mock(spec=WebDriver)
    title = _visible_element(text="保存“颜佳AI”的账号密码？")
    cancel = _visible_element(text="取消")
    driver.find_elements.side_effect = [[], [title], [cancel]]

    assert recovery._dismiss_android_credential_save_prompt(driver) is True
    assert driver.find_elements.call_args_list == [
        call(*recovery._ANDROID_ALERT_TITLE),
        call(*recovery._OEM_ALERT_TITLE),
        call(*recovery._ANDROID_ALERT_CANCEL),
    ]
    cancel.click.assert_called_once_with()


def test_unrelated_system_prompt_is_left_open() -> None:
    driver = Mock(spec=WebDriver)
    title = _visible_element(text="妍家AI无响应")
    driver.find_elements.return_value = [title]

    assert recovery._dismiss_android_credential_save_prompt(driver) is False
    driver.find_elements.assert_called_once_with(*recovery._ANDROID_ALERT_TITLE)


def test_account_error_prompt_is_not_mistaken_for_save_prompt() -> None:
    driver = Mock(spec=WebDriver)
    title = _visible_element(text="账号或密码错误")
    driver.find_elements.return_value = [title]

    assert recovery._dismiss_android_credential_save_prompt(driver) is False
    driver.find_elements.assert_called_once_with(*recovery._ANDROID_ALERT_TITLE)


def test_wait_for_home_dismisses_prompt_before_waiting_for_root(
    monkeypatch,
) -> None:
    driver = Mock(spec=WebDriver)
    home = Mock(spec=HomeScreen)
    home.root = HomeScreen.root
    home.is_visible.side_effect = [False, True]
    home.wait_loaded.return_value = home
    dismiss = Mock(return_value=True)
    monkeypatch.setattr(recovery, "_dismiss_android_credential_save_prompt", dismiss)

    result = recovery._wait_for_home(driver, home, timeout=2)

    assert result is home
    dismiss.assert_called_once_with(driver)
    home.wait_loaded.assert_called_once_with(timeout=2)


def test_recover_login_if_needed_signs_in_and_waits_for_home(monkeypatch) -> None:
    driver = Mock(spec=WebDriver)
    driver.current_activity = ".activity.LoginActivity"
    login = Mock()
    home = Mock(spec=HomeScreen)
    settings = Mock()
    credentials = object()
    settings.credentials.return_value = credentials
    settings.store_name = "测试门店"
    wait_for_home = Mock(return_value=home)

    monkeypatch.setattr(recovery, "LoginScreen", Mock(return_value=login))
    monkeypatch.setattr(recovery, "HomeScreen", Mock(return_value=home))
    monkeypatch.setattr(recovery, "_wait_for_home", wait_for_home)

    assert recovery.recover_login_if_needed(driver, settings, timeout=12) is True
    login.sign_in.assert_called_once_with(credentials, "测试门店")
    wait_for_home.assert_called_once_with(driver, home, timeout=12)


def test_recover_login_if_needed_leaves_non_login_page_unchanged(monkeypatch) -> None:
    driver = Mock(spec=WebDriver)
    driver.current_activity = ".activity.CustomerRecordsActivity"
    login = Mock()
    login.root = ("id", "login")
    login.is_visible.return_value = False
    settings = Mock()

    monkeypatch.setattr(recovery, "LoginScreen", Mock(return_value=login))

    assert recovery.recover_login_if_needed(driver, settings) is False
    login.sign_in.assert_not_called()
