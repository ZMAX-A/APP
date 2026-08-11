from __future__ import annotations

from selenium.webdriver.support.ui import WebDriverWait

from yanjia_automation.config import Credentials
from yanjia_automation.screens.base import BaseScreen, resource_id


class LoginScreen(BaseScreen):
    root = resource_id("login_username_et")
    username = resource_id("login_username_et")
    password = resource_id("login_pwd_et")
    submit = resource_id("login_tv")
    store_list = resource_id("login_store_rv")
    store_names = resource_id("a_login_store_tv")
    enter_store_buttons = resource_id("a_login_join_tv")
    error_confirm = resource_id("cover_prompt_v1_tv")

    def sign_in(self, credentials: Credentials, store_name: str | None = None) -> None:
        if self.is_visible(self.error_confirm, timeout=0.5):
            self.click(self.error_confirm)

        username = self.find(self.username)
        username.clear()
        username.send_keys(credentials.username)

        password = self.find(self.password)
        password.clear()
        password.send_keys(credentials.password)
        self.click(self.submit)

        WebDriverWait(self.driver, 20).until(
            lambda _: self.is_visible(self.store_list, timeout=0.2)
            or self.driver.current_activity.endswith(".MainActivity")
            or self.is_visible(self.error_confirm, timeout=0.2)
        )

        if self.is_visible(self.error_confirm, timeout=0.5):
            raise AssertionError("登录失败：应用提示账号或密码不正确")

        if not self.is_visible(self.store_list, timeout=0.5):
            return

        buttons = self.find_all(self.enter_store_buttons)
        if not buttons:
            raise AssertionError("The account returned no selectable store")

        index = 0
        if store_name:
            stores = self.find_all(self.store_names)
            matches = [position for position, item in enumerate(stores) if item.text == store_name]
            if not matches:
                raise AssertionError("Configured store was not present in the store picker")
            index = matches[0]

        buttons[index].click()
