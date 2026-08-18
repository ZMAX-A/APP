from __future__ import annotations

from yanjia_automation.screens.base import BaseScreen, resource_id


class SettingsScreen(BaseScreen):
    root = resource_id("set_logout_tv")
    personal_center = resource_id("set_personal_ll")
    version = resource_id("set_version_tv")
    logout = resource_id("set_logout_tv")
    confirm_logout = resource_id("cover_prompt_right_tv")

    def open_personal_center(self) -> None:
        self.click(self.personal_center)

    def logout_and_confirm(self) -> None:
        self.click(self.logout)
        self.click(self.confirm_logout)


class ProfileScreen(BaseScreen):
    root = resource_id("personal_account_ll")
    account = resource_id("personal_account_tv")
    name = resource_id("personal_name_tv")
    registered_at = resource_id("personal_register_time_tv")
    phone = resource_id("personal_phone_tv")
