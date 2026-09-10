from __future__ import annotations

from yanjia_automation.screens.base import BaseScreen, resource_id


class HomeScreen(BaseScreen):
    root = resource_id("main_fbl")
    greeting = resource_id("main_name_tv")
    search_input = resource_id("main_search_et")
    search_button = resource_id("main_search_tv")
    customer_records = resource_id("main_records_ll")
    academy = resource_id("main_meiji_ll")
    case_library = resource_id("main_case_ll")
    settings = resource_id("main_set_cl")

    def open_customer_records(self) -> None:
        self.click(self.customer_records)

    def open_case_library(self) -> None:
        self.click(self.case_library)

    def open_academy(self) -> None:
        self.click(self.academy)

    def open_search(self) -> None:
        self.click(self.search_button)

    def open_settings(self) -> None:
        self.click(self.settings)
