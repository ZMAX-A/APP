from __future__ import annotations

from yanjia_automation.screens.base import BaseScreen, resource_id


class CustomerListScreen(BaseScreen):
    root = resource_id("customer_records_rv")
    search_input = resource_id("customer_records_search_et")
    search_button = resource_id("customer_records_search_tv")
    date_filter = resource_id("customer_records_date_ll")
    other_filter = resource_id("customer_records_other_tv")
    cards = resource_id("a_records_cl")
    card_name = resource_id("a_records_name_tv")
    card_age = resource_id("a_records_age_tv")
    card_identifier = resource_id("a_records_unique_tv")
    card_last_detection = resource_id("a_records_time_tv")
    card_detection_count = resource_id("a_records_count_tv")

    def open_first_customer(self) -> None:
        cards = self.find_all(self.cards)
        if not cards:
            raise AssertionError("Customer seed data is required but the list was empty")
        cards[0].click()


class CustomerDetailScreen(BaseScreen):
    root = resource_id("customer_detail_name_tv")
    customer_info = resource_id("customer_detail_info_tv")
    edit_button = resource_id("customer_detail_edit_tv")
    manage_button = resource_id("customer_detail_manager_tv")
    images = resource_id("a_records_detail_all_head_ifv")
    image_dates = resource_id("a_records_detail_all_time_tv")
    empty_state = resource_id("empty_tv")

    def open_first_image(self) -> None:
        images = self.find_all(self.images)
        if not images:
            raise AssertionError("Image seed data is required but the customer had no image")
        images[0].click()
