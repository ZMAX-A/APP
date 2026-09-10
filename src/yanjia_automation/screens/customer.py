from __future__ import annotations

from datetime import date

from appium.webdriver.common.appiumby import AppiumBy
from appium.webdriver.webelement import WebElement
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait

from yanjia_automation.date_picker import date_after_picker_change, swipe_number_picker_steps
from yanjia_automation.screens.base import BaseScreen, resource_id


class CustomerListScreen(BaseScreen):
    root = resource_id("customer_records_rv")
    search_input = resource_id("customer_records_search_et")
    search_button = resource_id("customer_records_search_tv")
    date_filter = resource_id("customer_records_date_ll")
    other_filter = resource_id("customer_records_other_tv")
    other_filter_panel = resource_id("customer_records_other_cl")
    gender_male = resource_id("customer_records_other_sex_man_tv")
    gender_female = resource_id("customer_records_other_sex_female_tv")
    age_under_18 = resource_id("customer_records_other_age_0_tv")
    age_18_25 = resource_id("customer_records_other_age_1_tv")
    age_26_35 = resource_id("customer_records_other_age_2_tv")
    age_36_45 = resource_id("customer_records_other_age_3_tv")
    age_45_plus = resource_id("customer_records_other_age_4_tv")
    remark_input = resource_id("customer_records_other_remark_et")
    filter_reset = resource_id("customer_records_other_reset_tv")
    filter_confirm = resource_id("customer_records_other_confirm_tv")
    date_overlay = resource_id("customer_records_date_v")
    date_panel = resource_id("customer_records_date_cl")
    date_cancel = resource_id("customer_records_date_cancel_tv")
    date_confirm = resource_id("customer_records_date_confirm_tv")
    start_date_picker = resource_id("customer_records_date_start_dpv")
    end_date_picker = resource_id("customer_records_date_end_dpv")
    date_year_wheels = resource_id("date_picker_year_wheel")
    date_month_wheels = resource_id("date_picker_month_wheel")
    date_day_wheels = resource_id("date_picker_day_wheel")
    cards = resource_id("a_records_cl")
    card_name = resource_id("a_records_name_tv")
    card_age = resource_id("a_records_age_tv")
    card_identifier = resource_id("a_records_unique_tv")
    card_last_detection = resource_id("a_records_time_tv")
    card_detection_count = resource_id("a_records_count_tv")
    empty_state = resource_id("empty_tv")

    def open_first_customer(self) -> None:
        cards = self.find_all(self.cards)
        if not cards:
            raise AssertionError("Customer seed data is required but the list was empty")
        cards[0].click()

    def search_customers(
        self,
        query: str,
        *,
        timeout: float | None = None,
        expected_count: int | None = None,
    ) -> list[WebElement]:
        search = self.find(self.search_input, timeout)
        search.clear()
        search.send_keys(query)
        self.click(self.search_button, timeout)
        wait_timeout = timeout or self.timeout
        if expected_count is None:
            WebDriverWait(self.driver, wait_timeout).until(
                lambda _: (
                    self.find_all(self.cards) or self.is_visible(self.empty_state, timeout=0.2)
                )
            )
        else:
            try:
                WebDriverWait(self.driver, wait_timeout).until(
                    lambda _: len(self.find_all(self.cards)) == expected_count
                )
            except TimeoutException:
                # Return the final observed collection so the caller can apply
                # its redacted unique-match safety error without exposing data.
                pass
        return self.find_all(self.cards)


class CustomerDetailScreen(BaseScreen):
    root = resource_id("customer_detail_name_tv")
    customer_info = resource_id("customer_detail_info_tv")
    edit_button = resource_id("customer_detail_edit_tv")
    manage_button = resource_id("customer_detail_manager_tv")
    delete_button = resource_id("customer_detail_delete_tv")
    images = resource_id("a_records_detail_all_head_ifv")
    image_remarks = resource_id("a_records_detail_all_remark_ifv")
    image_dates = resource_id("a_records_detail_all_time_tv")
    image_selectors = resource_id("a_records_detail_all_select_ifv")
    consultation_results = resource_id("a_consultation_result_cl")
    consultation_remarks = resource_id("a_consultation_result_remark_ifv")
    consultation_detection_times = resource_id("a_consultation_result_time_2_tv")
    empty_state = resource_id("empty_tv")

    def open_first_image(self) -> None:
        images = self.find_all(self.images)
        if not images:
            raise AssertionError("Image seed data is required but the customer had no image")
        images[0].click()

    def open_editor(self) -> None:
        self.click(self.edit_button)


class CustomerEditScreen(BaseScreen):
    root = resource_id("customer_edit_v")
    close_button = resource_id("customer_edit_close_ifv")
    username = resource_id("customer_edit_username_et")
    phone = resource_id("customer_edit_unique_et")
    gender = resource_id("customer_edit_sex_tv")
    gender_male = resource_id("customer_edit_sex_man_tv")
    gender_female = resource_id("customer_edit_sex_female_tv")
    birthday = resource_id("customer_edit_birthday_tv")
    birthday_confirm = resource_id("customer_edit_birthday_confirm_tv")
    birthday_year_wheels = resource_id("date_picker_year_wheel")
    birthday_month_wheels = resource_id("date_picker_month_wheel")
    birthday_day_wheels = resource_id("date_picker_day_wheel")
    email = resource_id("customer_edit_email_et")
    marital = resource_id("customer_edit_marital_tv")
    marital_secret = resource_id("customer_edit_marital_secret_tv")
    marital_unmarried = resource_id("customer_edit_marital_no_tv")
    marital_married = resource_id("customer_edit_marital_yes_tv")
    address = resource_id("customer_edit_address_et")
    remark = resource_id("customer_edit_remark_et")
    profile_scroll_view = (AppiumBy.CLASS_NAME, "android.widget.ScrollView")
    save_button = resource_id("customer_edit_save_tv")

    def clear_username(self) -> None:
        self.find(self.username).clear()

    def clear_phone(self) -> None:
        self.find(self.phone).clear()

    def save(self) -> None:
        self.click(self.save_button)

    def close(self) -> None:
        self.click(self.close_button)

    def profile_values(self) -> tuple[str, str]:
        name = self.find(self.username).get_attribute("text")
        phone = self.find(self.phone).get_attribute("text")
        return str(name or ""), str(phone or "")

    def gender_value(self) -> str:
        value = self.find(self.gender).get_attribute("text")
        return str(value or "")

    def birthday_value(self) -> str:
        value = self.find(self.birthday).get_attribute("text")
        return str(value or "")

    def address_value(self) -> str:
        value = self.find(self.address).get_attribute("text")
        return str(value or "")

    def email_value(self) -> str:
        value = self.find(self.email).get_attribute("text")
        return str(value or "")

    def marital_value(self) -> str:
        value = self.find(self.marital).get_attribute("text")
        return str(value or "")

    def remark_value(self) -> str:
        field = self.scroll_to_remark()
        value = field.get_attribute("text")
        return str(value or "")

    def select_gender(self, gender: str) -> None:
        options = {"男": self.gender_male, "女": self.gender_female}
        try:
            option = options[gender]
        except KeyError as error:
            raise ValueError("Android 顾客编辑页仅支持男或女") from error

        self.click(self.gender)
        self.click(option)
        WebDriverWait(self.driver, self.timeout).until(lambda _: self.gender_value() == gender)

    def select_birthday(self, birthday: str) -> None:
        target = date.fromisoformat(birthday)
        current = date.fromisoformat(self.birthday_value())

        self.click(self.birthday)
        self._scroll_date_wheel(self.birthday_year_wheels, current.year, target.year)
        current = date_after_picker_change(current, year=target.year)
        self._scroll_date_wheel(self.birthday_month_wheels, current.month, target.month)
        current = date_after_picker_change(current, month=target.month)
        self._scroll_date_wheel(self.birthday_day_wheels, current.day, target.day)
        self.click(self.birthday_confirm)
        WebDriverWait(self.driver, self.timeout).until(
            lambda _: self.birthday_value() == target.isoformat()
        )

    def set_address(self, address: str) -> None:
        field = self.find(self.address)
        field.clear()
        if address:
            field.send_keys(address)

    def set_email(self, email: str) -> None:
        field = self.find(self.email)
        field.clear()
        if email:
            field.send_keys(email)

    def select_marital(self, marital: str) -> None:
        options = {
            "保密": self.marital_secret,
            "未婚": self.marital_unmarried,
            "已婚": self.marital_married,
        }
        try:
            option = options[marital]
        except KeyError as error:
            raise ValueError("Android 顾客编辑页婚姻状态仅支持保密、未婚或已婚") from error

        self.click(self.marital)
        self.click(option)
        WebDriverWait(self.driver, self.timeout).until(lambda _: self.marital_value() == marital)

    def scroll_to_top(self) -> WebElement:
        return self._scroll_profile_to(self.username, direction="up")

    def scroll_to_remark(self) -> WebElement:
        return self._scroll_profile_to(self.remark, direction="down")

    def set_remark(self, remark: str) -> None:
        field = self.scroll_to_remark()
        field.clear()
        if remark:
            field.send_keys(remark)

    def _scroll_profile_to(
        self,
        locator: tuple[str, str],
        *,
        direction: str,
    ) -> WebElement:
        visible = [element for element in self.find_all(locator) if element.is_displayed()]
        if visible:
            return visible[0]

        scroll_view = self.find(self.profile_scroll_view)
        for _ in range(3):
            self.driver.execute_script(
                "mobile: scrollGesture",
                {
                    "elementId": scroll_view.id,
                    "direction": direction,
                    "percent": 0.75,
                },
            )
            visible = [element for element in self.find_all(locator) if element.is_displayed()]
            if visible:
                return visible[0]
        raise AssertionError("顾客资料表单滚动后仍未找到目标字段")

    def _scroll_date_wheel(
        self,
        locator: tuple[str, str],
        current_value: int,
        target_value: int,
    ) -> None:
        strategy, selector = locator
        if strategy != AppiumBy.ID:
            raise AssertionError("生日滚轮必须使用 Android resource-id")
        delta = target_value - current_value
        if delta == 0:
            return
        wheels = self.driver.find_elements(strategy, selector)
        if len(wheels) != 1:
            raise AssertionError(f"生日滚轮数量必须为1：{selector}")
        swipe_number_picker_steps(self.driver, wheels[0].rect, delta)

    def set_profile(
        self,
        *,
        name: str,
        phone: str,
        gender: str | None = None,
        birthday: str | None = None,
        email: str | None = None,
        marital: str | None = None,
        address: str | None = None,
        remark: str | None = None,
    ) -> None:
        self.scroll_to_top()
        username = self.find(self.username)
        username.clear()
        username.send_keys(name)

        unique_phone = self.find(self.phone)
        unique_phone.clear()
        unique_phone.send_keys(phone)

        if gender is not None and self.gender_value() != gender:
            self.select_gender(gender)
        if birthday is not None and self.birthday_value() != birthday:
            self.select_birthday(birthday)
        if email is not None and self.email_value() != email:
            self.set_email(email)
        if marital is not None and self.marital_value() != marital:
            self.select_marital(marital)
        if address is not None and self.address_value() != address:
            self.set_address(address)
        if remark is not None and self.remark_value() != remark:
            self.set_remark(remark)


class RecordsRemarkScreen(BaseScreen):
    root = resource_id("records_remark_remark_v")
    close_button = resource_id("records_remark_close_ifv")
    remark_input = resource_id("records_remark_remark_et")
    submit_button = resource_id("records_remark_remark_commit_tv")
    tag_add_button = resource_id("a_records_remark_tag_add_ll")
    tag_dialog = resource_id("records_remark_tag_cl")
    tag_input = resource_id("records_remark_tag_et")
    tag_save_button = resource_id("records_remark_tag_save_tv")
    tag_cancel_button = resource_id("records_remark_tag_cancel_tv")
    records = resource_id("records_remark_records_rv")
    record_time = resource_id("a_records_remark_list_time_tv")
    record_content = resource_id("a_records_remark_list_content_tv")
    record_delete = resource_id("a_records_remark_list_delete_ifv")
