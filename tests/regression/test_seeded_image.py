from __future__ import annotations

import pytest

from yanjia_automation.screens.customer import CustomerDetailScreen, CustomerListScreen
from yanjia_automation.screens.home import HomeScreen
from yanjia_automation.screens.image_viewer import ImageViewerScreen

pytestmark = [
    pytest.mark.tablet,
    pytest.mark.readonly,
    pytest.mark.requires_auth,
    pytest.mark.requires_seed,
]


@pytest.mark.p0
@pytest.mark.extended
@pytest.mark.case_id("TC-IMAGE-001")
def test_existing_image_opens_viewer(home: HomeScreen) -> None:
    home.open_customer_records()
    customer_list = CustomerListScreen(home.driver).wait_loaded(timeout=15)
    customer_list.open_first_customer()
    detail = CustomerDetailScreen(home.driver).wait_loaded(timeout=15)
    detail.open_first_image()

    viewer = ImageViewerScreen(home.driver).wait_loaded(timeout=20)
    assert home.driver.current_activity.endswith(".SkinResultActivity")
    assert viewer.find(viewer.pager).is_displayed()
