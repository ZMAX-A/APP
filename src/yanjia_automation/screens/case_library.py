from __future__ import annotations

from yanjia_automation.screens.base import BaseScreen, resource_id


class CaseLibraryScreen(BaseScreen):
    root = resource_id("case_rv")
    search_input = resource_id("case_search_et")
    search_button = resource_id("case_search_tv")
    tags = resource_id("a_case_tag_tv")
    tag_expand = resource_id("case_tag_expand_tv")
    categories = resource_id("a_case_category_tv")
    cards = resource_id("a_case_image_cl")
    empty_state = resource_id("empty_tv")
    manage_button = resource_id("case_manager_tv")
