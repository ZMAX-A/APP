from __future__ import annotations

from yanjia_automation.screens.base import BaseScreen, resource_id


class ImageViewerScreen(BaseScreen):
    root = resource_id("skin_result_back_ll")
    surface = resource_id("skin_result_l3d")
    pager = resource_id("skin_result_vp2")
    download_progress = resource_id("skin_result_download_pb")
