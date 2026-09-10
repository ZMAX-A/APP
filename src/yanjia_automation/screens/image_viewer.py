from __future__ import annotations

from yanjia_automation.screens.base import BaseScreen, resource_id


class ImageViewerScreen(BaseScreen):
    root = resource_id("skin_result_back_ll")
    surface = resource_id("skin_result_l3d")
    pager = resource_id("skin_result_vp2")
    download_progress = resource_id("skin_result_download_pb")
    view_case = resource_id("f_skin_result_info_view_case_tv")
    menu = resource_id("f_skin_result_info_menu_ifv")
    report_menu = resource_id("f_skin_result_info_menu_report_ll")
    discard_exit = resource_id("cover_prompt_v1_tv")
    comprehensive_scores = resource_id("f_skin_result_info_comprehensive_rv")
    radar_score_names = resource_id("a_skin_result_info_radar_child_name_tv")
    radar_scores = resource_id("a_skin_result_info_radar_child_score_tv")
    feature_score_names = resource_id("a_skin_result_info_feature_child_name_tv")
    feature_scores = resource_id("a_skin_result_info_feature_child_score_tv")


class SkinReportScreen(BaseScreen):
    root = resource_id("skin_report_pdfv")
    back = resource_id("header_back_ll")
    print_action = resource_id("skin_report_print_tv")
    send_action = resource_id("skin_report_send_tv")
    share_action = resource_id("skin_report_share_tv")
