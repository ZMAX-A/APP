from __future__ import annotations

from scripts.prepare_excel import CURRENT_STEPS
from yanjia_automation.screens import ImageViewerScreen, SkinReportScreen


def _step(case_id: str, name: str) -> tuple[object, ...]:
    return next(step for step in CURRENT_STEPS[case_id] if step[0] == name)


def test_image_navigation_cases_use_unique_controlled_seed() -> None:
    for case_id in ("TC-IMAGE-032", "TC-IMAGE-033", "TC-IMAGE-034"):
        steps = CURRENT_STEPS[case_id]
        unique_seed = _step(case_id, "校验种子顾客唯一匹配")
        assert unique_seed[2] == "id=com.xiaofutech.yanjia_ai:id/a_records_cl"
        assert unique_seed[4:6] == ("element_count_equals", "1")
        assert _step(case_id, "输入影像导航种子顾客")[3] == "咨询单特定"
        assert _step(case_id, "等待影像操作栏就绪")[6] == 60
        assert len(steps) >= 16


def test_view_case_navigation_uses_probed_native_contract() -> None:
    assert ImageViewerScreen.view_case[1].endswith("f_skin_result_info_view_case_tv")
    assert _step("TC-IMAGE-032", "校验案例库Activity")[5] == ".CaseActivity"
    assert str(_step("TC-IMAGE-032", "校验案例列表")[2]).endswith("case_rv")


def test_report_navigation_uses_current_report_contract() -> None:
    assert ImageViewerScreen.menu[1].endswith("f_skin_result_info_menu_ifv")
    assert ImageViewerScreen.report_menu[1].endswith("f_skin_result_info_menu_report_ll")
    assert SkinReportScreen.root[1].endswith("skin_report_pdfv")
    assert _step("TC-IMAGE-033", "校验报告Activity")[5] == ".SkinReportActivity"
    assert str(_step("TC-IMAGE-033", "校验报告分享入口")[2]).endswith(
        "skin_report_share_tv"
    )


def test_discard_exit_returns_to_customer_detail_without_saving() -> None:
    discard = _step("TC-IMAGE-034", "校验退出确认提示")
    assert ImageViewerScreen.discard_exit[1].endswith("cover_prompt_v1_tv")
    assert discard[4:6] == ("text_equals", "不保存直接退出")
    assert _step("TC-IMAGE-034", "校验返回顾客详情Activity")[5] == (
        ".CustomerDetailActivity"
    )
