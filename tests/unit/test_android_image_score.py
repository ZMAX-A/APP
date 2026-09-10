from __future__ import annotations

from scripts.prepare_excel import CURRENT_STEPS, image_score_value
from yanjia_automation.screens import ImageViewerScreen


def _step(name: str) -> tuple[object, ...]:
    return next(step for step in CURRENT_STEPS["TC-IMAGE-030"] if step[0] == name)


def test_image_score_case_uses_unique_controlled_seed_and_readonly_contract() -> None:
    steps = CURRENT_STEPS["TC-IMAGE-030"]
    assert len(steps) == 19
    assert _step("输入影像导航种子顾客")[3] == "咨询单特定"
    assert _step("校验种子顾客唯一匹配")[4:6] == ("element_count_equals", "1")
    assert _step("等待影像操作栏就绪")[6] == 60
    assert _step("校验AI综合分析评分区")[2] == (
        "id=com.xiaofutech.yanjia_ai:id/f_skin_result_info_comprehensive_rv"
    )
    assert ImageViewerScreen.comprehensive_scores[1].endswith(
        "f_skin_result_info_comprehensive_rv"
    )


def test_image_score_xpath_anchors_each_label_to_its_own_nonempty_score() -> None:
    expected = {
        "炎敏": "radar",
        "肤色": "radar",
        "色素": "feature",
        "亮度": "feature",
        "红区": "feature",
    }
    for label, group in expected.items():
        locator = image_score_value(label, group)
        assert locator.startswith("xpath=//*[@resource-id=")
        assert f"child_name_tv' and @text='{label}']/parent::*" in locator
        assert locator.endswith(f"a_skin_result_info_{group}_child_score_tv']")
        step = _step(f"校验{label}评分非空")
        assert step[2] == locator
        assert step[4] == "text_not_empty"
        assert "不记录评分值" in str(step[10])


def test_image_score_screen_exposes_probed_native_collections() -> None:
    assert ImageViewerScreen.radar_score_names[1].endswith(
        "a_skin_result_info_radar_child_name_tv"
    )
    assert ImageViewerScreen.radar_scores[1].endswith(
        "a_skin_result_info_radar_child_score_tv"
    )
    assert ImageViewerScreen.feature_score_names[1].endswith(
        "a_skin_result_info_feature_child_name_tv"
    )
    assert ImageViewerScreen.feature_scores[1].endswith(
        "a_skin_result_info_feature_child_score_tv"
    )
