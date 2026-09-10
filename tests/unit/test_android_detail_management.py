from __future__ import annotations

from scripts.prepare_excel import CURRENT_STEPS


def test_unselected_image_delete_uses_exact_prompt_without_selecting_image() -> None:
    steps = CURRENT_STEPS["TC-DETAIL-029"]
    selectors = "id=com.xiaofutech.yanjia_ai:id/a_records_detail_all_select_ifv"
    delete_button = "id=com.xiaofutech.yanjia_ai:id/customer_detail_delete_tv"
    prompt = "请选择需要删除的影像"

    assert len(steps) == 20
    assert steps[3][3] == "咨询单特定"
    assert steps[5][4:6] == ("element_count_equals", "1")
    assert not any(action == "click" and locator == selectors for _, action, locator, *_ in steps)

    delete_clicks = [
        step for step in steps if step[1] == "click" and step[2] == delete_button
    ]
    assert len(delete_clicks) == 1
    assert delete_clicks[0][0] == "未选择影像直接点击删除"

    prompt_step = steps[15]
    assert prompt_step[0] == "校验未选择影像提示"
    assert prompt_step[2] is None
    assert prompt_step[4:6] == ("page_source_contains", prompt)
