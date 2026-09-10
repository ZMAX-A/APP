from __future__ import annotations

import argparse
import shutil
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

from yanjia_automation.excel.android_catalog import ANDROID_STEP_SHEET_CASE_IDS

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_WORKBOOK = ROOT / "test_case.xlsx"
CASE_SHEET = "自动化测试用例"
STEP_SHEET = "自动化执行步骤"
KEYWORD_SHEET = "自动化关键字说明"

CASE_RESULT_COLUMNS = (
    "是否执行",
    "最后执行时间",
    "执行耗时(秒)",
    "错误信息",
    "运行编号",
    "Allure报告目录",
)

STEP_HEADERS = (
    "用例ID",
    "步骤序号",
    "步骤名称",
    "操作类型",
    "元素定位器",
    "输入数据",
    "断言类型",
    "期望值",
    "超时(秒)",
    "元素索引",
    "失败继续",
    "是否执行",
    "备注",
)


def rid(name: str) -> str:
    return f"id=com.xiaofutech.yanjia_ai:id/{name}"


def exact_case_tag(tag: str) -> str:
    return f"xpath=//*[@resource-id='com.xiaofutech.yanjia_ai:id/a_case_tag_tv' and @text='{tag}']"


def case_card_tag(tag: str) -> str:
    return (
        f"xpath=//*[@resource-id='com.xiaofutech.yanjia_ai:id/a_case_image_cl']//*[@text='{tag}']"
    )


def image_score_value(label: str, group: str) -> str:
    base = "com.xiaofutech.yanjia_ai:id/"
    return (
        f"xpath=//*[@resource-id='{base}a_skin_result_info_{group}_child_name_tv' "
        f"and @text='{label}']/parent::*"
        f"//*[@resource-id='{base}a_skin_result_info_{group}_child_score_tv']"
    )


def s(
    name: str,
    action: str,
    locator: str | None = None,
    input_value: str | None = None,
    assertion: str | None = None,
    expected: str | None = None,
    timeout: int = 10,
    index: int = 0,
    continue_on_failure: str = "否",
    note: str | None = None,
) -> tuple[Any, ...]:
    return (
        name,
        action,
        locator,
        input_value,
        assertion,
        expected,
        timeout,
        index,
        continue_on_failure,
        "是",
        note,
    )


def controlled_image_entry_steps() -> list[tuple[Any, ...]]:
    """Open the first image for the unique, non-personal image-navigation seed."""

    return [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "输入影像导航种子顾客",
            "input",
            rid("customer_records_search_et"),
            "咨询单特定",
        ),
        s("执行顾客搜索", "click", rid("customer_records_search_tv")),
        s(
            "校验种子顾客唯一匹配",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_equals",
            expected="1",
            timeout=15,
        ),
        s("打开唯一顾客", "click", rid("a_records_cl"), timeout=15),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s(
            "校验存在历史影像",
            "assert",
            rid("a_records_detail_all_head_ifv"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
        s(
            "打开第一张历史影像",
            "click",
            rid("a_records_detail_all_head_ifv"),
            timeout=15,
        ),
        s(
            "校验影像Activity",
            "assert",
            assertion="activity_endswith",
            expected=".SkinResultActivity",
            timeout=20,
        ),
        s(
            "校验影像翻页控件",
            "assert",
            rid("skin_result_vp2"),
            assertion="element_visible",
            timeout=20,
        ),
        s(
            "等待影像操作栏就绪",
            "assert",
            rid("f_skin_result_info_view_case_tv"),
            assertion="element_visible",
            timeout=60,
        ),
    ]


def detail_consultation_create_steps() -> list[tuple[Any, ...]]:
    """Create and retain a consultation from the first currently unlinked image."""

    keep_note = "新增咨询单按用户要求保留，不执行删除或回滚"
    return [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "输入影像种子顾客",
            "input",
            rid("customer_records_search_et"),
            "咨询单特定",
        ),
        s("执行顾客搜索", "click", rid("customer_records_search_tv")),
        s(
            "校验影像种子顾客唯一匹配",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_equals",
            expected="1",
            timeout=15,
        ),
        s("打开唯一顾客卡片", "click", rid("a_records_cl"), timeout=15),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s(
            "校验存在历史影像",
            "assert",
            rid("a_records_detail_all_head_ifv"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
        s(
            "打开首张未关联咨询单的影像",
            "open_first_unlinked_image",
            timeout=15,
            note="按检测时间动态选择；若当前可见影像均已关联则安全失败",
        ),
        s(
            "校验进入影像结果页",
            "assert",
            assertion="activity_endswith",
            expected=".SkinResultActivity",
            timeout=15,
        ),
        s(
            "校验影像结果内容",
            "assert",
            rid("skin_result_vp2"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "校验咨询单入口就绪",
            "assert",
            rid("f_skin_result_info_view_case_tv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "返回并触发保存退出确认",
            "click",
            rid("skin_result_back_ll"),
        ),
        s(
            "校验保存并退出选项",
            "assert",
            rid("cover_prompt_v2_tv"),
            assertion="text_equals",
            expected="保存并退出",
        ),
        s(
            "保存咨询单并退出",
            "click",
            rid("cover_prompt_v2_tv"),
            timeout=15,
            note=keep_note,
        ),
        s(
            "校验返回顾客详情页",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s(
            "校验新增咨询单已保留",
            "verify_selected_image_consultation",
            timeout=15,
            note="以所选影像检测时间出现在咨询单卡片中为准",
        ),
    ]


def detail_gender_steps(gender: str) -> list[tuple[Any, ...]]:
    """Build a guarded dedicated-customer gender write and verification flow."""

    option = {
        "男": "customer_edit_sex_man_tv",
        "女": "customer_edit_sex_female_tv",
    }[gender]
    restore_note = "执行后强制恢复原姓名、手机号和性别，并重新打开编辑页逐字段比对"
    return [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "输入专用顾客查询",
            "input",
            rid("customer_records_search_et"),
            input_value="${YANJIA_MUTATION_CUSTOMER_QUERY}",
        ),
        s("执行专用顾客搜索", "click", rid("customer_records_search_tv")),
        s(
            "校验专用顾客唯一匹配",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_equals",
            expected="1",
            timeout=15,
        ),
        s("打开唯一顾客卡片", "click", rid("a_records_cl"), timeout=15),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s("进入详情编辑", "click", rid("customer_detail_edit_tv")),
        s("校验编辑表单", "assert", rid("customer_edit_v"), assertion="element_visible"),
        s("展开性别选项", "click", rid("customer_edit_sex_tv")),
        s(f"选择性别{gender}", "click", rid(option), note=restore_note),
        s("保存性别", "click", rid("customer_edit_save_tv"), note=restore_note),
        s("重新打开详情编辑", "click", rid("customer_detail_edit_tv"), timeout=15),
        s(
            "校验保存后的性别",
            "assert",
            rid("customer_edit_sex_tv"),
            assertion="text_equals",
            expected=gender,
            note=restore_note,
        ),
    ]


def detail_profile_entry_steps() -> list[tuple[Any, ...]]:
    """Open the guarded dedicated-customer editor used by profile writes."""

    return [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "输入专用顾客查询",
            "input",
            rid("customer_records_search_et"),
            input_value="${YANJIA_MUTATION_CUSTOMER_QUERY}",
        ),
        s("执行专用顾客搜索", "click", rid("customer_records_search_tv")),
        s(
            "校验专用顾客唯一匹配",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_equals",
            expected="1",
            timeout=15,
        ),
        s("打开唯一顾客卡片", "click", rid("a_records_cl"), timeout=15),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s("进入详情编辑", "click", rid("customer_detail_edit_tv")),
        s("校验编辑表单", "assert", rid("customer_edit_v"), assertion="element_visible"),
    ]


def detail_birthday_steps() -> list[tuple[Any, ...]]:
    """Build the reconciled birthday write, reopen, and verification flow."""

    birthday = "2022-08-04"
    restore_note = "原用例步骤中的03日与输入/验证点冲突，统一为04日；执行后恢复八项资料"
    return [
        *detail_profile_entry_steps(),
        s(
            "选择生日2022-08-04",
            "select_birthday",
            input_value=birthday,
            timeout=30,
            note=restore_note,
        ),
        s("保存生日", "click", rid("customer_edit_save_tv"), note=restore_note),
        s("重新打开详情编辑", "click", rid("customer_detail_edit_tv"), timeout=15),
        s(
            "校验保存后的生日",
            "assert",
            rid("customer_edit_birthday_tv"),
            assertion="text_equals",
            expected=birthday,
            note=restore_note,
        ),
    ]


def detail_address_steps(case_id: str) -> list[tuple[Any, ...]]:
    """Build guarded address writes whose original value is always restored."""

    restore_note = (
        "执行后恢复运行前姓名、手机号、性别、生日、邮箱、婚姻状态、地址和个人备注，"
        "并重新打开逐字段比对"
    )
    mutations = {
        "TC-DETAIL-008": ("清空地址", "clear", None, "attribute_equals", "text|"),
        "TC-DETAIL-009": (
            "输入地址特殊字符",
            "input",
            "……&&*……*&",
            "text_equals",
            "……&&*……*&",
        ),
        "TC-DETAIL-010": (
            "输入单个空格地址",
            "input",
            "${SPACE}",
            "attribute_equals",
            "text|",
        ),
    }
    name, action, input_value, assertion, expected = mutations[case_id]
    return [
        *detail_profile_entry_steps(),
        s(
            name,
            action,
            rid("customer_edit_address_et"),
            input_value=input_value,
            note=restore_note,
        ),
        s("保存地址", "click", rid("customer_edit_save_tv"), note=restore_note),
        s("重新打开详情编辑", "click", rid("customer_detail_edit_tv"), timeout=15),
        s(
            "校验保存后的地址",
            "assert",
            rid("customer_edit_address_et"),
            assertion=assertion,
            expected=expected,
            note=restore_note,
        ),
    ]


def detail_marital_steps() -> list[tuple[Any, ...]]:
    """Build a guarded marital-state write with eight-field restoration."""

    restore_note = (
        "执行后恢复运行前姓名、手机号、性别、生日、邮箱、婚姻状态、地址和个人备注，并逐字段比对"
    )
    return [
        *detail_profile_entry_steps(),
        s("展开婚姻状态选项", "click", rid("customer_edit_marital_tv"), note=restore_note),
        s(
            "选择婚姻状态保密",
            "click",
            rid("customer_edit_marital_secret_tv"),
            note=restore_note,
        ),
        s("保存婚姻状态", "click", rid("customer_edit_save_tv"), note=restore_note),
        s("重新打开详情编辑", "click", rid("customer_detail_edit_tv"), timeout=15),
        s(
            "校验保存后的婚姻状态",
            "assert",
            rid("customer_edit_marital_tv"),
            assertion="text_equals",
            expected="保密",
            note=restore_note,
        ),
    ]


def detail_empty_tag_steps() -> list[tuple[Any, ...]]:
    """Validate an empty image tag without creating or deleting customer data."""

    safety_note = "空值校验后取消弹窗，不创建标签，不执行删除"
    return [
        *detail_profile_entry_steps()[:8],
        s(
            "校验存在历史影像",
            "assert",
            rid("a_records_detail_all_remark_ifv"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
            note=safety_note,
        ),
        s(
            "进入第一张影像备注",
            "click",
            rid("a_records_detail_all_remark_ifv"),
            timeout=10,
            note=safety_note,
        ),
        s(
            "校验备注Activity",
            "assert",
            assertion="activity_endswith",
            expected=".RecordsRemarkActivity",
            timeout=15,
        ),
        s("打开新增标签弹窗", "click", rid("a_records_remark_tag_add_ll")),
        s(
            "校验标签弹窗",
            "assert",
            rid("records_remark_tag_cl"),
            assertion="element_visible",
        ),
        s(
            "空标签点击保存",
            "click",
            rid("records_remark_tag_save_tv"),
            note=safety_note,
        ),
        s(
            "校验空标签提示",
            "assert",
            assertion="page_source_contains",
            expected="请输入标签内容",
            note=safety_note,
        ),
        s("取消新增标签", "click", rid("records_remark_tag_cancel_tv")),
        s("关闭备注编辑器", "click", rid("records_remark_close_ifv")),
        s(
            "校验返回顾客详情",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
        ),
    ]


def detail_whitespace_tag_steps(case_id: str) -> list[tuple[Any, ...]]:
    """Validate image/consultation tag whitespace without persisting a tag."""

    if case_id not in {"TC-DETAIL-019", "TC-DETAIL-032", "TC-DETAIL-034"}:
        raise ValueError(f"Unsupported tag validation case: {case_id}")
    image_case = case_id == "TC-DETAIL-019"
    whitespace_case = case_id in {"TC-DETAIL-019", "TC-DETAIL-032"}
    safety_note = (
        "空格或空值保存仅校验请输入标签内容，随后取消弹窗；"
        "真机确认标签集合不变，不创建标签，不执行删除"
    )
    if image_case:
        entry_steps = detail_profile_entry_steps()[:8]
        seed_name = "校验存在历史影像"
        seed_locator = "a_records_detail_all_remark_ifv"
        open_name = "进入第一张影像备注"
        open_locator = "a_records_detail_all_remark_ifv"
    else:
        entry_steps = [
            s("恢复首页", "restart_to_home", timeout=30),
            s("进入顾客档案", "click", rid("main_records_ll")),
            s(
                "校验顾客列表",
                "assert",
                rid("customer_records_rv"),
                assertion="element_visible",
                timeout=15,
            ),
            s(
                "输入咨询单种子顾客",
                "input",
                rid("customer_records_search_et"),
                "咨询单特定",
            ),
            s("执行顾客搜索", "click", rid("customer_records_search_tv")),
            s(
                "校验咨询单种子顾客存在",
                "assert",
                rid("a_records_cl"),
                assertion="element_count_gte",
                expected="1",
                timeout=15,
            ),
            s("点击第一张顾客卡片", "click", rid("a_records_cl"), timeout=15),
            s(
                "校验顾客详情Activity",
                "assert",
                assertion="activity_endswith",
                expected=".CustomerDetailActivity",
                timeout=15,
            ),
        ]
        seed_name = "校验存在咨询单"
        seed_locator = "a_consultation_result_cl"
        open_name = "进入第一条咨询单备注"
        open_locator = "a_consultation_result_remark_ifv"

    steps = [
        *entry_steps,
        s(
            seed_name,
            "assert",
            rid(seed_locator),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
            note=safety_note,
        ),
        s(open_name, "click", rid(open_locator), timeout=10, note=safety_note),
        s(
            "校验备注Activity",
            "assert",
            assertion="activity_endswith",
            expected=".RecordsRemarkActivity",
            timeout=15,
        ),
        s("打开新增标签弹窗", "click", rid("a_records_remark_tag_add_ll")),
        s(
            "校验标签弹窗",
            "assert",
            rid("records_remark_tag_cl"),
            assertion="element_visible",
        ),
    ]
    if whitespace_case:
        steps.append(
            s(
                "输入单个空格标签",
                "input",
                rid("records_remark_tag_et"),
                "${SPACE}",
                note=safety_note,
            )
        )
    steps.extend(
        [
            s(
                "空格标签点击保存" if whitespace_case else "空标签点击保存",
                "click",
                rid("records_remark_tag_save_tv"),
                note=safety_note,
            ),
            s(
                "校验空标签提示",
                "assert",
                assertion="page_source_contains",
                expected="请输入标签内容",
                note=safety_note,
            ),
            s("取消新增标签", "click", rid("records_remark_tag_cancel_tv")),
            s("关闭备注编辑器", "click", rid("records_remark_close_ifv")),
            s(
                "校验返回顾客详情",
                "assert",
                assertion="activity_endswith",
                expected=".CustomerDetailActivity",
            ),
        ]
    )
    return steps


def detail_tag_roundtrip_steps(case_id: str) -> list[tuple[Any, ...]]:
    """Create one unique tag, then restore the exact pre-run tag collection."""

    labels = {
        "TC-DETAIL-016": "自动化标签${RUN_TOKEN}",
        "TC-DETAIL-018": "%……&*${RUN_TOKEN}",
        "TC-DETAIL-020": "自动化删除标签${RUN_TOKEN}",
        "TC-DETAIL-031": "咨询单标签${RUN_TOKEN}",
        "TC-DETAIL-033": "%……&*${RUN_TOKEN}",
    }
    if case_id not in labels:
        raise ValueError(f"Unsupported tag roundtrip case: {case_id}")
    image_case = case_id in {"TC-DETAIL-016", "TC-DETAIL-018", "TC-DETAIL-020"}
    entry_locator = (
        "a_records_detail_all_remark_ifv" if image_case else "a_consultation_result_remark_ifv"
    )
    safety_note = (
        "保存前记录标签集合；断言结束或异常退出时仅删除本轮差集新增标签，"
        "并验证原标签文本、数量和可删除控件数量全部恢复"
    )
    steps = [
        *detail_profile_entry_steps()[:8],
        s(
            "校验存在历史影像" if image_case else "校验存在咨询单",
            "assert",
            rid(entry_locator),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
            note=safety_note,
        ),
        s(
            "进入第一张影像备注" if image_case else "进入第一条咨询单备注",
            "click",
            rid(entry_locator),
            timeout=10,
            note=safety_note,
        ),
        s(
            "校验备注Activity",
            "assert",
            assertion="activity_endswith",
            expected=".RecordsRemarkActivity",
            timeout=15,
        ),
        s(
            "打开新增标签弹窗",
            "click",
            rid("a_records_remark_tag_add_ll"),
            note=safety_note,
        ),
        s(
            "校验标签弹窗",
            "assert",
            rid("records_remark_tag_cl"),
            assertion="element_visible",
        ),
        s(
            "输入本轮唯一标签",
            "input",
            rid("records_remark_tag_et"),
            labels[case_id],
            note=safety_note,
        ),
        s(
            "保存标签",
            "click",
            rid("records_remark_tag_save_tv"),
            note=safety_note,
        ),
        s(
            "校验新增标签显示",
            "assert",
            rid("a_records_remark_tag_tv"),
            assertion="text_contains",
            expected=labels[case_id],
            note=safety_note,
        ),
    ]
    if case_id == "TC-DETAIL-020":
        steps.extend(
            [
                s(
                    "删除本轮唯一标签并确认",
                    "delete_current_run_tag",
                    timeout=15,
                    note=safety_note,
                ),
                s(
                    "校验标签集合已恢复",
                    "assert",
                    assertion="activity_endswith",
                    expected=".RecordsRemarkActivity",
                    note=safety_note,
                ),
            ]
        )
    return steps


def detail_remark_roundtrip_steps(case_id: str) -> list[tuple[Any, ...]]:
    """Create one unique remark and force the original collection to be restored."""

    labels = {
        "TC-DETAIL-012": "测试备注${RUN_TOKEN}",
        "TC-DETAIL-021": "影像备注${RUN_TOKEN}",
        "TC-DETAIL-022": "……&&*……*&${RUN_TOKEN}",
        "TC-DETAIL-035": "咨询单备注${RUN_TOKEN}",
        "TC-DETAIL-037": "自动化删除备注${RUN_TOKEN}",
    }
    if case_id not in labels:
        raise ValueError(f"Unsupported remark roundtrip case: {case_id}")
    label = labels[case_id]
    image_case = case_id in {"TC-DETAIL-012", "TC-DETAIL-021", "TC-DETAIL-022"}
    seed_name = "校验存在历史影像" if image_case else "校验存在咨询单"
    entry_id = (
        "a_records_detail_all_remark_ifv" if image_case else "a_consultation_result_remark_ifv"
    )
    open_name = "进入第一张影像备注" if image_case else "进入第一条咨询单备注"
    safety_note = (
        "提交前记录备注内容与可删除控件集合；断言结束或异常退出时仅删除"
        "本轮唯一新增备注，并验证原备注集合严格恢复"
    )
    steps = [
        *detail_profile_entry_steps()[:8],
        s(
            seed_name,
            "assert",
            rid(entry_id),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
            note=safety_note,
        ),
        s(
            open_name,
            "click",
            rid(entry_id),
            note=safety_note,
        ),
        s(
            "校验备注Activity",
            "assert",
            assertion="activity_endswith",
            expected=".RecordsRemarkActivity",
            timeout=15,
        ),
        s(
            "校验备注输入框",
            "assert",
            rid("records_remark_remark_et"),
            assertion="element_visible",
        ),
        s(
            "输入本轮唯一备注",
            "input",
            rid("records_remark_remark_et"),
            label,
            note=safety_note,
        ),
        s(
            "校验提交按钮启用",
            "assert",
            rid("records_remark_remark_commit_tv"),
            assertion="element_enabled",
        ),
        s(
            "提交本轮备注",
            "click",
            rid("records_remark_remark_commit_tv"),
            note=safety_note,
        ),
        s(
            "校验新增备注显示",
            "assert",
            rid("a_records_remark_list_content_tv"),
            assertion="text_contains",
            expected=label,
            timeout=15,
            note=safety_note,
        ),
    ]
    if case_id == "TC-DETAIL-037":
        steps.extend(
            [
                s(
                    "删除本轮唯一备注并确认",
                    "delete_current_run_remark",
                    timeout=15,
                    note=safety_note,
                ),
                s(
                    "校验备注集合已恢复",
                    "assert",
                    assertion="activity_endswith",
                    expected=".RecordsRemarkActivity",
                    note=safety_note,
                ),
            ]
        )
    return steps


def detail_card_sync_steps(case_id: str) -> list[tuple[Any, ...]]:
    """Change one profile key, assert its list-card value, then restore all fields."""

    contracts = {
        "TC-DETAIL-026": (
            "customer_edit_username_et",
            "自动化${RUN_TOKEN}",
            "a_records_name_tv",
            "姓名",
        ),
        "TC-DETAIL-027": (
            "customer_edit_unique_et",
            "${RUN_PHONE}",
            "a_records_unique_tv",
            "手机号",
        ),
    }
    if case_id not in contracts:
        raise ValueError(f"Unsupported detail card sync case: {case_id}")
    field_id, target_value, card_id, label = contracts[case_id]
    restore_note = (
        "卡片同步断言结束或异常退出后恢复运行前姓名、手机号、性别、生日、"
        "邮箱、婚姻状态、地址和个人备注，并逐字段比对"
    )
    return [
        *detail_profile_entry_steps(),
        s(
            f"输入本轮{label}",
            "input",
            rid(field_id),
            target_value,
            note=restore_note,
        ),
        s(
            f"保存本轮{label}",
            "click",
            rid("customer_edit_save_tv"),
            note=restore_note,
        ),
        s("返回顾客列表", "back", note=restore_note),
        s(
            "校验顾客列表Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerRecordsActivity",
            timeout=15,
            note=restore_note,
        ),
        s(
            f"校验顾客卡片{label}同步",
            "assert",
            rid(card_id),
            assertion="text_equals",
            expected=target_value,
            timeout=15,
            note=restore_note,
        ),
    ]


def detail_profile_remark_steps(case_id: str) -> list[tuple[Any, ...]]:
    """Build guarded profile-remark writes with eight-field restoration."""

    restore_note = "执行后恢复运行前八项资料（含个人备注），并重新打开逐字段比对"
    if case_id == "TC-DETAIL-013":
        mutation_name, action, input_value = "清空个人备注", "clear", None
        assertion, expected = "attribute_equals", "text|"
    else:
        mutation_name, action, input_value = (
            "输入个人备注特殊字符",
            "input",
            "……&&*……*&",
        )
        assertion, expected = "text_equals", "……&&*……*&"
    return [
        *detail_profile_entry_steps(),
        s("滚动到个人备注", "scroll_profile_remark", timeout=15, note=restore_note),
        s(
            mutation_name,
            action,
            rid("customer_edit_remark_et"),
            input_value=input_value,
            note=restore_note,
        ),
        s("保存个人备注", "click", rid("customer_edit_save_tv"), note=restore_note),
        s("重新打开详情编辑", "click", rid("customer_detail_edit_tv"), timeout=15),
        s("重新滚动到个人备注", "scroll_profile_remark", timeout=15),
        s(
            "校验保存后的个人备注",
            "assert",
            rid("customer_edit_remark_et"),
            assertion=assertion,
            expected=expected,
            note=restore_note,
        ),
    ]


def detail_unselected_delete_prompt_steps() -> list[tuple[Any, ...]]:
    """Verify the no-selection delete prompt without selecting or deleting images."""

    prompt = "请选择需要删除的影像"
    return [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "输入影像种子顾客",
            "input",
            rid("customer_records_search_et"),
            "咨询单特定",
        ),
        s("执行顾客搜索", "click", rid("customer_records_search_tv")),
        s(
            "校验影像种子顾客唯一匹配",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_equals",
            expected="1",
            timeout=15,
        ),
        s("打开唯一顾客卡片", "click", rid("a_records_cl"), timeout=15),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s(
            "校验存在历史影像",
            "assert",
            rid("a_records_detail_all_head_ifv"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
        s(
            "校验管理入口初始状态",
            "assert",
            rid("customer_detail_manager_tv"),
            assertion="text_equals",
            expected="管理",
        ),
        s("进入影像管理态", "click", rid("customer_detail_manager_tv")),
        s(
            "校验完成按钮",
            "assert",
            rid("customer_detail_manager_tv"),
            assertion="text_equals",
            expected="完成",
        ),
        s(
            "校验影像选择控件",
            "assert",
            rid("a_records_detail_all_select_ifv"),
            assertion="element_visible",
        ),
        s(
            "校验删除按钮",
            "assert",
            rid("customer_detail_delete_tv"),
            assertion="element_visible",
        ),
        s(
            "未选择影像直接点击删除",
            "click",
            rid("customer_detail_delete_tv"),
            note="不勾选任何影像，仅验证未选择删除提示",
        ),
        s(
            "校验未选择影像提示",
            "assert",
            assertion="page_source_contains",
            expected=prompt,
            timeout=15,
        ),
        s("退出影像管理态", "click", rid("customer_detail_manager_tv")),
        s(
            "校验恢复管理按钮",
            "assert",
            rid("customer_detail_manager_tv"),
            assertion="text_equals",
            expected="管理",
        ),
        s(
            "校验选择控件已隐藏",
            "assert",
            rid("a_records_detail_all_select_ifv"),
            assertion="element_not_visible",
        ),
        s(
            "校验删除按钮已隐藏",
            "assert",
            rid("customer_detail_delete_tv"),
            assertion="element_not_visible",
        ),
    ]


CURRENT_STEPS: dict[str, list[tuple[Any, ...]]] = {
    "TC-LOGIN-001": [
        s("恢复到登录页", "restart_to_login", timeout=20),
        s("输入错误账号", "input", rid("login_username_et"), "${INVALID_USERNAME}"),
        s("输入有效密码", "input", rid("login_pwd_et"), "${TEST_PASSWORD}"),
        s("点击登录", "click", rid("login_tv")),
        s("校验登录失败提示", "assert", rid("cover_prompt_cl"), assertion="element_visible"),
    ],
    "TC-LOGIN-002": [
        s("恢复到登录页", "restart_to_login", timeout=20),
        s("输入有效账号", "input", rid("login_username_et"), "${TEST_USERNAME}"),
        s("输入错误密码", "input", rid("login_pwd_et"), "${INVALID_PASSWORD}"),
        s("点击登录", "click", rid("login_tv")),
        s(
            "校验账号密码错误提示",
            "assert",
            rid("cover_prompt_desc_tv"),
            assertion="text_contains",
            expected="账号密码错误，请重新输入",
        ),
    ],
    "TC-LOGIN-003": [
        s("恢复到登录页", "restart_to_login", timeout=20),
        s("清空账号", "clear", rid("login_username_et")),
        s("清空密码", "clear", rid("login_pwd_et")),
        s("点击登录", "click", rid("login_tv")),
        s(
            "校验空字段未进入首页",
            "assert",
            assertion="activity_endswith",
            expected=".LoginActivity",
            timeout=10,
        ),
    ],
    "TC-LOGIN-004": [
        s("恢复到登录页", "restart_to_login", timeout=20),
        s("输入有效账号", "input", rid("login_username_et"), "${TEST_USERNAME}"),
        s("清空密码", "clear", rid("login_pwd_et")),
        s("点击登录", "click", rid("login_tv")),
        s(
            "校验空密码未进入首页",
            "assert",
            assertion="activity_endswith",
            expected=".LoginActivity",
            timeout=10,
        ),
    ],
    "TC-LOGIN-005": [
        s("恢复到登录页", "restart_to_login", timeout=20),
        s("输入有效账号", "input", rid("login_username_et"), "${TEST_USERNAME}"),
        s("输入有效密码", "input", rid("login_pwd_et"), "${TEST_PASSWORD}"),
        s("点击登录", "click", rid("login_tv")),
        s(
            "校验门店选择器可见",
            "assert",
            rid("login_store_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "校验存在可选门店",
            "assert",
            rid("a_login_join_tv"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
        s(
            "校验未选择门店未进入首页",
            "assert",
            assertion="activity_endswith",
            expected=".LoginActivity",
            timeout=10,
        ),
    ],
    "TC-LOGIN-006": [
        s("恢复到登录页", "restart_to_login", timeout=20),
        s(
            "点击用户协议文字链接",
            "click",
            rid("login_protocol_tv"),
            "0.84,0.5",
        ),
        s(
            "校验进入用户协议Activity",
            "assert",
            assertion="activity_endswith",
            expected=".WebViewActivity",
            timeout=10,
        ),
        s(
            "校验用户协议标题",
            "assert",
            rid("header_title_tv"),
            assertion="text_equals",
            expected="用户协议",
            timeout=15,
        ),
    ],
    "TC-LOGIN-007": [
        s("恢复已认证首页，必要时自动登录", "restart_to_home", timeout=30),
        s(
            "校验进入MainActivity",
            "assert",
            assertion="activity_endswith",
            expected=".MainActivity",
        ),
        s("校验首页根元素", "assert", rid("main_fbl"), assertion="element_visible"),
    ],
    "TC-HOME-001": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("校验顾客档案入口", "assert", rid("main_records_ll"), assertion="element_visible"),
        s("校验美际学院入口", "assert", rid("main_meiji_ll"), assertion="element_visible"),
        s("校验案例管理入口", "assert", rid("main_case_ll"), assertion="element_visible"),
        s("校验设置入口", "assert", rid("main_set_cl"), assertion="element_visible"),
    ],
    "TC-HOME-002": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("输入首页搜索词", "input", rid("main_search_et"), "不存在的用户xyz123", timeout=5),
        s("点击首页搜索", "click", rid("main_search_tv"), timeout=5),
        s(
            "校验搜索结果页",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerRecordsActivity",
            timeout=15,
        ),
        s("校验空结果提示", "assert", rid("empty_tv"), assertion="element_visible"),
    ],
    "TC-HOME-003": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("输入首页搜索词", "input", rid("main_search_et"), "100000000000", timeout=5),
        s("点击首页搜索", "click", rid("main_search_tv"), timeout=5),
        s(
            "校验搜索结果页",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerRecordsActivity",
            timeout=15,
        ),
        s("校验空结果提示", "assert", rid("empty_tv"), assertion="element_visible"),
    ],
    "TC-HOME-004": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("输入首页搜索词", "input", rid("main_search_et"), "t", timeout=5),
        s("点击首页搜索", "click", rid("main_search_tv"), timeout=5),
        s(
            "校验搜索结果页",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerRecordsActivity",
            timeout=15,
        ),
        s(
            "校验搜索到顾客卡片",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-HOME-005": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("输入首页搜索词", "input", rid("main_search_et"), "余婷", timeout=5),
        s("点击首页搜索", "click", rid("main_search_tv"), timeout=5),
        s(
            "校验搜索结果页",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerRecordsActivity",
            timeout=15,
        ),
        s(
            "校验搜索到顾客卡片",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-HOME-006": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("输入手机号片段", "input", rid("main_search_et"), "186", timeout=5),
        s("点击首页搜索", "click", rid("main_search_tv"), timeout=5),
        s(
            "校验搜索结果页",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerRecordsActivity",
            timeout=15,
        ),
        s(
            "校验搜索到顾客卡片",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
        s(
            "校验结果手机号包含搜索片段",
            "assert",
            rid("a_records_unique_tv"),
            assertion="text_contains",
            expected="186",
            timeout=15,
        ),
    ],
    "TC-HOME-007": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s(
            "输入完整手机号变量",
            "input",
            rid("main_search_et"),
            "${SEEDED_CUSTOMER_PHONE}",
            timeout=5,
        ),
        s("点击首页搜索", "click", rid("main_search_tv"), timeout=5),
        s(
            "校验搜索结果页",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerRecordsActivity",
            timeout=15,
        ),
        s(
            "校验仅有一张顾客卡片",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_equals",
            expected="1",
            timeout=15,
        ),
        s(
            "校验结果手机号精确匹配",
            "assert",
            rid("a_records_unique_tv"),
            assertion="text_equals",
            expected="${SEEDED_CUSTOMER_PHONE}",
            timeout=15,
            note="真实手机号仅在本地运行时解析，不写入日志、报告或工作簿",
        ),
    ],
    "TC-HOME-008": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("点击顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerRecordsActivity",
        ),
        s("校验顾客列表", "assert", rid("customer_records_rv"), assertion="element_visible"),
    ],
    "TC-HOME-009": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("点击首页搜索", "click", rid("main_search_tv")),
        s(
            "校验搜索进入顾客列表",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerRecordsActivity",
            timeout=15,
        ),
    ],
    "TC-HOME-010": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("点击案例管理", "click", rid("main_case_ll")),
        s("校验案例Activity", "assert", assertion="activity_endswith", expected=".CaseActivity"),
        s("校验案例列表", "assert", rid("case_rv"), assertion="element_visible"),
    ],
    "TC-HOME-011": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("点击设置", "click", rid("main_set_cl")),
        s("校验设置Activity", "assert", assertion="activity_endswith", expected=".SetActivity"),
        s("校验设置页根元素", "assert", rid("set_logout_tv"), assertion="element_visible"),
    ],
    "TC-HOME-012": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("点击美际学院", "click", rid("main_meiji_ll")),
        s(
            "校验美际学院Activity",
            "assert",
            assertion="activity_endswith",
            expected=".WebViewPCActivity",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-001": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "输入顾客搜索词",
            "input",
            rid("customer_records_search_et"),
            "不存在的用户xyz123",
        ),
        s("点击顾客搜索", "click", rid("customer_records_search_tv")),
        s("校验空结果提示", "assert", rid("empty_tv"), assertion="element_visible"),
    ],
    "TC-CUSTOMER-002": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("输入顾客搜索词", "input", rid("customer_records_search_et"), "t"),
        s("点击顾客搜索", "click", rid("customer_records_search_tv")),
        s(
            "校验搜索到顾客卡片",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-003": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("输入顾客搜索词", "input", rid("customer_records_search_et"), "余婷"),
        s("点击顾客搜索", "click", rid("customer_records_search_tv")),
        s(
            "校验搜索到顾客卡片",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-004": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "输入顾客搜索词",
            "input",
            rid("customer_records_search_et"),
            "100000000000",
        ),
        s("点击顾客搜索", "click", rid("customer_records_search_tv")),
        s("校验空结果提示", "assert", rid("empty_tv"), assertion="element_visible"),
    ],
    "TC-CUSTOMER-005": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验至少一张顾客卡片",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
        ),
        s("校验姓名字段", "assert", rid("a_records_name_tv"), assertion="element_visible"),
        s("校验年龄字段", "assert", rid("a_records_age_tv"), assertion="element_visible"),
        s("校验唯一标识字段", "assert", rid("a_records_unique_tv"), assertion="element_visible"),
        s("校验上次检测时间", "assert", rid("a_records_time_tv"), assertion="element_visible"),
        s("校验检测次数字段", "assert", rid("a_records_count_tv"), assertion="element_visible"),
    ],
    "TC-CUSTOMER-006": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s("点击第一张顾客卡片", "click", rid("a_records_cl"), index=0),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
        ),
        s("校验详情姓名", "assert", rid("customer_detail_name_tv"), assertion="element_visible"),
        s("校验顾客信息", "assert", rid("customer_detail_info_tv"), assertion="element_visible"),
        s("校验编辑按钮", "assert", rid("customer_detail_edit_tv"), assertion="element_visible"),
    ],
    "TC-CUSTOMER-007": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("打开日期范围筛选", "click", rid("customer_records_date_ll")),
        s(
            "校验日期弹层",
            "assert",
            rid("customer_records_date_cl"),
            assertion="element_visible",
        ),
        s(
            "滚动日期范围：2015-02-04~2015-02-13",
            "select_date_range",
            input_value="2015-02-04~2015-02-13",
            timeout=30,
        ),
        s("确认日期范围", "click", rid("customer_records_date_confirm_tv")),
        s(
            "校验日期范围文本",
            "assert",
            rid("customer_records_date_tv"),
            assertion="text_equals",
            expected="2015-02-04~2015-02-13",
            timeout=15,
        ),
        s(
            "校验日期筛选为空",
            "assert",
            rid("empty_tv"),
            assertion="element_visible",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-008": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("打开日期范围筛选", "click", rid("customer_records_date_ll")),
        s(
            "校验日期弹层",
            "assert",
            rid("customer_records_date_cl"),
            assertion="element_visible",
        ),
        s(
            "滚动日期范围：2026-01-01~2026-06-28",
            "select_date_range",
            input_value="2026-01-01~2026-06-28",
            timeout=30,
        ),
        s("确认日期范围", "click", rid("customer_records_date_confirm_tv")),
        s(
            "校验日期范围文本",
            "assert",
            rid("customer_records_date_tv"),
            assertion="text_equals",
            expected="2026-01-01~2026-06-28",
            timeout=15,
        ),
        s(
            "校验日期筛选结果",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-009": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("打开顾客筛选", "click", rid("customer_records_other_tv")),
        s(
            "校验筛选弹层",
            "assert",
            rid("customer_records_other_cl"),
            assertion="element_visible",
        ),
        s("选择筛选项：男", "click", rid("customer_records_other_sex_man_tv")),
        s("确认顾客筛选", "click", rid("customer_records_other_confirm_tv")),
        s(
            "校验筛选结果",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-010": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("打开顾客筛选", "click", rid("customer_records_other_tv")),
        s(
            "校验筛选弹层",
            "assert",
            rid("customer_records_other_cl"),
            assertion="element_visible",
        ),
        s("选择筛选项：女", "click", rid("customer_records_other_sex_female_tv")),
        s("确认顾客筛选", "click", rid("customer_records_other_confirm_tv")),
        s(
            "校验筛选结果",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-012": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("打开顾客筛选", "click", rid("customer_records_other_tv")),
        s(
            "校验筛选弹层",
            "assert",
            rid("customer_records_other_cl"),
            assertion="element_visible",
        ),
        s("选择筛选项：18-25岁", "click", rid("customer_records_other_age_1_tv")),
        s("确认顾客筛选", "click", rid("customer_records_other_confirm_tv")),
        s(
            "校验筛选结果",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-013": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("打开顾客筛选", "click", rid("customer_records_other_tv")),
        s(
            "校验筛选弹层",
            "assert",
            rid("customer_records_other_cl"),
            assertion="element_visible",
        ),
        s("选择筛选项：26-35岁", "click", rid("customer_records_other_age_2_tv")),
        s("确认顾客筛选", "click", rid("customer_records_other_confirm_tv")),
        s(
            "校验筛选结果",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-014": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("打开顾客筛选", "click", rid("customer_records_other_tv")),
        s(
            "校验筛选弹层",
            "assert",
            rid("customer_records_other_cl"),
            assertion="element_visible",
        ),
        s("选择筛选项：36-45岁", "click", rid("customer_records_other_age_3_tv")),
        s("确认顾客筛选", "click", rid("customer_records_other_confirm_tv")),
        s(
            "校验筛选结果",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-015": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("打开顾客筛选", "click", rid("customer_records_other_tv")),
        s(
            "校验筛选弹层",
            "assert",
            rid("customer_records_other_cl"),
            assertion="element_visible",
        ),
        s("选择筛选项：45岁以上", "click", rid("customer_records_other_age_4_tv")),
        s("确认顾客筛选", "click", rid("customer_records_other_confirm_tv")),
        s("校验空结果提示", "assert", rid("empty_tv"), assertion="element_visible"),
    ],
    "TC-CUSTOMER-016": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("打开顾客筛选", "click", rid("customer_records_other_tv")),
        s(
            "校验筛选弹层",
            "assert",
            rid("customer_records_other_cl"),
            assertion="element_visible",
        ),
        s("选择筛选项：18岁以下", "click", rid("customer_records_other_age_0_tv")),
        s("确认顾客筛选", "click", rid("customer_records_other_confirm_tv")),
        s(
            "校验筛选结果",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-017": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("打开日期范围筛选", "click", rid("customer_records_date_ll")),
        s(
            "校验日期弹层",
            "assert",
            rid("customer_records_date_cl"),
            assertion="element_visible",
        ),
        s(
            "滚动日期范围：2015-02-04~2015-02-28",
            "select_date_range",
            input_value="2015-02-04~2015-02-28",
            timeout=30,
        ),
        s("确认日期范围", "click", rid("customer_records_date_confirm_tv")),
        s(
            "校验日期范围文本",
            "assert",
            rid("customer_records_date_tv"),
            assertion="text_equals",
            expected="2015-02-04~2015-02-28",
            timeout=15,
        ),
        s("打开其他筛选", "click", rid("customer_records_other_tv")),
        s(
            "校验其他筛选弹层",
            "assert",
            rid("customer_records_other_cl"),
            assertion="element_visible",
        ),
        s("选择联合筛选项：男", "click", rid("customer_records_other_sex_man_tv")),
        s("选择联合筛选项：18-25岁", "click", rid("customer_records_other_age_1_tv")),
        s("确认联合筛选", "click", rid("customer_records_other_confirm_tv")),
        s(
            "校验日期筛选为空",
            "assert",
            rid("empty_tv"),
            assertion="element_visible",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-018": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("打开日期范围筛选", "click", rid("customer_records_date_ll")),
        s(
            "校验日期弹层",
            "assert",
            rid("customer_records_date_cl"),
            assertion="element_visible",
        ),
        s(
            "滚动日期范围：2026-01-01~2026-07-30",
            "select_date_range",
            input_value="2026-01-01~2026-07-30",
            timeout=30,
        ),
        s("确认日期范围", "click", rid("customer_records_date_confirm_tv")),
        s(
            "校验日期范围文本",
            "assert",
            rid("customer_records_date_tv"),
            assertion="text_equals",
            expected="2026-01-01~2026-07-30",
            timeout=15,
        ),
        s("打开其他筛选", "click", rid("customer_records_other_tv")),
        s(
            "校验其他筛选弹层",
            "assert",
            rid("customer_records_other_cl"),
            assertion="element_visible",
        ),
        s("选择联合筛选项：男", "click", rid("customer_records_other_sex_man_tv")),
        s("选择联合筛选项：18-25岁", "click", rid("customer_records_other_age_1_tv")),
        s("确认联合筛选", "click", rid("customer_records_other_confirm_tv")),
        s(
            "校验日期筛选结果",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-020": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "校验顾客卡片",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
        s(
            "校验头像或性别图像",
            "assert",
            rid("a_records_head_ifv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "校验姓名字段",
            "assert",
            rid("a_records_name_tv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "校验年龄字段",
            "assert",
            rid("a_records_age_tv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "校验手机号字段",
            "assert",
            rid("a_records_unique_tv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "校验上次检测时间",
            "assert",
            rid("a_records_time_tv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "校验检测次数字段",
            "assert",
            rid("a_records_count_tv"),
            assertion="element_visible",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-021": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("点击第一张顾客卡片", "click", rid("a_records_cl"), timeout=15),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s(
            "校验详情页面",
            "assert",
            rid("customer_detail_name_tv"),
            assertion="element_visible",
            timeout=15,
        ),
    ],
    "TC-CUSTOMER-022": [
        s("恢复到首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("打开其他筛选", "click", rid("customer_records_other_tv")),
        s(
            "校验其他筛选面板",
            "assert",
            rid("customer_records_other_cl"),
            assertion="element_visible",
        ),
        s("选择性别：男", "click", rid("customer_records_other_sex_man_tv")),
        s(
            "校验性别已选中",
            "assert",
            rid("customer_records_other_sex_man_tv"),
            assertion="element_selected",
        ),
        s("选择年龄：18-25岁", "click", rid("customer_records_other_age_1_tv")),
        s(
            "校验年龄已选中",
            "assert",
            rid("customer_records_other_age_1_tv"),
            assertion="element_selected",
        ),
        s(
            "输入备注筛选词",
            "input",
            rid("customer_records_other_remark_et"),
            input_value="123",
        ),
        s(
            "校验备注筛选词",
            "assert",
            rid("customer_records_other_remark_et"),
            assertion="attribute_equals",
            expected="text|123",
        ),
        s("重置筛选条件", "click", rid("customer_records_other_reset_tv")),
        s(
            "校验性别已取消",
            "assert",
            rid("customer_records_other_sex_man_tv"),
            assertion="attribute_equals",
            expected="selected|false",
        ),
        s(
            "校验年龄已取消",
            "assert",
            rid("customer_records_other_age_1_tv"),
            assertion="attribute_equals",
            expected="selected|false",
        ),
        s(
            "校验备注恢复占位文本",
            "assert",
            rid("customer_records_other_remark_et"),
            assertion="attribute_equals",
            expected="text|搜索备注记录",
        ),
        s(
            "校验重置后面板仍打开",
            "assert",
            rid("customer_records_other_cl"),
            assertion="element_visible",
        ),
    ],
    "TC-DETAIL-001": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "输入专用顾客查询",
            "input",
            rid("customer_records_search_et"),
            input_value="${YANJIA_MUTATION_CUSTOMER_QUERY}",
        ),
        s("执行专用顾客搜索", "click", rid("customer_records_search_tv")),
        s(
            "校验专用顾客唯一匹配",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_equals",
            expected="1",
            timeout=15,
        ),
        s("打开唯一顾客卡片", "click", rid("a_records_cl"), timeout=15),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s("进入详情编辑", "click", rid("customer_detail_edit_tv")),
        s("校验编辑表单", "assert", rid("customer_edit_v"), assertion="element_visible"),
        s("清空姓名", "clear", rid("customer_edit_username_et")),
        s("点击保存", "click", rid("customer_edit_save_tv")),
        s("校验姓名必填提示", "assert", assertion="page_source_contains", expected="姓名不能为空"),
    ],
    "TC-DETAIL-003": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "输入专用顾客查询",
            "input",
            rid("customer_records_search_et"),
            input_value="${YANJIA_MUTATION_CUSTOMER_QUERY}",
        ),
        s("执行专用顾客搜索", "click", rid("customer_records_search_tv")),
        s(
            "校验专用顾客唯一匹配",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_equals",
            expected="1",
            timeout=15,
        ),
        s("打开唯一顾客卡片", "click", rid("a_records_cl"), timeout=15),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s("进入详情编辑", "click", rid("customer_detail_edit_tv")),
        s("校验编辑表单", "assert", rid("customer_edit_v"), assertion="element_visible"),
        s(
            "输入33字符姓名",
            "input",
            rid("customer_edit_username_et"),
            "DQWDQWDWDWDWDWDWDWDDDD11111112EF1",
            note="仅用于专用顾客边界校验；执行后强制恢复原姓名和手机号",
        ),
        s("点击保存", "click", rid("customer_edit_save_tv")),
        s(
            "校验姓名长度提示",
            "assert",
            assertion="page_source_contains",
            expected="客户名称长度不合法",
            note="真机已验证提示存在；无论断言结果均强制恢复并重新打开编辑页比对",
        ),
    ],
    "TC-DETAIL-004": detail_gender_steps("男"),
    "TC-DETAIL-005": detail_gender_steps("女"),
    "TC-DETAIL-007": detail_birthday_steps(),
    "TC-DETAIL-008": detail_address_steps("TC-DETAIL-008"),
    "TC-DETAIL-009": detail_address_steps("TC-DETAIL-009"),
    "TC-DETAIL-010": detail_address_steps("TC-DETAIL-010"),
    "TC-DETAIL-011": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "输入专用顾客查询",
            "input",
            rid("customer_records_search_et"),
            input_value="${YANJIA_MUTATION_CUSTOMER_QUERY}",
        ),
        s("执行专用顾客搜索", "click", rid("customer_records_search_tv")),
        s(
            "校验专用顾客唯一匹配",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_equals",
            expected="1",
            timeout=15,
        ),
        s("打开唯一顾客卡片", "click", rid("a_records_cl"), timeout=15),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s("进入详情编辑", "click", rid("customer_detail_edit_tv")),
        s("校验编辑表单", "assert", rid("customer_edit_v"), assertion="element_visible"),
        s("清空手机号", "clear", rid("customer_edit_unique_et")),
        s("点击保存", "click", rid("customer_edit_save_tv")),
        s(
            "校验手机号必填提示",
            "assert",
            assertion="page_source_contains",
            expected="手机号不能为空",
        ),
    ],
    "TC-DETAIL-012": detail_remark_roundtrip_steps("TC-DETAIL-012"),
    "TC-DETAIL-013": detail_profile_remark_steps("TC-DETAIL-013"),
    "TC-DETAIL-014": detail_profile_remark_steps("TC-DETAIL-014"),
    "TC-DETAIL-015": detail_marital_steps(),
    "TC-DETAIL-016": detail_tag_roundtrip_steps("TC-DETAIL-016"),
    "TC-DETAIL-017": detail_empty_tag_steps(),
    "TC-DETAIL-018": detail_tag_roundtrip_steps("TC-DETAIL-018"),
    "TC-DETAIL-019": detail_whitespace_tag_steps("TC-DETAIL-019"),
    "TC-DETAIL-020": detail_tag_roundtrip_steps("TC-DETAIL-020"),
    "TC-DETAIL-023": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("点击第一张顾客卡片", "click", rid("a_records_cl"), timeout=15),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s(
            "校验存在历史影像",
            "assert",
            rid("a_records_detail_all_head_ifv"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
        s("进入第一张影像备注", "click", rid("a_records_detail_all_remark_ifv")),
        s(
            "校验备注Activity",
            "assert",
            assertion="activity_endswith",
            expected=".RecordsRemarkActivity",
            timeout=15,
        ),
        s(
            "校验备注编辑器",
            "assert",
            rid("records_remark_remark_et"),
            assertion="element_visible",
        ),
        s("输入单个空格", "input", rid("records_remark_remark_et"), "${SPACE}"),
        s(
            "点击提交并校验备注未新增",
            "assert",
            rid("records_remark_remark_commit_tv"),
            assertion="element_count_unchanged_after_click",
            expected=rid("a_records_remark_list_content_tv"),
            timeout=10,
        ),
        s("关闭备注编辑器", "click", rid("records_remark_close_ifv")),
        s(
            "校验返回顾客详情",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
        ),
    ],
    "TC-DETAIL-024": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s("点击第一张顾客卡片", "click", rid("a_records_cl"), timeout=15),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s(
            "校验存在历史影像",
            "assert",
            rid("a_records_detail_all_head_ifv"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
        s(
            "校验管理入口初始状态",
            "assert",
            rid("customer_detail_manager_tv"),
            assertion="text_equals",
            expected="管理",
        ),
        s("进入影像管理态", "click", rid("customer_detail_manager_tv")),
        s(
            "校验完成按钮",
            "assert",
            rid("customer_detail_manager_tv"),
            assertion="text_equals",
            expected="完成",
        ),
        s(
            "校验影像选择控件",
            "assert",
            rid("a_records_detail_all_select_ifv"),
            assertion="element_visible",
        ),
        s(
            "校验删除按钮",
            "assert",
            rid("customer_detail_delete_tv"),
            assertion="element_visible",
        ),
        s("退出影像管理态", "click", rid("customer_detail_manager_tv")),
        s(
            "校验恢复管理按钮",
            "assert",
            rid("customer_detail_manager_tv"),
            assertion="text_equals",
            expected="管理",
        ),
        s(
            "校验选择控件已隐藏",
            "assert",
            rid("a_records_detail_all_select_ifv"),
            assertion="element_not_visible",
        ),
        s(
            "校验删除按钮已隐藏",
            "assert",
            rid("customer_detail_delete_tv"),
            assertion="element_not_visible",
        ),
    ],
    "TC-DETAIL-029": detail_unselected_delete_prompt_steps(),
    "TC-DETAIL-030": detail_consultation_create_steps(),
    "TC-DETAIL-036": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s(
            "校验顾客列表",
            "assert",
            rid("customer_records_rv"),
            assertion="element_visible",
            timeout=15,
        ),
        s(
            "输入咨询单种子顾客",
            "input",
            rid("customer_records_search_et"),
            "咨询单特定",
        ),
        s("执行顾客搜索", "click", rid("customer_records_search_tv")),
        s(
            "校验咨询单种子顾客存在",
            "assert",
            rid("a_records_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
        s("点击第一张顾客卡片", "click", rid("a_records_cl"), timeout=15),
        s(
            "校验顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s(
            "校验存在咨询单",
            "assert",
            rid("a_consultation_result_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
        s("进入第一条咨询单备注", "click", rid("a_consultation_result_remark_ifv")),
        s(
            "校验备注Activity",
            "assert",
            assertion="activity_endswith",
            expected=".RecordsRemarkActivity",
            timeout=15,
        ),
        s(
            "校验备注编辑器",
            "assert",
            rid("records_remark_remark_et"),
            assertion="element_visible",
        ),
        s("输入单个空格", "input", rid("records_remark_remark_et"), "${SPACE}"),
        s(
            "点击提交并校验备注未新增",
            "assert",
            rid("records_remark_remark_commit_tv"),
            assertion="element_count_unchanged_after_click",
            expected=rid("a_records_remark_list_content_tv"),
            timeout=10,
        ),
        s("关闭备注编辑器", "click", rid("records_remark_close_ifv")),
        s(
            "校验返回顾客详情",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
        ),
    ],
    "TC-IMAGE-001": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入顾客档案", "click", rid("main_records_ll")),
        s("点击第一张顾客卡片", "click", rid("a_records_cl"), index=0),
        s("点击第一张历史影像", "click", rid("a_records_detail_all_head_ifv"), index=0, timeout=15),
        s(
            "校验影像Activity",
            "assert",
            assertion="activity_endswith",
            expected=".SkinResultActivity",
            timeout=20,
        ),
        s(
            "校验影像翻页控件",
            "assert",
            rid("skin_result_vp2"),
            assertion="element_visible",
            timeout=20,
        ),
    ],
    "TC-DETAIL-026": detail_card_sync_steps("TC-DETAIL-026"),
    "TC-DETAIL-027": detail_card_sync_steps("TC-DETAIL-027"),
    "TC-DETAIL-021": detail_remark_roundtrip_steps("TC-DETAIL-021"),
    "TC-DETAIL-022": detail_remark_roundtrip_steps("TC-DETAIL-022"),
    "TC-DETAIL-031": detail_tag_roundtrip_steps("TC-DETAIL-031"),
    "TC-DETAIL-032": detail_whitespace_tag_steps("TC-DETAIL-032"),
    "TC-DETAIL-033": detail_tag_roundtrip_steps("TC-DETAIL-033"),
    "TC-DETAIL-034": detail_whitespace_tag_steps("TC-DETAIL-034"),
    "TC-DETAIL-035": detail_remark_roundtrip_steps("TC-DETAIL-035"),
    "TC-DETAIL-037": detail_remark_roundtrip_steps("TC-DETAIL-037"),
    "TC-IMAGE-030": [
        *controlled_image_entry_steps(),
        s(
            "校验AI综合分析评分区",
            "assert",
            rid("f_skin_result_info_comprehensive_rv"),
            assertion="element_visible",
            timeout=20,
            note="只验证评分区可见，不记录受控顾客的具体评分值",
        ),
        s(
            "校验炎敏评分非空",
            "assert",
            image_score_value("炎敏", "radar"),
            assertion="text_not_empty",
            note="XPath 已在真机验证唯一匹配；不记录评分值",
        ),
        s(
            "校验肤色评分非空",
            "assert",
            image_score_value("肤色", "radar"),
            assertion="text_not_empty",
            note="XPath 已在真机验证唯一匹配；不记录评分值",
        ),
        s(
            "校验色素评分非空",
            "assert",
            image_score_value("色素", "feature"),
            assertion="text_not_empty",
            note="XPath 已在真机验证唯一匹配；不记录评分值",
        ),
        s(
            "校验亮度评分非空",
            "assert",
            image_score_value("亮度", "feature"),
            assertion="text_not_empty",
            note="XPath 已在真机验证唯一匹配；不记录评分值",
        ),
        s(
            "校验红区评分非空",
            "assert",
            image_score_value("红区", "feature"),
            assertion="text_not_empty",
            note="XPath 已在真机验证唯一匹配；不记录评分值",
        ),
    ],
    "TC-IMAGE-032": [
        *controlled_image_entry_steps(),
        s("查看关联案例", "click", rid("f_skin_result_info_view_case_tv")),
        s(
            "校验案例库Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CaseActivity",
            timeout=15,
        ),
        s("校验案例列表", "assert", rid("case_rv"), assertion="element_visible", timeout=15),
    ],
    "TC-IMAGE-033": [
        *controlled_image_entry_steps(),
        s("展开影像操作菜单", "click", rid("f_skin_result_info_menu_ifv")),
        s(
            "校验查看报告入口",
            "assert",
            rid("f_skin_result_info_menu_report_ll"),
            assertion="element_visible",
        ),
        s("查看影像报告", "click", rid("f_skin_result_info_menu_report_ll")),
        s(
            "校验报告Activity",
            "assert",
            assertion="activity_endswith",
            expected=".SkinReportActivity",
            timeout=20,
        ),
        s(
            "校验报告预览",
            "assert",
            rid("skin_report_pdfv"),
            assertion="element_visible",
            timeout=20,
        ),
        s(
            "校验报告分享入口",
            "assert",
            rid("skin_report_share_tv"),
            assertion="element_visible",
        ),
    ],
    "TC-IMAGE-034": [
        *controlled_image_entry_steps(),
        s("返回顾客详情", "click", rid("skin_result_back_ll")),
        s(
            "校验退出确认提示",
            "assert",
            rid("cover_prompt_v1_tv"),
            assertion="text_equals",
            expected="不保存直接退出",
        ),
        s("不保存直接退出", "click", rid("cover_prompt_v1_tv")),
        s(
            "校验返回顾客详情Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CustomerDetailActivity",
            timeout=15,
        ),
        s(
            "校验顾客详情已恢复",
            "assert",
            rid("customer_detail_name_tv"),
            assertion="element_visible",
        ),
    ],
    "TC-CASE-001": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入案例管理", "click", rid("main_case_ll")),
        s("校验搜索框", "assert", rid("case_search_et"), assertion="element_visible"),
        s(
            "校验至少一个标签",
            "assert",
            rid("a_case_tag_tv"),
            assertion="element_count_gte",
            expected="1",
        ),
        s(
            "校验至少一个类别",
            "assert",
            rid("a_case_category_tv"),
            assertion="element_count_gte",
            expected="1",
        ),
        s("校验案例列表", "assert", rid("case_rv"), assertion="element_visible"),
    ],
    "TC-CASE-002": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入案例库", "click", rid("main_case_ll")),
        s(
            "校验案例库Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CaseActivity",
            timeout=15,
        ),
        s("校验案例列表", "assert", rid("case_rv"), assertion="element_visible", timeout=15),
        s(
            "输入不存在的案例标签",
            "input",
            rid("case_search_et"),
            "${NON_EXISTENT_CASE_TAG}",
        ),
        s("执行案例搜索", "click", rid("case_search_tv")),
        s("校验空结果提示", "assert", rid("empty_tv"), assertion="element_visible", timeout=15),
        s(
            "校验不存在案例卡片",
            "assert",
            rid("a_case_image_cl"),
            assertion="element_not_visible",
            timeout=15,
        ),
    ],
    "TC-CASE-003": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入案例库", "click", rid("main_case_ll")),
        s(
            "校验案例库Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CaseActivity",
            timeout=15,
        ),
        s("校验案例列表", "assert", rid("case_rv"), assertion="element_visible", timeout=15),
        s("输入案例标签", "input", rid("case_search_et"), "火"),
        s("执行案例搜索", "click", rid("case_search_tv")),
        s(
            "校验存在案例卡片",
            "assert",
            rid("a_case_image_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
        s(
            "校验案例卡片包含标签火",
            "assert",
            case_card_tag("火"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-CASE-004": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入案例库", "click", rid("main_case_ll")),
        s(
            "校验案例库Activity",
            "assert",
            assertion="activity_endswith",
            expected=".CaseActivity",
            timeout=15,
        ),
        s("校验案例列表", "assert", rid("case_rv"), assertion="element_visible", timeout=15),
        s("展开案例标签", "click", rid("case_tag_expand_tv")),
        s("筛选案例标签火", "click", exact_case_tag("火")),
        s(
            "校验存在案例卡片",
            "assert",
            rid("a_case_image_cl"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
        s(
            "校验案例卡片包含标签火",
            "assert",
            case_card_tag("火"),
            assertion="element_count_gte",
            expected="1",
            timeout=15,
        ),
    ],
    "TC-PROFILE-001": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入设置", "click", rid("main_set_cl")),
        s("进入个人中心", "click", rid("set_personal_ll")),
        s(
            "校验个人中心Activity",
            "assert",
            assertion="activity_endswith",
            expected=".PersonalActivity",
        ),
        s("校验账号非空", "assert", rid("personal_account_tv"), assertion="text_not_empty"),
    ],
    "TC-PROFILE-002": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入设置", "click", rid("main_set_cl")),
        s("进入个人中心", "click", rid("set_personal_ll")),
        s("校验姓名非空", "assert", rid("personal_name_tv"), assertion="text_not_empty"),
    ],
    "TC-PROFILE-003": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入设置", "click", rid("main_set_cl")),
        s("进入个人中心", "click", rid("set_personal_ll")),
        s(
            "校验注册时间非空",
            "assert",
            rid("personal_register_time_tv"),
            assertion="text_not_empty",
        ),
    ],
    "TC-PROFILE-004": [
        s("恢复首页", "restart_to_home", timeout=30),
        s("进入设置", "click", rid("main_set_cl")),
        s("进入个人中心", "click", rid("set_personal_ll")),
        s("校验手机号字段存在", "assert", rid("personal_phone_tv"), assertion="element_visible"),
    ],
}


KEYWORDS = [
    ("操作", "restart_to_login", "重启应用并恢复登录页", "否", "无", ""),
    ("操作", "restart_to_home", "重启应用并恢复首页，必要时自动登录", "否", "无", ""),
    ("操作", "ensure_home", "从当前页面返回首页，必要时自动登录", "否", "无", ""),
    (
        "操作",
        "click",
        "等待并点击指定元素；可按元素内部相对位置点击局部链接",
        "是",
        "可选：相对X,相对Y（0到1）",
        "0.84,0.5",
    ),
    ("操作", "long_click", "长按指定元素", "是", "输入为秒数，默认1", "1.5"),
    (
        "操作",
        "input",
        "清空后输入；步骤为空时使用主表输入数据",
        "是",
        "文本或${变量名}",
        "${SEEDED_CASE_TAG}",
    ),
    ("操作", "clear", "清空输入框", "是", "无", ""),
    (
        "操作",
        "delete_current_run_tag",
        "仅删除受控恢复会话中本轮唯一新增标签并确认集合恢复",
        "否",
        "无",
        "",
    ),
    (
        "操作",
        "delete_current_run_remark",
        "仅删除受控恢复会话中本轮唯一新增备注并确认集合恢复",
        "否",
        "无",
        "",
    ),
    (
        "操作",
        "open_first_unlinked_image",
        "按检测时间打开第一张尚未关联咨询单的历史影像；若无候选则不写入并失败",
        "否",
        "无",
        "",
    ),
    (
        "操作",
        "verify_selected_image_consultation",
        "校验本轮所选影像的检测时间已出现在咨询单卡片中",
        "否",
        "无",
        "",
    ),
    ("操作", "back", "Android返回键", "否", "无", ""),
    ("操作", "hide_keyboard", "隐藏软键盘", "否", "无", ""),
    ("操作", "press_keycode", "发送Android按键码", "否", "按键码", "66"),
    ("操作", "wait_visible", "等待元素可见", "是", "无", ""),
    ("操作", "wait_invisible", "等待元素不可见", "是", "无", ""),
    ("操作", "swipe", "按坐标或屏幕比例滑动", "否", "x1,y1,x2,y2[,毫秒]", "0.8,0.5,0.2,0.5,500"),
    ("操作", "scroll_to_text", "滚动到指定文本", "否", "目标文本", "用户协议"),
    (
        "操作",
        "scroll_profile_remark",
        "滚动顾客资料表单并显示个人备注输入框",
        "否",
        "无",
        "",
    ),
    ("操作", "activate_app", "激活颜佳AI", "否", "无", ""),
    ("操作", "select_first_store", "选择登录页首个可用门店", "否", "无", ""),
    ("操作", "terminate_app", "终止颜佳AI", "否", "无", ""),
    ("操作", "background_app", "应用进入后台后恢复", "否", "后台秒数", "3"),
    ("操作", "set_orientation", "设置横竖屏", "否", "LANDSCAPE或PORTRAIT", "LANDSCAPE"),
    (
        "操作",
        "select_birthday",
        "按真实 Android 单组日期滚轮选择生日",
        "否",
        "YYYY-MM-DD",
        "2022-08-04",
    ),
    (
        "操作",
        "select_date_range",
        "按真实 Android 日期滚轮选择日期范围",
        "否",
        "YYYY/MM/DD-YYYY/MM/DD",
        "2026/01/01-2026/06/28",
    ),
    ("操作", "pause", "固定暂停，限制0到30秒，优先使用显式等待", "否", "秒数", "1"),
    ("操作", "assert", "执行断言类型列指定的断言", "按断言", "期望值按断言填写", ""),
    ("操作", "verify", "assert的同义关键字", "按断言", "期望值按断言填写", ""),
    ("操作", "noop", "空操作，用于仅记录说明的步骤", "否", "无", ""),
    ("断言", "element_visible", "指定索引元素可见", "是", "无", ""),
    ("断言", "element_exists", "指定索引元素存在", "是", "无", ""),
    ("断言", "element_not_visible", "没有可见匹配元素", "是", "无", ""),
    ("断言", "all_elements_visible", "所有匹配元素均可见且至少一个", "是", "无", ""),
    ("断言", "element_count_gte", "元素数量大于等于期望", "是", "整数", "1"),
    ("断言", "element_count_equals", "元素数量等于期望", "是", "整数", "2"),
    ("断言", "element_disabled", "元素可见且未启用", "是", "无", ""),
    (
        "断言",
        "element_count_unchanged_after_click",
        "对定位控件执行中心点击，并在超时窗口内持续校验期望定位器的元素数量不变且Activity不变",
        "是",
        "待观察元素的定位器",
        "id=com.example:id/record_content",
    ),
    ("断言", "element_enabled", "元素可见且已启用", "是", "无", ""),
    ("断言", "element_selected", "元素处于选中状态", "是", "无", ""),
    ("断言", "text_not_empty", "元素文本非空", "是", "无", ""),
    ("断言", "text_equals", "元素文本等于期望", "是", "文本或${变量名}", "男"),
    ("断言", "text_contains", "元素文本包含期望", "是", "文本或${变量名}", "${SEEDED_CASE_TAG}"),
    ("断言", "activity_endswith", "当前Activity以后缀结尾", "否", "Activity后缀", ".MainActivity"),
    (
        "断言",
        "activity_not_endswith",
        "当前Activity不应以后缀结尾",
        "否",
        "Activity后缀",
        ".LoginActivity",
    ),
    ("断言", "activity_equals", "当前Activity等于期望", "否", "完整Activity", ""),
    ("断言", "package_equals", "当前包名等于期望", "否", "包名", "com.xiaofutech.yanjia_ai"),
    ("断言", "orientation_equals", "屏幕方向等于期望", "否", "LANDSCAPE或PORTRAIT", "LANDSCAPE"),
    ("断言", "attribute_equals", "元素属性等于期望", "是", "属性名|期望值", "checked|true"),
    (
        "断言",
        "attribute_contains",
        "元素属性包含期望",
        "是",
        "属性名|期望值",
        "contentDescription|设置",
    ),
    ("断言", "page_source_contains", "页面结构包含文本", "否", "文本或${变量名}", "暂无内容"),
    ("定位", "id", "Android resource-id", "-", "id=完整resource-id", rid("main_fbl")),
    ("定位", "accessibility_id", "无障碍标识", "-", "accessibility_id=值", "accessibility_id=返回"),
    (
        "定位",
        "uiautomator",
        "Android UiSelector",
        "-",
        "uiautomator=表达式",
        'uiautomator=new UiSelector().text("确定")',
    ),
    ("定位", "xpath", "XPath，最后选择", "-", "xpath=表达式", "xpath=//*[@text='确定']"),
    (
        "定位",
        "候选定位器",
        "从左到右依次尝试，首个满足条件的定位器生效",
        "-",
        "多个定位器用 || 分隔",
        f"{rid('main_fbl')} || accessibility_id=首页",
    ),
    (
        "变量",
        "${变量名}",
        "从系统环境或.env解析；账号密码兼容env.txt",
        "-",
        "未配置会明确失败",
        "${YANJIA_USERNAME}",
    ),
]


def _validate_current_step_catalog() -> None:
    """Fail early when step definitions drift from the shared Android catalog."""

    actual = frozenset(CURRENT_STEPS)
    if actual != ANDROID_STEP_SHEET_CASE_IDS:
        missing = sorted(ANDROID_STEP_SHEET_CASE_IDS - actual)
        extra = sorted(actual - ANDROID_STEP_SHEET_CASE_IDS)
        raise RuntimeError(
            f"prepare_excel.py 与 Android 用例目录不一致；目录缺少={missing}，脚本多出={extra}"
        )


def main() -> None:
    _validate_current_step_catalog()
    parser = argparse.ArgumentParser(description="Prepare the Excel-driven test workbook.")
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument(
        "--replace-current-steps",
        action="store_true",
        help="Replace steps only for the currently implemented case IDs.",
    )
    args = parser.parse_args()
    workbook_path = args.workbook.resolve()
    if not workbook_path.is_file():
        raise FileNotFoundError(workbook_path)

    backup_dir = ROOT / "reports" / "excel-backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup = backup_dir / f"{workbook_path.stem}_schema_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(workbook_path, backup)

    workbook = load_workbook(workbook_path)
    try:
        cases = workbook[CASE_SHEET]
        case_headers = ensure_columns(cases, CASE_RESULT_COLUMNS)
        case_rows = case_row_map(cases, case_headers)
        for case_id, row in case_rows.items():
            cases.cell(row, case_headers["是否执行"], "是" if case_id in CURRENT_STEPS else "否")
        style_case_sheet(cases, case_headers)

        prepare_keyword_sheet(workbook)
        prepare_step_sheet(workbook, replace=args.replace_current_steps)
        workbook.save(workbook_path)
    finally:
        workbook.close()
    print(f"Prepared: {workbook_path}")
    print(f"Backup:   {backup}")
    print(f"Enabled cases: {len(CURRENT_STEPS)}")


def prepare_step_sheet(workbook, *, replace: bool) -> Worksheet:
    if STEP_SHEET not in workbook.sheetnames:
        sheet = workbook.create_sheet(STEP_SHEET, 1)
        sheet.append(STEP_HEADERS)
    else:
        sheet = workbook[STEP_SHEET]
        existing_headers = header_map(sheet)
        if tuple(existing_headers) != STEP_HEADERS:
            for name in STEP_HEADERS:
                if name not in existing_headers:
                    sheet.cell(1, sheet.max_column + 1, name)

    headers = header_map(sheet)
    existing_cases = {
        str(sheet.cell(row, headers["用例ID"]).value or "").strip()
        for row in range(2, sheet.max_row + 1)
    }
    if replace:
        rows_to_delete = [
            row
            for row in range(2, sheet.max_row + 1)
            if str(sheet.cell(row, headers["用例ID"]).value or "").strip() in CURRENT_STEPS
        ]
        for row in reversed(rows_to_delete):
            sheet.delete_rows(row)
        existing_cases -= set(CURRENT_STEPS)

    for case_id, steps in CURRENT_STEPS.items():
        if case_id in existing_cases:
            continue
        for order, values in enumerate(steps, 1):
            row_values = (case_id, order, *values)
            sheet.append(row_values)

    style_table(sheet, widths=(20, 10, 34, 22, 62, 32, 26, 34, 12, 10, 12, 12, 40))
    sheet.auto_filter.ref = sheet.dimensions
    sheet.freeze_panes = "A2"
    sheet.data_validations.dataValidation.clear()
    sheet.conditional_formatting = ConditionalFormattingList()
    yes_no = DataValidation(type="list", formula1='"是,否"', allow_blank=False)
    sheet.add_data_validation(yes_no)
    yes_no.add(f"K2:L{max(sheet.max_row, 1000)}")
    actions = DataValidation(
        type="list",
        formula1="=ExcelActionKeywords",
        allow_blank=False,
        error="请选择关键字说明表中的操作类型",
        errorTitle="无效操作类型",
    )
    assertions = DataValidation(
        type="list",
        formula1="=ExcelAssertionKeywords",
        allow_blank=True,
        error="请选择关键字说明表中的断言类型",
        errorTitle="无效断言类型",
    )
    sheet.add_data_validation(actions)
    sheet.add_data_validation(assertions)
    actions.add(f"D2:D{max(sheet.max_row, 1000)}")
    assertions.add(f"G2:G{max(sheet.max_row, 1000)}")
    sheet.conditional_formatting.add(
        f"A2:M{sheet.max_row}",
        FormulaRule(formula=['$L2="否"'], fill=PatternFill("solid", fgColor="E7E6E6")),
    )
    return sheet


def prepare_keyword_sheet(workbook) -> Worksheet:
    if KEYWORD_SHEET in workbook.sheetnames:
        del workbook[KEYWORD_SHEET]
    sheet = workbook.create_sheet(KEYWORD_SHEET)
    sheet.append(("类型", "关键字", "用途", "需要定位器", "输入数据/期望值格式", "示例"))
    for row in KEYWORDS:
        sheet.append(row)
    style_table(sheet, widths=(12, 26, 48, 14, 38, 70))
    sheet.auto_filter.ref = sheet.dimensions
    sheet.freeze_panes = "A2"
    action_rows = [index for index, row in enumerate(KEYWORDS, 2) if row[0] == "操作"]
    assertion_rows = [index for index, row in enumerate(KEYWORDS, 2) if row[0] == "断言"]
    named_ranges = {
        "ExcelActionKeywords": (
            f"'{KEYWORD_SHEET}'!$B$" + str(min(action_rows)) + ":$B$" + str(max(action_rows))
        ),
        "ExcelAssertionKeywords": (
            f"'{KEYWORD_SHEET}'!$B$" + str(min(assertion_rows)) + ":$B$" + str(max(assertion_rows))
        ),
    }
    for name, reference in named_ranges.items():
        if name in workbook.defined_names:
            del workbook.defined_names[name]
        workbook.defined_names.add(DefinedName(name, attr_text=reference))
    return sheet


def ensure_columns(sheet: Worksheet, names: tuple[str, ...]) -> dict[str, int]:
    headers = header_map(sheet)
    source = sheet.cell(1, headers.get("实际结果", 1))
    for name in names:
        if name in headers:
            continue
        target = sheet.cell(1, sheet.max_column + 1, name)
        copy_style(source, target)
        headers[name] = target.column
    return headers


def case_row_map(sheet: Worksheet, headers: dict[str, int]) -> dict[str, int]:
    rows: dict[str, int] = {}
    for row in range(2, sheet.max_row + 1):
        case_id = str(sheet.cell(row, headers["用例ID"]).value or "").strip()
        if case_id:
            rows[case_id] = row
    return rows


def style_case_sheet(sheet: Worksheet, headers: dict[str, int]) -> None:
    widths = {
        "是否执行": 12,
        "最后执行时间": 22,
        "执行耗时(秒)": 16,
        "错误信息": 60,
        "运行编号": 22,
        "Allure报告目录": 60,
    }
    for name, width in widths.items():
        column = get_column_letter(headers[name])
        sheet.column_dimensions[column].width = width
    sheet.data_validations.dataValidation.clear()
    sheet.conditional_formatting = ConditionalFormattingList()
    yes_no = DataValidation(type="list", formula1='"是,否"', allow_blank=False)
    sheet.add_data_validation(yes_no)
    column = get_column_letter(headers["是否执行"])
    yes_no.add(f"{column}2:{column}{max(sheet.max_row, 1000)}")
    result_column = headers.get("实际结果")
    if result_column:
        result_letter = get_column_letter(result_column)
        result_range = f"{result_letter}2:{result_letter}{max(sheet.max_row, 1000)}"
        result_colors = {
            "PASS": "C6EFCE",
            "FAIL": "FFC7CE",
            "ERROR": "FFC7CE",
            "SKIP": "FFEB9C",
            "NOT_RUN": "E7E6E6",
        }
        for status, color in result_colors.items():
            sheet.conditional_formatting.add(
                result_range,
                FormulaRule(
                    formula=["$" + result_letter + f'2="{status}"'],
                    fill=PatternFill("solid", fgColor=color),
                ),
            )
    sheet.auto_filter.ref = sheet.dimensions
    sheet.freeze_panes = "A2"


def style_table(sheet: Worksheet, *, widths: tuple[int, ...]) -> None:
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def header_map(sheet: Worksheet) -> dict[str, int]:
    headers: dict[str, int] = {}
    for cell in sheet[1]:
        if isinstance(cell.value, str) and isinstance(cell.column, int):
            headers[cell.value.strip()] = cell.column
    return headers


def copy_style(source, target) -> None:
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.alignment = copy(source.alignment)
    target.number_format = source.number_format
    target.protection = copy(source.protection)


if __name__ == "__main__":
    main()
