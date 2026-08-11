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


CURRENT_STEPS: dict[str, list[tuple[Any, ...]]] = {
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
    ("操作", "restart_to_home", "重启应用并恢复首页，必要时自动登录", "否", "无", ""),
    ("操作", "ensure_home", "从当前页面返回首页，必要时自动登录", "否", "无", ""),
    ("操作", "click", "等待并点击指定元素", "是", "无", rid("main_records_ll")),
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
    ("操作", "back", "Android返回键", "否", "无", ""),
    ("操作", "hide_keyboard", "隐藏软键盘", "否", "无", ""),
    ("操作", "press_keycode", "发送Android按键码", "否", "按键码", "66"),
    ("操作", "wait_visible", "等待元素可见", "是", "无", ""),
    ("操作", "wait_invisible", "等待元素不可见", "是", "无", ""),
    ("操作", "swipe", "按坐标或屏幕比例滑动", "否", "x1,y1,x2,y2[,毫秒]", "0.8,0.5,0.2,0.5,500"),
    ("操作", "scroll_to_text", "滚动到指定文本", "否", "目标文本", "用户协议"),
    ("操作", "activate_app", "激活颜佳AI", "否", "无", ""),
    ("操作", "terminate_app", "终止颜佳AI", "否", "无", ""),
    ("操作", "background_app", "应用进入后台后恢复", "否", "后台秒数", "3"),
    ("操作", "set_orientation", "设置横竖屏", "否", "LANDSCAPE或PORTRAIT", "LANDSCAPE"),
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
    ("断言", "element_enabled", "元素可见且已启用", "是", "无", ""),
    ("断言", "element_selected", "元素处于选中状态", "是", "无", ""),
    ("断言", "text_not_empty", "元素文本非空", "是", "无", ""),
    ("断言", "text_equals", "元素文本等于期望", "是", "文本或${变量名}", "男"),
    ("断言", "text_contains", "元素文本包含期望", "是", "文本或${变量名}", "${SEEDED_CASE_TAG}"),
    ("断言", "activity_endswith", "当前Activity以后缀结尾", "否", "Activity后缀", ".MainActivity"),
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


def main() -> None:
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
            f"'{KEYWORD_SHEET}'!$B$"
            + str(min(assertion_rows))
            + ":$B$"
            + str(max(assertion_rows))
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
