from __future__ import annotations

import re
from datetime import date
from fnmatch import fnmatchcase
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from yanjia_automation.excel.android_catalog import (
    SINGLE_SHEET_SOURCE,
    android_case_ids,
    get_android_case_definition,
)
from yanjia_automation.excel.models import ExcelCase, ExcelStep
from yanjia_automation.excel.safety import (
    CUSTOMER_MUTATION_CASE_IDS,
    REMARK_MUTATION_CASE_IDS,
    TAG_MUTATION_CASE_IDS,
)
from yanjia_automation.screens.base import PACKAGE

SINGLE_SHEET_REQUIRED_HEADERS = {
    "用例ID",
    "模块",
    "测试场景",
    "测试点",
    "优先级",
    "前置条件",
    "输入数据",
    "期望结果",
    "超时(秒)",
    "是否执行",
}

ANDROID_LOGIN_CASES = android_case_ids(
    source=SINGLE_SHEET_SOURCE,
    categories={"login"},
)
ANDROID_HOME_CASES = android_case_ids(
    source=SINGLE_SHEET_SOURCE,
    categories={"home", "case-library"},
)
ANDROID_CUSTOMER_CASES = android_case_ids(
    source=SINGLE_SHEET_SOURCE,
    categories={"customer"},
)
ANDROID_DETAIL_CASES = android_case_ids(
    source=SINGLE_SHEET_SOURCE,
    categories={"detail"},
)
ANDROID_HOME_SEARCH_CASES = android_case_ids(
    source=SINGLE_SHEET_SOURCE,
    flows={"home_search"},
)
ANDROID_CASE_SEARCH_CASES = android_case_ids(
    source=SINGLE_SHEET_SOURCE,
    flows={"case_search"},
)
ANDROID_CASE_FILTER_CASES = android_case_ids(
    source=SINGLE_SHEET_SOURCE,
    flows={"case_filter"},
)
ANDROID_CUSTOMER_SEARCH_CASES = android_case_ids(
    source=SINGLE_SHEET_SOURCE,
    flows={"customer_search"},
)
ANDROID_CUSTOMER_FILTER_CASES = android_case_ids(
    source=SINGLE_SHEET_SOURCE,
    flows={"customer_filter"},
)
ANDROID_CUSTOMER_FILTER_RESET_CASES = android_case_ids(
    source=SINGLE_SHEET_SOURCE,
    flows={"customer_filter_reset"},
)
ANDROID_CUSTOMER_DATE_FILTER_CASES = android_case_ids(
    source=SINGLE_SHEET_SOURCE,
    flows={"customer_date_filter"},
)

ANDROID_CUSTOMER_DETAIL_CASES = ANDROID_CUSTOMER_CASES | ANDROID_DETAIL_CASES
ANDROID_SINGLE_SHEET_CASES = (
    ANDROID_LOGIN_CASES | ANDROID_HOME_CASES | ANDROID_CUSTOMER_DETAIL_CASES
)

HOME_DESTINATIONS = {
    "TC-HOME-008": ("main_records_ll", ".CustomerRecordsActivity", "打开顾客档案"),
    "TC-HOME-009": ("main_search_tv", ".CustomerRecordsActivity", "打开搜索"),
    "TC-HOME-010": ("main_case_ll", ".CaseActivity", "打开案例库"),
    "TC-HOME-011": ("main_set_cl", ".SetActivity", "打开设置"),
    "TC-HOME-012": ("main_meiji_ll", ".WebViewPCActivity", "打开美际学院"),
    "TC-CASE-001": ("main_case_ll", ".CaseActivity", "打开案例库"),
}

ANDROID_ERROR_EXPECTATIONS = {
    "TC-LOGIN-002": "账号密码错误，请重新输入",
}


class SingleSheetMigrationError(ValueError):
    """Raised when a single-sheet Web case cannot be safely mapped to Android."""


def load_android_single_sheet_cases(
    sheet: Worksheet,
    headers: dict[str, int],
    *,
    case_patterns: tuple[str, ...],
) -> list[ExcelCase]:
    if not case_patterns:
        raise SingleSheetMigrationError(
            "检测到单表 Web/Playwright 用例格式；请用 --excel-case-id 选择已迁移的 Android 用例。"
        )

    cases: list[ExcelCase] = []
    seen_ids: set[str] = set()
    for row in range(2, sheet.max_row + 1):
        case_id = _text(_value(sheet, row, headers, "用例ID"))
        if not case_id or not any(fnmatchcase(case_id, pattern) for pattern in case_patterns):
            continue
        if case_id in seen_ids:
            raise SingleSheetMigrationError(f"主表第{row}行用例ID重复：{case_id}")
        seen_ids.add(case_id)

        if not _boolean(_value(sheet, row, headers, "是否执行"), default=True):
            continue
        if case_id not in ANDROID_SINGLE_SHEET_CASES:
            raise SingleSheetMigrationError(
                f"{case_id} 尚未迁移为 Android Appium 用例；"
                "当前表中的定位器是 Web/Playwright 语法。"
            )

        cases.append(_build_single_sheet_case(sheet, headers, row, case_id))

    return cases


def _build_single_sheet_case(
    sheet: Worksheet,
    headers: dict[str, int],
    row: int,
    case_id: str,
) -> ExcelCase:
    input_data = _optional_text(_value(sheet, row, headers, "输入数据"))
    expected = _optional_text(_value(sheet, row, headers, "期望结果"))
    timeout = _number(_value(sheet, row, headers, "超时(秒)"), default=10.0)
    definition = get_android_case_definition(case_id)
    if definition is None or not definition.supports(SINGLE_SHEET_SOURCE):
        raise SingleSheetMigrationError(f"不支持的 Android 单表用例：{case_id}")
    if case_id in ANDROID_LOGIN_CASES:
        steps = _login_steps(
            row=row,
            case_id=case_id,
            input_data=input_data,
            expected=expected,
            timeout=timeout,
        )
    elif case_id in ANDROID_HOME_CASES:
        steps = _home_steps(
            row=row,
            case_id=case_id,
            input_data=input_data,
            timeout=timeout,
        )
    elif case_id in ANDROID_CUSTOMER_CASES:
        steps = _customer_steps(
            row=row,
            case_id=case_id,
            input_data=input_data,
            timeout=timeout,
        )
    elif case_id in ANDROID_DETAIL_CASES:
        steps = _detail_validation_steps(
            row=row,
            case_id=case_id,
            input_data=input_data,
            timeout=timeout,
        )
    else:
        raise SingleSheetMigrationError(f"不支持的 Android 单表用例：{case_id}")
    return ExcelCase(
        source_row=row,
        case_id=case_id,
        module=_text(_value(sheet, row, headers, "模块")),
        scenario=_text(_value(sheet, row, headers, "测试场景")),
        test_point=_text(_value(sheet, row, headers, "测试点")),
        priority=_text(_value(sheet, row, headers, "优先级")).upper(),
        precondition=_optional_text(_value(sheet, row, headers, "前置条件")),
        input_data=input_data,
        expected_result=expected,
        timeout=timeout,
        tags=definition.tags,
        enabled=True,
        automation_status=definition.automation_status_for(SINGLE_SHEET_SOURCE),
        steps=tuple(steps),
    )


def _login_steps(
    *,
    row: int,
    case_id: str,
    input_data: str | None,
    expected: str | None,
    timeout: float,
) -> list[ExcelStep]:
    steps = [
        _step(
            row=row,
            case_id=case_id,
            order=1,
            name="恢复到登录页",
            action="restart_to_login",
            timeout=max(timeout, 20.0),
        )
    ]

    if case_id in {"TC-LOGIN-001", "TC-LOGIN-002", "TC-LOGIN-007"}:
        username, password = _input_pair(input_data)
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=2,
                    name="输入账号",
                    action="input",
                    locator=_id_locator("login_username_et"),
                    input_value=username,
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=3,
                    name="输入密码",
                    action="input",
                    locator=_id_locator("login_pwd_et"),
                    input_value=password,
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=4,
                    name="点击登录",
                    action="click",
                    locator=_id_locator("login_tv"),
                    timeout=timeout,
                ),
            ]
        )

        if case_id == "TC-LOGIN-007":
            steps.extend(
                [
                    _step(
                        row=row,
                        case_id=case_id,
                        order=5,
                        name="选择首个门店",
                        action="select_first_store",
                        timeout=max(timeout, 20.0),
                    ),
                    _step(
                        row=row,
                        case_id=case_id,
                        order=6,
                        name="校验登录进入首页",
                        action="assert",
                        assertion="activity_endswith",
                        expected=".MainActivity",
                        timeout=max(timeout, 20.0),
                    ),
                ]
            )
        else:
            error_expected = ANDROID_ERROR_EXPECTATIONS.get(case_id)
            steps.append(
                _step(
                    row=row,
                    case_id=case_id,
                    order=5,
                    name="校验登录失败提示",
                    action="assert",
                    locator=(
                        _id_locator("cover_prompt_desc_tv")
                        if error_expected
                        else _id_locator("cover_prompt_cl")
                    ),
                    assertion="text_contains" if error_expected else "element_visible",
                    expected=error_expected,
                    timeout=max(timeout, 10.0),
                )
            )
        return steps

    if case_id == "TC-LOGIN-003":
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=2,
                    name="清空账号",
                    action="clear",
                    locator=_id_locator("login_username_et"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=3,
                    name="清空密码",
                    action="clear",
                    locator=_id_locator("login_pwd_et"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=4,
                    name="点击登录",
                    action="click",
                    locator=_id_locator("login_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=5,
                    name="校验空字段未进入首页",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".LoginActivity",
                    timeout=max(timeout, 5.0),
                ),
            ]
        )
        return steps

    if case_id == "TC-LOGIN-004":
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=2,
                    name="输入有效账号",
                    action="input",
                    locator=_id_locator("login_username_et"),
                    input_value="${TEST_USERNAME}",
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=3,
                    name="清空密码",
                    action="clear",
                    locator=_id_locator("login_pwd_et"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=4,
                    name="点击登录",
                    action="click",
                    locator=_id_locator("login_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=5,
                    name="校验空密码未进入首页",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".LoginActivity",
                    timeout=max(timeout, 5.0),
                ),
            ]
        )
        return steps

    if case_id == "TC-LOGIN-005":
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=2,
                    name="输入有效账号",
                    action="input",
                    locator=_id_locator("login_username_et"),
                    input_value="${TEST_USERNAME}",
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=3,
                    name="输入有效密码",
                    action="input",
                    locator=_id_locator("login_pwd_et"),
                    input_value="${TEST_PASSWORD}",
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=4,
                    name="点击登录",
                    action="click",
                    locator=_id_locator("login_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=5,
                    name="校验门店选择器可见",
                    action="assert",
                    locator=_id_locator("login_store_rv"),
                    assertion="element_visible",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=6,
                    name="校验存在可选门店",
                    action="assert",
                    locator=_id_locator("a_login_join_tv"),
                    assertion="element_count_gte",
                    expected="1",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=7,
                    name="校验未选择门店未进入首页",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".LoginActivity",
                    timeout=max(timeout, 5.0),
                ),
            ]
        )
        return steps

    if case_id == "TC-LOGIN-006":
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=2,
                    name="点击用户协议文字链接",
                    action="click",
                    locator=_id_locator("login_protocol_tv"),
                    input_value="0.84,0.5",
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=3,
                    name="校验进入用户协议Activity",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".WebViewActivity",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=4,
                    name="校验用户协议标题",
                    action="assert",
                    locator=_id_locator("header_title_tv"),
                    assertion="text_equals",
                    expected="用户协议",
                    timeout=max(timeout, 15.0),
                ),
            ]
        )
        return steps

    raise SingleSheetMigrationError(f"不支持的 Android 登录用例：{case_id}")


def _home_steps(
    *,
    row: int,
    case_id: str,
    input_data: str | None,
    timeout: float,
) -> list[ExcelStep]:
    steps = [
        _step(
            row=row,
            case_id=case_id,
            order=1,
            name="恢复到首页",
            action="restart_to_home",
            timeout=max(timeout, 30.0),
        )
    ]
    if case_id in ANDROID_HOME_SEARCH_CASES:
        return _home_search_steps(
            row=row,
            case_id=case_id,
            input_data=input_data,
            timeout=timeout,
            steps=steps,
        )
    if case_id in ANDROID_CASE_SEARCH_CASES | ANDROID_CASE_FILTER_CASES:
        return _case_library_steps(
            row=row,
            case_id=case_id,
            input_data=input_data,
            timeout=timeout,
            steps=steps,
        )
    if case_id == "TC-HOME-001":
        modules = (
            ("顾客档案", "main_records_ll"),
            ("美际学院", "main_meiji_ll"),
            ("案例库", "main_case_ll"),
            ("设置", "main_set_cl"),
        )
        for order, (name, resource_name) in enumerate(modules, start=2):
            steps.append(
                _step(
                    row=row,
                    case_id=case_id,
                    order=order,
                    name=f"校验{name}入口可见",
                    action="assert",
                    locator=_id_locator(resource_name),
                    assertion="element_visible",
                    timeout=timeout,
                )
            )
        steps.append(
            _step(
                row=row,
                case_id=case_id,
                order=6,
                name="校验仍在首页",
                action="assert",
                assertion="activity_endswith",
                expected=".MainActivity",
                timeout=timeout,
            )
        )
        return steps

    destination = HOME_DESTINATIONS.get(case_id)
    if destination is None:
        raise SingleSheetMigrationError(f"不支持的 Android 首页用例：{case_id}")
    resource_name, activity, action_name = destination
    steps.extend(
        [
            _step(
                row=row,
                case_id=case_id,
                order=2,
                name=action_name,
                action="click",
                locator=_id_locator(resource_name),
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=3,
                name="校验目标页面",
                action="assert",
                assertion="activity_endswith",
                expected=activity,
                timeout=max(timeout, 15.0),
            ),
        ]
    )
    return steps


def _home_search_steps(
    *,
    row: int,
    case_id: str,
    input_data: str | None,
    timeout: float,
    steps: list[ExcelStep],
) -> list[ExcelStep]:
    steps.extend(
        [
            _step(
                row=row,
                case_id=case_id,
                order=2,
                name="输入首页搜索词",
                action="input",
                locator=_id_locator("main_search_et"),
                input_value=input_data,
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=3,
                name="点击首页搜索",
                action="click",
                locator=_id_locator("main_search_tv"),
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=4,
                name="校验搜索结果页",
                action="assert",
                assertion="activity_endswith",
                expected=".CustomerRecordsActivity",
                timeout=max(timeout, 15.0),
            ),
        ]
    )
    definition = get_android_case_definition(case_id)
    if definition is None:
        raise SingleSheetMigrationError(f"未登记的 Android 首页搜索用例：{case_id}")
    if definition.result_state == "empty":
        steps.append(
            _step(
                row=row,
                case_id=case_id,
                order=5,
                name="校验空结果提示",
                action="assert",
                locator=_id_locator("empty_tv"),
                assertion="element_visible",
                timeout=max(timeout, 10.0),
            )
        )
    else:
        exact_match = definition.result_state == "exact"
        steps.append(
            _step(
                row=row,
                case_id=case_id,
                order=5,
                name=("校验仅有一张顾客卡片" if exact_match else "校验搜索到顾客卡片"),
                action="assert",
                locator=_id_locator("a_records_cl"),
                assertion=("element_count_equals" if exact_match else "element_count_gte"),
                expected="1",
                timeout=max(timeout, 15.0),
            )
        )
        if definition.filter_locator:
            expected = (input_data or definition.filter_value or "").strip()
            if not expected:
                raise SingleSheetMigrationError(f"{case_id} 缺少搜索结果文本断言")
            steps.append(
                _step(
                    row=row,
                    case_id=case_id,
                    order=6,
                    name=(
                        "校验结果手机号精确匹配"
                        if exact_match
                        else "校验结果手机号包含搜索片段"
                    ),
                    action="assert",
                    locator=_id_locator(definition.filter_locator),
                    assertion=("text_equals" if exact_match else "text_contains"),
                    expected=expected,
                    timeout=max(timeout, 15.0),
                    note=(
                        "真实手机号仅在本地运行时解析，不写入日志、报告或工作簿"
                        if exact_match
                        else None
                    ),
                )
            )
    return steps


def _case_library_steps(
    *,
    row: int,
    case_id: str,
    input_data: str | None,
    timeout: float,
    steps: list[ExcelStep],
) -> list[ExcelStep]:
    definition = get_android_case_definition(case_id)
    if definition is None:
        raise SingleSheetMigrationError(f"未登记的 Android 案例库用例：{case_id}")
    tag = (input_data or definition.filter_value or "").strip()
    if not tag:
        raise SingleSheetMigrationError(f"{case_id} 缺少案例标签")

    steps.extend(
        [
            _step(
                row=row,
                case_id=case_id,
                order=2,
                name="进入案例库",
                action="click",
                locator=_id_locator("main_case_ll"),
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=3,
                name="校验案例库Activity",
                action="assert",
                assertion="activity_endswith",
                expected=".CaseActivity",
                timeout=max(timeout, 15.0),
            ),
            _step(
                row=row,
                case_id=case_id,
                order=4,
                name="校验案例列表",
                action="assert",
                locator=_id_locator("case_rv"),
                assertion="element_visible",
                timeout=max(timeout, 15.0),
            ),
        ]
    )
    if case_id in ANDROID_CASE_SEARCH_CASES:
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=5,
                    name="输入案例标签",
                    action="input",
                    locator=_id_locator("case_search_et"),
                    input_value=tag,
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=6,
                    name="执行案例搜索",
                    action="click",
                    locator=_id_locator("case_search_tv"),
                    timeout=timeout,
                ),
            ]
        )
    elif case_id in ANDROID_CASE_FILTER_CASES:
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=5,
                    name="展开案例标签",
                    action="click",
                    locator=_id_locator("case_tag_expand_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=6,
                    name=f"筛选案例标签{tag}",
                    action="click",
                    locator=_case_tag_locator(tag),
                    timeout=timeout,
                ),
            ]
        )
    else:
        raise SingleSheetMigrationError(f"不支持的 Android 案例库用例：{case_id}")

    if definition.result_state == "empty":
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=7,
                    name="校验空结果提示",
                    action="assert",
                    locator=_id_locator("empty_tv"),
                    assertion="element_visible",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=8,
                    name="校验不存在案例卡片",
                    action="assert",
                    locator=_id_locator("a_case_image_cl"),
                    assertion="element_not_visible",
                    timeout=max(timeout, 15.0),
                ),
            ]
        )
        return steps

    steps.extend(
        [
            _step(
                row=row,
                case_id=case_id,
                order=7,
                name="校验存在案例卡片",
                action="assert",
                locator=_id_locator("a_case_image_cl"),
                assertion="element_count_gte",
                expected="1",
                timeout=max(timeout, 15.0),
            ),
            _step(
                row=row,
                case_id=case_id,
                order=8,
                name=f"校验案例卡片包含标签{tag}",
                action="assert",
                locator=_case_card_tag_locator(tag),
                assertion="element_count_gte",
                expected="1",
                timeout=max(timeout, 15.0),
            ),
        ]
    )
    return steps


def _customer_list_steps(*, row: int, case_id: str, timeout: float) -> list[ExcelStep]:
    return [
        _step(
            row=row,
            case_id=case_id,
            order=1,
            name="恢复到首页",
            action="restart_to_home",
            timeout=max(timeout, 30.0),
        ),
        _step(
            row=row,
            case_id=case_id,
            order=2,
            name="进入顾客档案",
            action="click",
            locator=_id_locator("main_records_ll"),
            timeout=timeout,
        ),
        _step(
            row=row,
            case_id=case_id,
            order=3,
            name="校验顾客列表",
            action="assert",
            locator=_id_locator("customer_records_rv"),
            assertion="element_visible",
            timeout=max(timeout, 15.0),
        ),
    ]


def _customer_steps(
    *,
    row: int,
    case_id: str,
    input_data: str | None,
    timeout: float,
) -> list[ExcelStep]:
    steps = _customer_list_steps(row=row, case_id=case_id, timeout=timeout)
    if case_id in ANDROID_CUSTOMER_SEARCH_CASES:
        return _customer_search_steps(
            row=row,
            case_id=case_id,
            input_data=input_data,
            timeout=timeout,
            steps=steps,
        )
    if case_id in ANDROID_CUSTOMER_FILTER_CASES:
        return _customer_filter_steps(
            row=row,
            case_id=case_id,
            timeout=timeout,
            steps=steps,
        )
    if case_id in ANDROID_CUSTOMER_FILTER_RESET_CASES:
        return _customer_filter_reset_steps(
            row=row,
            case_id=case_id,
            input_data=input_data,
            timeout=timeout,
            steps=steps,
        )
    if case_id in ANDROID_CUSTOMER_DATE_FILTER_CASES:
        return _customer_date_filter_steps(
            row=row,
            case_id=case_id,
            input_data=input_data,
            timeout=timeout,
            steps=steps,
        )
    if case_id == "TC-CUSTOMER-020":
        fields = (
            ("顾客卡片", "a_records_cl", "element_count_gte", "1"),
            ("头像/性别图像", "a_records_head_ifv", "element_visible", None),
            ("姓名字段", "a_records_name_tv", "element_visible", None),
            ("年龄字段", "a_records_age_tv", "element_visible", None),
            ("手机号字段", "a_records_unique_tv", "element_visible", None),
            ("上次检测时间", "a_records_time_tv", "element_visible", None),
            ("检测次数字段", "a_records_count_tv", "element_visible", None),
        )
        for order, (name, resource_name, assertion, expected) in enumerate(fields, start=4):
            steps.append(
                _step(
                    row=row,
                    case_id=case_id,
                    order=order,
                    name=f"校验{name}",
                    action="assert",
                    locator=_id_locator(resource_name),
                    assertion=assertion,
                    expected=expected,
                    timeout=max(timeout, 15.0),
                )
            )
        return steps

    if case_id == "TC-CUSTOMER-021":
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=4,
                    name="点击第一张顾客卡片",
                    action="click",
                    locator=_id_locator("a_records_cl"),
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=5,
                    name="校验顾客详情Activity",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".CustomerDetailActivity",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=6,
                    name="校验详情页面",
                    action="assert",
                    locator=_id_locator("customer_detail_name_tv"),
                    assertion="element_visible",
                    timeout=max(timeout, 15.0),
                ),
            ]
        )
        return steps

    raise SingleSheetMigrationError(f"不支持的 Android 顾客用例：{case_id}")


def _customer_search_steps(
    *,
    row: int,
    case_id: str,
    input_data: str | None,
    timeout: float,
    steps: list[ExcelStep],
) -> list[ExcelStep]:
    steps.extend(
        [
            _step(
                row=row,
                case_id=case_id,
                order=4,
                name="输入顾客搜索词",
                action="input",
                locator=_id_locator("customer_records_search_et"),
                input_value=input_data,
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=5,
                name="点击顾客搜索",
                action="click",
                locator=_id_locator("customer_records_search_tv"),
                timeout=timeout,
            ),
        ]
    )
    definition = get_android_case_definition(case_id)
    if definition is None:
        raise SingleSheetMigrationError(f"未登记的 Android 顾客搜索用例：{case_id}")
    if definition.result_state == "empty":
        steps.append(
            _step(
                row=row,
                case_id=case_id,
                order=6,
                name="校验空结果提示",
                action="assert",
                locator=_id_locator("empty_tv"),
                assertion="element_visible",
                timeout=max(timeout, 10.0),
            )
        )
    else:
        steps.append(
            _step(
                row=row,
                case_id=case_id,
                order=6,
                name="校验搜索到顾客卡片",
                action="assert",
                locator=_id_locator("a_records_cl"),
                assertion="element_count_gte",
                expected="1",
                timeout=max(timeout, 15.0),
            )
        )
    return steps


def _customer_filter_steps(
    *,
    row: int,
    case_id: str,
    timeout: float,
    steps: list[ExcelStep],
) -> list[ExcelStep]:
    definition = get_android_case_definition(case_id)
    if definition is None or not definition.filter_value or not definition.filter_locator:
        raise SingleSheetMigrationError(f"未登记的 Android 顾客筛选用例：{case_id}")

    steps.extend(
        [
            _step(
                row=row,
                case_id=case_id,
                order=4,
                name="打开顾客筛选",
                action="click",
                locator=_id_locator("customer_records_other_tv"),
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=5,
                name="校验筛选弹层",
                action="assert",
                locator=_id_locator("customer_records_other_cl"),
                assertion="element_visible",
                timeout=max(timeout, 10.0),
            ),
            _step(
                row=row,
                case_id=case_id,
                order=6,
                name=f"选择筛选项：{definition.filter_value}",
                action="click",
                locator=_id_locator(definition.filter_locator),
                timeout=max(timeout, 10.0),
            ),
            _step(
                row=row,
                case_id=case_id,
                order=7,
                name="确认顾客筛选",
                action="click",
                locator=_id_locator("customer_records_other_confirm_tv"),
                timeout=max(timeout, 10.0),
            ),
        ]
    )
    if definition.result_state == "empty":
        steps.append(
            _step(
                row=row,
                case_id=case_id,
                order=8,
                name="校验空结果提示",
                action="assert",
                locator=_id_locator("empty_tv"),
                assertion="element_visible",
                timeout=max(timeout, 10.0),
            )
        )
    else:
        steps.append(
            _step(
                row=row,
                case_id=case_id,
                order=8,
                name="校验筛选结果",
                action="assert",
                locator=_id_locator("a_records_cl"),
                assertion="element_count_gte",
                expected="1",
                timeout=max(timeout, 15.0),
            )
        )
    return steps


def _customer_filter_reset_steps(
    *,
    row: int,
    case_id: str,
    input_data: str | None,
    timeout: float,
    steps: list[ExcelStep],
) -> list[ExcelStep]:
    remark = input_data or "123"
    male_locator = _id_locator("customer_records_other_sex_man_tv")
    age_locator = _id_locator("customer_records_other_age_1_tv")
    remark_locator = _id_locator("customer_records_other_remark_et")
    panel_locator = _id_locator("customer_records_other_cl")
    steps.extend(
        [
            _step(
                row=row,
                case_id=case_id,
                order=4,
                name="打开其他筛选",
                action="click",
                locator=_id_locator("customer_records_other_tv"),
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=5,
                name="校验其他筛选面板",
                action="assert",
                locator=panel_locator,
                assertion="element_visible",
                timeout=max(timeout, 10.0),
            ),
            _step(
                row=row,
                case_id=case_id,
                order=6,
                name="选择性别：男",
                action="click",
                locator=male_locator,
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=7,
                name="校验性别已选中",
                action="assert",
                locator=male_locator,
                assertion="element_selected",
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=8,
                name="选择年龄：18-25岁",
                action="click",
                locator=age_locator,
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=9,
                name="校验年龄已选中",
                action="assert",
                locator=age_locator,
                assertion="element_selected",
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=10,
                name="输入备注筛选词",
                action="input",
                locator=remark_locator,
                input_value=remark,
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=11,
                name="校验备注筛选词",
                action="assert",
                locator=remark_locator,
                assertion="attribute_equals",
                expected=f"text|{remark}",
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=12,
                name="重置筛选条件",
                action="click",
                locator=_id_locator("customer_records_other_reset_tv"),
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=13,
                name="校验性别已取消",
                action="assert",
                locator=male_locator,
                assertion="attribute_equals",
                expected="selected|false",
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=14,
                name="校验年龄已取消",
                action="assert",
                locator=age_locator,
                assertion="attribute_equals",
                expected="selected|false",
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=15,
                name="校验备注恢复占位文本",
                action="assert",
                locator=remark_locator,
                assertion="attribute_equals",
                expected="text|搜索备注记录",
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=16,
                name="校验重置后面板仍打开",
                action="assert",
                locator=panel_locator,
                assertion="element_visible",
                timeout=timeout,
            ),
        ]
    )
    return steps


def _customer_date_filter_steps(
    *,
    row: int,
    case_id: str,
    input_data: str | None,
    timeout: float,
    steps: list[ExcelStep],
) -> list[ExcelStep]:
    definition = get_android_case_definition(case_id)
    if definition is None or not definition.filter_value or not definition.filter_locator:
        raise SingleSheetMigrationError(f"未登记的 Android 日期筛选用例：{case_id}")

    date_range = _normalize_date_range(input_data or definition.filter_value)
    steps.extend(
        [
            _step(
                row=row,
                case_id=case_id,
                order=4,
                name="打开日期范围筛选",
                action="click",
                locator=_id_locator(definition.filter_locator),
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=5,
                name="校验日期弹层",
                action="assert",
                locator=_id_locator("customer_records_date_cl"),
                assertion="element_visible",
                timeout=max(timeout, 10.0),
            ),
            _step(
                row=row,
                case_id=case_id,
                order=6,
                name=f"滚动日期范围：{date_range}",
                action="select_date_range",
                input_value=date_range,
                timeout=max(timeout, 30.0),
            ),
            _step(
                row=row,
                case_id=case_id,
                order=7,
                name="确认日期范围",
                action="click",
                locator=_id_locator("customer_records_date_confirm_tv"),
                timeout=max(timeout, 10.0),
            ),
            _step(
                row=row,
                case_id=case_id,
                order=8,
                name="校验日期范围文本",
                action="assert",
                locator=_id_locator("customer_records_date_tv"),
                assertion="text_equals",
                expected=date_range.replace("/", "-"),
                timeout=max(timeout, 15.0),
            ),
        ]
    )
    if definition.filter_options:
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=9,
                    name="打开其他筛选",
                    action="click",
                    locator=_id_locator("customer_records_other_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=10,
                    name="校验其他筛选弹层",
                    action="assert",
                    locator=_id_locator("customer_records_other_cl"),
                    assertion="element_visible",
                    timeout=max(timeout, 10.0),
                ),
            ]
        )
        for order, (locator, value) in enumerate(definition.filter_options, start=11):
            steps.append(
                _step(
                    row=row,
                    case_id=case_id,
                    order=order,
                    name=f"选择联合筛选项：{value}",
                    action="click",
                    locator=_id_locator(locator),
                    timeout=max(timeout, 10.0),
                )
            )
        steps.append(
            _step(
                row=row,
                case_id=case_id,
                order=13,
                name="确认联合筛选",
                action="click",
                locator=_id_locator("customer_records_other_confirm_tv"),
                timeout=max(timeout, 10.0),
            )
        )
        result_order = 14
    else:
        result_order = 9

    if definition.result_state == "empty":
        steps.append(
            _step(
                row=row,
                case_id=case_id,
                order=result_order,
                name="校验日期筛选为空",
                action="assert",
                locator=_id_locator("empty_tv"),
                assertion="element_visible",
                timeout=max(timeout, 15.0),
            )
        )
    else:
        steps.append(
            _step(
                row=row,
                case_id=case_id,
                order=result_order,
                name="校验日期筛选结果",
                action="assert",
                locator=_id_locator("a_records_cl"),
                assertion="element_count_gte",
                expected="1",
                timeout=max(timeout, 15.0),
            )
        )
    return steps


def _detail_validation_steps(
    *,
    row: int,
    case_id: str,
    input_data: str | None,
    timeout: float,
) -> list[ExcelStep]:
    steps = _customer_list_steps(row=row, case_id=case_id, timeout=timeout)
    detail_order = 5
    if case_id in (
        CUSTOMER_MUTATION_CASE_IDS | TAG_MUTATION_CASE_IDS | REMARK_MUTATION_CASE_IDS
    ) or case_id in {
        "TC-DETAIL-017",
        "TC-DETAIL-019",
    }:
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=4,
                    name="输入专用顾客查询",
                    action="input",
                    locator=_id_locator("customer_records_search_et"),
                    input_value="${YANJIA_MUTATION_CUSTOMER_QUERY}",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=5,
                    name="执行专用顾客搜索",
                    action="click",
                    locator=_id_locator("customer_records_search_tv"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=6,
                    name="校验专用顾客唯一匹配",
                    action="assert",
                    locator=_id_locator("a_records_cl"),
                    assertion="element_count_equals",
                    expected="1",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=7,
                    name="打开唯一顾客卡片",
                    action="click",
                    locator=_id_locator("a_records_cl"),
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=8,
                    name="校验顾客详情Activity",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".CustomerDetailActivity",
                    timeout=max(timeout, 15.0),
                ),
            ]
        )
        detail_order = 8
    elif case_id in {
        "TC-DETAIL-029",
        "TC-DETAIL-030",
        "TC-DETAIL-032",
        "TC-DETAIL-034",
        "TC-DETAIL-036",
    }:
        search_value = (input_data or "").split("|", 1)[0].strip()
        if not search_value:
            raise SingleSheetMigrationError(f"{case_id} 缺少影像或咨询单种子顾客搜索值")
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=4,
                    name=(
                        "输入影像种子顾客"
                        if case_id in {"TC-DETAIL-029", "TC-DETAIL-030"}
                        else "输入咨询单种子顾客"
                    ),
                    action="input",
                    locator=_id_locator("customer_records_search_et"),
                    input_value=search_value,
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=5,
                    name="执行顾客搜索",
                    action="click",
                    locator=_id_locator("customer_records_search_tv"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=6,
                    name=(
                        "校验影像种子顾客唯一匹配"
                        if case_id in {"TC-DETAIL-029", "TC-DETAIL-030"}
                        else "校验咨询单种子顾客存在"
                    ),
                    action="assert",
                    locator=_id_locator("a_records_cl"),
                    assertion=(
                        "element_count_equals"
                        if case_id in {"TC-DETAIL-029", "TC-DETAIL-030"}
                        else "element_count_gte"
                    ),
                    expected="1",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=7,
                    name=(
                        "打开唯一顾客卡片"
                        if case_id in {"TC-DETAIL-029", "TC-DETAIL-030"}
                        else "点击第一张顾客卡片"
                    ),
                    action="click",
                    locator=_id_locator("a_records_cl"),
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=8,
                    name="校验顾客详情Activity",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".CustomerDetailActivity",
                    timeout=max(timeout, 15.0),
                ),
            ]
        )
        detail_order = 8
    else:
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=4,
                    name="点击第一张顾客卡片",
                    action="click",
                    locator=_id_locator("a_records_cl"),
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=5,
                    name="校验顾客详情Activity",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".CustomerDetailActivity",
                    timeout=max(timeout, 15.0),
                ),
            ]
        )

    if case_id == "TC-DETAIL-029":
        prompt = "请选择需要删除的影像"
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 1,
                    name="校验存在历史影像",
                    action="assert",
                    locator=_id_locator("a_records_detail_all_head_ifv"),
                    assertion="element_count_gte",
                    expected="1",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 2,
                    name="校验管理入口初始状态",
                    action="assert",
                    locator=_id_locator("customer_detail_manager_tv"),
                    assertion="text_equals",
                    expected="管理",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 3,
                    name="进入影像管理态",
                    action="click",
                    locator=_id_locator("customer_detail_manager_tv"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 4,
                    name="校验完成按钮",
                    action="assert",
                    locator=_id_locator("customer_detail_manager_tv"),
                    assertion="text_equals",
                    expected="完成",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 5,
                    name="校验影像选择控件",
                    action="assert",
                    locator=_id_locator("a_records_detail_all_select_ifv"),
                    assertion="element_visible",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 6,
                    name="校验删除按钮",
                    action="assert",
                    locator=_id_locator("customer_detail_delete_tv"),
                    assertion="element_visible",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 7,
                    name="未选择影像直接点击删除",
                    action="click",
                    locator=_id_locator("customer_detail_delete_tv"),
                    timeout=max(timeout, 10.0),
                    note="不勾选任何影像，仅验证未选择删除提示",
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 8,
                    name="校验未选择影像提示",
                    action="assert",
                    assertion="page_source_contains",
                    expected=prompt,
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 9,
                    name="退出影像管理态",
                    action="click",
                    locator=_id_locator("customer_detail_manager_tv"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 10,
                    name="校验恢复管理按钮",
                    action="assert",
                    locator=_id_locator("customer_detail_manager_tv"),
                    assertion="text_equals",
                    expected="管理",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 11,
                    name="校验选择控件已隐藏",
                    action="assert",
                    locator=_id_locator("a_records_detail_all_select_ifv"),
                    assertion="element_not_visible",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 12,
                    name="校验删除按钮已隐藏",
                    action="assert",
                    locator=_id_locator("customer_detail_delete_tv"),
                    assertion="element_not_visible",
                    timeout=max(timeout, 10.0),
                ),
            ]
        )
        return steps

    if case_id == "TC-DETAIL-030":
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 1,
                    name="校验存在历史影像",
                    action="assert",
                    locator=_id_locator("a_records_detail_all_head_ifv"),
                    assertion="element_count_gte",
                    expected="1",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 2,
                    name="打开首张未关联咨询单的影像",
                    action="open_first_unlinked_image",
                    timeout=max(timeout, 15.0),
                    note="按检测时间动态选择；若当前可见影像均已关联则安全失败",
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 3,
                    name="校验进入影像结果页",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".SkinResultActivity",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 4,
                    name="校验影像结果内容",
                    action="assert",
                    locator=_id_locator("skin_result_vp2"),
                    assertion="element_visible",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 5,
                    name="校验咨询单入口就绪",
                    action="assert",
                    locator=_id_locator("f_skin_result_info_view_case_tv"),
                    assertion="element_visible",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 6,
                    name="返回并触发保存退出确认",
                    action="click",
                    locator=_id_locator("skin_result_back_ll"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 7,
                    name="校验保存并退出选项",
                    action="assert",
                    locator=_id_locator("cover_prompt_v2_tv"),
                    assertion="text_equals",
                    expected="保存并退出",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 8,
                    name="保存咨询单并退出",
                    action="click",
                    locator=_id_locator("cover_prompt_v2_tv"),
                    timeout=max(timeout, 15.0),
                    note="新增咨询单按用户要求保留，不执行删除或回滚",
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 9,
                    name="校验返回顾客详情页",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".CustomerDetailActivity",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 10,
                    name="校验新增咨询单已保留",
                    action="verify_selected_image_consultation",
                    timeout=max(timeout, 15.0),
                    note="以所选影像检测时间出现在咨询单卡片中为准",
                ),
            ]
        )
        return steps

    if case_id == "TC-DETAIL-024":
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=6,
                    name="校验存在历史影像",
                    action="assert",
                    locator=_id_locator("a_records_detail_all_head_ifv"),
                    assertion="element_count_gte",
                    expected="1",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=7,
                    name="校验管理入口初始状态",
                    action="assert",
                    locator=_id_locator("customer_detail_manager_tv"),
                    assertion="text_equals",
                    expected="管理",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=8,
                    name="进入影像管理态",
                    action="click",
                    locator=_id_locator("customer_detail_manager_tv"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=9,
                    name="校验完成按钮",
                    action="assert",
                    locator=_id_locator("customer_detail_manager_tv"),
                    assertion="text_equals",
                    expected="完成",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=10,
                    name="校验影像选择控件",
                    action="assert",
                    locator=_id_locator("a_records_detail_all_select_ifv"),
                    assertion="element_visible",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=11,
                    name="校验删除按钮",
                    action="assert",
                    locator=_id_locator("customer_detail_delete_tv"),
                    assertion="element_visible",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=12,
                    name="退出影像管理态",
                    action="click",
                    locator=_id_locator("customer_detail_manager_tv"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=13,
                    name="校验恢复管理按钮",
                    action="assert",
                    locator=_id_locator("customer_detail_manager_tv"),
                    assertion="text_equals",
                    expected="管理",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=14,
                    name="校验选择控件已隐藏",
                    action="assert",
                    locator=_id_locator("a_records_detail_all_select_ifv"),
                    assertion="element_not_visible",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=15,
                    name="校验删除按钮已隐藏",
                    action="assert",
                    locator=_id_locator("customer_detail_delete_tv"),
                    assertion="element_not_visible",
                    timeout=max(timeout, 10.0),
                ),
            ]
        )
        return steps

    if case_id in TAG_MUTATION_CASE_IDS:
        image_case = case_id in {"TC-DETAIL-016", "TC-DETAIL-018", "TC-DETAIL-020"}
        label = {
            "TC-DETAIL-016": "自动化标签${RUN_TOKEN}",
            "TC-DETAIL-018": "%……&*${RUN_TOKEN}",
            "TC-DETAIL-020": "自动化删除标签${RUN_TOKEN}",
            "TC-DETAIL-031": "咨询单标签${RUN_TOKEN}",
            "TC-DETAIL-033": "%……&*${RUN_TOKEN}",
        }[case_id]
        if image_case:
            seed_name = "校验存在历史影像"
            entry_locator = "a_records_detail_all_remark_ifv"
            open_name = "进入第一张影像备注"
        else:
            seed_name = "校验存在咨询单"
            entry_locator = "a_consultation_result_remark_ifv"
            open_name = "进入第一条咨询单备注"
        restore_note = (
            "保存前记录标签集合；断言结束或异常退出时仅删除本轮差集新增标签，"
            "并验证原标签文本、数量和可删除控件数量全部恢复"
        )
        start_order = detail_order + 1
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order,
                    name=seed_name,
                    action="assert",
                    locator=_id_locator(entry_locator),
                    assertion="element_count_gte",
                    expected="1",
                    timeout=max(timeout, 15.0),
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 1,
                    name=open_name,
                    action="click",
                    locator=_id_locator(entry_locator),
                    timeout=max(timeout, 10.0),
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 2,
                    name="校验备注Activity",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".RecordsRemarkActivity",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 3,
                    name="打开新增标签弹窗",
                    action="click",
                    locator=_id_locator("a_records_remark_tag_add_ll"),
                    timeout=max(timeout, 10.0),
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 4,
                    name="校验标签弹窗",
                    action="assert",
                    locator=_id_locator("records_remark_tag_cl"),
                    assertion="element_visible",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 5,
                    name="输入本轮唯一标签",
                    action="input",
                    locator=_id_locator("records_remark_tag_et"),
                    input_value=label,
                    timeout=max(timeout, 10.0),
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 6,
                    name="保存标签",
                    action="click",
                    locator=_id_locator("records_remark_tag_save_tv"),
                    timeout=max(timeout, 10.0),
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 7,
                    name="校验新增标签显示",
                    action="assert",
                    locator=_id_locator("a_records_remark_tag_tv"),
                    assertion="text_contains",
                    expected=label,
                    timeout=max(timeout, 10.0),
                    note=restore_note,
                ),
            ]
        )
        if case_id == "TC-DETAIL-020":
            steps.extend(
                [
                    _step(
                        row=row,
                        case_id=case_id,
                        order=start_order + 8,
                        name="删除本轮唯一标签并确认",
                        action="delete_current_run_tag",
                        timeout=max(timeout, 15.0),
                        note=restore_note,
                    ),
                    _step(
                        row=row,
                        case_id=case_id,
                        order=start_order + 9,
                        name="校验标签集合已恢复",
                        action="assert",
                        assertion="activity_endswith",
                        expected=".RecordsRemarkActivity",
                        timeout=max(timeout, 10.0),
                        note=restore_note,
                    ),
                ]
            )
        return steps

    if case_id in REMARK_MUTATION_CASE_IDS:
        label = {
            "TC-DETAIL-012": "测试备注${RUN_TOKEN}",
            "TC-DETAIL-021": "影像备注${RUN_TOKEN}",
            "TC-DETAIL-022": "……&&*……*&${RUN_TOKEN}",
            "TC-DETAIL-035": "咨询单备注${RUN_TOKEN}",
            "TC-DETAIL-037": "自动化删除备注${RUN_TOKEN}",
        }[case_id]
        image_case = case_id in {"TC-DETAIL-012", "TC-DETAIL-021", "TC-DETAIL-022"}
        if image_case:
            seed_name = "校验存在历史影像"
            entry_locator = "a_records_detail_all_remark_ifv"
            open_name = "进入第一张影像备注"
        else:
            seed_name = "校验存在咨询单"
            entry_locator = "a_consultation_result_remark_ifv"
            open_name = "进入第一条咨询单备注"
        restore_note = (
            "提交前记录备注内容与可删除控件集合；断言结束或异常退出时仅删除"
            "本轮唯一新增备注，并验证原备注集合严格恢复"
        )
        start_order = detail_order + 1
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order,
                    name=seed_name,
                    action="assert",
                    locator=_id_locator(entry_locator),
                    assertion="element_count_gte",
                    expected="1",
                    timeout=max(timeout, 15.0),
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 1,
                    name=open_name,
                    action="click",
                    locator=_id_locator(entry_locator),
                    timeout=max(timeout, 10.0),
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 2,
                    name="校验备注Activity",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".RecordsRemarkActivity",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 3,
                    name="校验备注输入框",
                    action="assert",
                    locator=_id_locator("records_remark_remark_et"),
                    assertion="element_visible",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 4,
                    name="输入本轮唯一备注",
                    action="input",
                    locator=_id_locator("records_remark_remark_et"),
                    input_value=label,
                    timeout=max(timeout, 10.0),
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 5,
                    name="校验提交按钮启用",
                    action="assert",
                    locator=_id_locator("records_remark_remark_commit_tv"),
                    assertion="element_enabled",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 6,
                    name="提交本轮备注",
                    action="click",
                    locator=_id_locator("records_remark_remark_commit_tv"),
                    timeout=max(timeout, 10.0),
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 7,
                    name="校验新增备注显示",
                    action="assert",
                    locator=_id_locator("a_records_remark_list_content_tv"),
                    assertion="text_contains",
                    expected=label,
                    timeout=max(timeout, 15.0),
                    note=restore_note,
                ),
            ]
        )
        if case_id == "TC-DETAIL-037":
            steps.extend(
                [
                    _step(
                        row=row,
                        case_id=case_id,
                        order=start_order + 8,
                        name="删除本轮唯一备注并确认",
                        action="delete_current_run_remark",
                        timeout=max(timeout, 15.0),
                        note=restore_note,
                    ),
                    _step(
                        row=row,
                        case_id=case_id,
                        order=start_order + 9,
                        name="校验备注集合已恢复",
                        action="assert",
                        assertion="activity_endswith",
                        expected=".RecordsRemarkActivity",
                        timeout=max(timeout, 10.0),
                        note=restore_note,
                    ),
                ]
            )
        return steps

    if case_id in {
        "TC-DETAIL-017",
        "TC-DETAIL-019",
        "TC-DETAIL-032",
        "TC-DETAIL-034",
    }:
        image_case = case_id in {"TC-DETAIL-017", "TC-DETAIL-019"}
        whitespace_case = case_id in {"TC-DETAIL-019", "TC-DETAIL-032"}
        if image_case:
            seed_name = "校验存在历史影像"
            seed_locator = "a_records_detail_all_remark_ifv"
            open_name = "进入第一张影像备注"
            open_locator = "a_records_detail_all_remark_ifv"
        else:
            seed_name = "校验存在咨询单"
            seed_locator = "a_consultation_result_cl"
            open_name = "进入第一条咨询单备注"
            open_locator = "a_consultation_result_remark_ifv"
        start_order = detail_order + 1
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order,
                    name=seed_name,
                    action="assert",
                    locator=_id_locator(seed_locator),
                    assertion="element_count_gte",
                    expected="1",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 1,
                    name=open_name,
                    action="click",
                    locator=_id_locator(open_locator),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 2,
                    name="校验备注Activity",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".RecordsRemarkActivity",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 3,
                    name="打开新增标签弹窗",
                    action="click",
                    locator=_id_locator("a_records_remark_tag_add_ll"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 4,
                    name="校验标签弹窗",
                    action="assert",
                    locator=_id_locator("records_remark_tag_cl"),
                    assertion="element_visible",
                    timeout=max(timeout, 10.0),
                ),
            ]
        )
        next_order = start_order + 5
        if whitespace_case:
            steps.append(
                _step(
                    row=row,
                    case_id=case_id,
                    order=next_order,
                    name="输入单个空格标签",
                    action="input",
                    locator=_id_locator("records_remark_tag_et"),
                    input_value="${SPACE}",
                    timeout=max(timeout, 10.0),
                )
            )
            next_order += 1
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=next_order,
                    name="空格标签点击保存" if whitespace_case else "空标签点击保存",
                    action="click",
                    locator=_id_locator("records_remark_tag_save_tv"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=next_order + 1,
                    name="校验空标签提示",
                    action="assert",
                    assertion="page_source_contains",
                    expected="请输入标签内容",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=next_order + 2,
                    name="取消新增标签",
                    action="click",
                    locator=_id_locator("records_remark_tag_cancel_tv"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=next_order + 3,
                    name="关闭备注编辑器",
                    action="click",
                    locator=_id_locator("records_remark_close_ifv"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=next_order + 4,
                    name="校验返回顾客详情",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".CustomerDetailActivity",
                    timeout=max(timeout, 10.0),
                ),
            ]
        )
        return steps

    if case_id in {"TC-DETAIL-023", "TC-DETAIL-036"}:
        if case_id == "TC-DETAIL-023":
            seed_name = "校验存在历史影像"
            seed_locator = "a_records_detail_all_head_ifv"
            open_name = "进入第一张影像备注"
            open_locator = "a_records_detail_all_remark_ifv"
        else:
            seed_name = "校验存在咨询单"
            seed_locator = "a_consultation_result_cl"
            open_name = "进入第一条咨询单备注"
            open_locator = "a_consultation_result_remark_ifv"
        start_order = detail_order + 1
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order,
                    name=seed_name,
                    action="assert",
                    locator=_id_locator(seed_locator),
                    assertion="element_count_gte",
                    expected="1",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 1,
                    name=open_name,
                    action="click",
                    locator=_id_locator(open_locator),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 2,
                    name="校验备注Activity",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".RecordsRemarkActivity",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 3,
                    name="校验备注编辑器",
                    action="assert",
                    locator=_id_locator("records_remark_remark_et"),
                    assertion="element_visible",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 4,
                    name="输入单个空格",
                    action="input",
                    locator=_id_locator("records_remark_remark_et"),
                    input_value="${SPACE}",
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 5,
                    name="点击提交并校验备注未新增",
                    action="assert",
                    locator=_id_locator("records_remark_remark_commit_tv"),
                    assertion="element_count_unchanged_after_click",
                    expected=_id_locator("a_records_remark_list_content_tv"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 6,
                    name="关闭备注编辑器",
                    action="click",
                    locator=_id_locator("records_remark_close_ifv"),
                    timeout=max(timeout, 10.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=start_order + 7,
                    name="校验返回顾客详情",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".CustomerDetailActivity",
                    timeout=max(timeout, 10.0),
                ),
            ]
        )
        return steps

    steps.extend(
        [
            _step(
                row=row,
                case_id=case_id,
                order=detail_order + 1,
                name="校验详情编辑入口",
                action="click",
                locator=_id_locator("customer_detail_edit_tv"),
                timeout=max(timeout, 10.0),
            ),
            _step(
                row=row,
                case_id=case_id,
                order=detail_order + 2,
                name="校验编辑表单",
                action="assert",
                locator=_id_locator("customer_edit_v"),
                assertion="element_visible",
                timeout=max(timeout, 10.0),
            ),
        ]
    )

    if case_id in {"TC-DETAIL-026", "TC-DETAIL-027"}:
        if case_id == "TC-DETAIL-026":
            field_name = "customer_edit_username_et"
            target_value = "自动化${RUN_TOKEN}"
            card_locator = "a_records_name_tv"
            save_name = "保存本轮姓名"
            assert_name = "校验顾客卡片姓名同步"
        else:
            field_name = "customer_edit_unique_et"
            target_value = "${RUN_PHONE}"
            card_locator = "a_records_unique_tv"
            save_name = "保存本轮手机号"
            assert_name = "校验顾客卡片手机号同步"
        restore_note = (
            "卡片同步断言结束或异常退出后恢复运行前姓名、手机号、性别、生日、"
            "邮箱、婚姻状态、地址和个人备注，并逐字段比对"
        )
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 3,
                    name="输入本轮唯一姓名" if case_id == "TC-DETAIL-026" else "输入本轮手机号",
                    action="input",
                    locator=_id_locator(field_name),
                    input_value=target_value,
                    timeout=timeout,
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 4,
                    name=save_name,
                    action="click",
                    locator=_id_locator("customer_edit_save_tv"),
                    timeout=timeout,
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 5,
                    name="返回顾客列表",
                    action="back",
                    timeout=max(timeout, 10.0),
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 6,
                    name="校验顾客列表Activity",
                    action="assert",
                    assertion="activity_endswith",
                    expected=".CustomerRecordsActivity",
                    timeout=max(timeout, 15.0),
                    note=restore_note,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 7,
                    name=assert_name,
                    action="assert",
                    locator=_id_locator(card_locator),
                    assertion="text_equals",
                    expected=target_value,
                    timeout=max(timeout, 15.0),
                    note=restore_note,
                ),
            ]
        )
        return steps

    if case_id == "TC-DETAIL-007":
        birthday = (input_data or "").strip().strip("|")
        try:
            parsed_birthday = date.fromisoformat(birthday)
        except ValueError as error:
            raise SingleSheetMigrationError("TC-DETAIL-007 生日输入必须为 YYYY-MM-DD") from error
        target_birthday = parsed_birthday.isoformat()
        if target_birthday != "2022-08-04":
            raise SingleSheetMigrationError("TC-DETAIL-007 已按输入数据和验证点统一为 2022-08-04")
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 3,
                    name="选择生日2022-08-04",
                    action="select_birthday",
                    input_value=target_birthday,
                    timeout=max(timeout, 30.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 4,
                    name="保存生日",
                    action="click",
                    locator=_id_locator("customer_edit_save_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 5,
                    name="重新打开详情编辑",
                    action="click",
                    locator=_id_locator("customer_detail_edit_tv"),
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 6,
                    name="校验保存后的生日",
                    action="assert",
                    locator=_id_locator("customer_edit_birthday_tv"),
                    assertion="text_equals",
                    expected=target_birthday,
                    timeout=max(timeout, 10.0),
                ),
            ]
        )
        return steps

    if case_id in {"TC-DETAIL-008", "TC-DETAIL-009", "TC-DETAIL-010"}:
        if case_id == "TC-DETAIL-008":
            mutation_name = "清空地址"
            mutation_action = "clear"
            mutation_input = None
            assertion = "attribute_equals"
            expected = "text|"
        elif case_id == "TC-DETAIL-009":
            mutation_name = "输入地址特殊字符"
            mutation_action = "input"
            mutation_input = (input_data or "").strip().strip("|")
            if not mutation_input:
                raise SingleSheetMigrationError("TC-DETAIL-009 缺少地址特殊字符输入")
            assertion = "text_equals"
            expected = mutation_input
        else:
            mutation_name = "输入单个空格地址"
            mutation_action = "input"
            mutation_input = "${SPACE}"
            assertion = "attribute_equals"
            expected = "text|"
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 3,
                    name=mutation_name,
                    action=mutation_action,
                    locator=_id_locator("customer_edit_address_et"),
                    input_value=mutation_input,
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 4,
                    name="保存地址",
                    action="click",
                    locator=_id_locator("customer_edit_save_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 5,
                    name="重新打开详情编辑",
                    action="click",
                    locator=_id_locator("customer_detail_edit_tv"),
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 6,
                    name="校验保存后的地址",
                    action="assert",
                    locator=_id_locator("customer_edit_address_et"),
                    assertion=assertion,
                    expected=expected,
                    timeout=max(timeout, 10.0),
                ),
            ]
        )
        return steps

    if case_id == "TC-DETAIL-015":
        marital = (input_data or "").strip().strip("|")
        if marital != "保密":
            raise SingleSheetMigrationError("TC-DETAIL-015 婚姻状态输入必须为保密")
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 3,
                    name="展开婚姻状态选项",
                    action="click",
                    locator=_id_locator("customer_edit_marital_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 4,
                    name="选择婚姻状态保密",
                    action="click",
                    locator=_id_locator("customer_edit_marital_secret_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 5,
                    name="保存婚姻状态",
                    action="click",
                    locator=_id_locator("customer_edit_save_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 6,
                    name="重新打开详情编辑",
                    action="click",
                    locator=_id_locator("customer_detail_edit_tv"),
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 7,
                    name="校验保存后的婚姻状态",
                    action="assert",
                    locator=_id_locator("customer_edit_marital_tv"),
                    assertion="text_equals",
                    expected="保密",
                    timeout=max(timeout, 10.0),
                ),
            ]
        )
        return steps

    if case_id in {"TC-DETAIL-013", "TC-DETAIL-014"}:
        if case_id == "TC-DETAIL-013":
            mutation_name = "清空个人备注"
            mutation_action = "clear"
            mutation_input = None
            assertion = "attribute_equals"
            expected = "text|"
        else:
            mutation_name = "输入个人备注特殊字符"
            mutation_action = "input"
            mutation_input = (input_data or "").strip().strip("|")
            if not mutation_input:
                raise SingleSheetMigrationError("TC-DETAIL-014 缺少个人备注特殊字符输入")
            assertion = "text_equals"
            expected = mutation_input
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 3,
                    name="滚动到个人备注",
                    action="scroll_profile_remark",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 4,
                    name=mutation_name,
                    action=mutation_action,
                    locator=_id_locator("customer_edit_remark_et"),
                    input_value=mutation_input,
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 5,
                    name="保存个人备注",
                    action="click",
                    locator=_id_locator("customer_edit_save_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 6,
                    name="重新打开详情编辑",
                    action="click",
                    locator=_id_locator("customer_detail_edit_tv"),
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 7,
                    name="重新滚动到个人备注",
                    action="scroll_profile_remark",
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 8,
                    name="校验保存后的个人备注",
                    action="assert",
                    locator=_id_locator("customer_edit_remark_et"),
                    assertion=assertion,
                    expected=expected,
                    timeout=max(timeout, 10.0),
                ),
            ]
        )
        return steps

    if case_id in {"TC-DETAIL-004", "TC-DETAIL-005"}:
        expected_gender = "男" if case_id == "TC-DETAIL-004" else "女"
        input_gender = (input_data or "").strip().strip("|")
        if input_gender != expected_gender:
            raise SingleSheetMigrationError(f"{case_id} 性别输入必须为{expected_gender}")
        option_id = (
            "customer_edit_sex_man_tv" if expected_gender == "男" else "customer_edit_sex_female_tv"
        )
        steps.extend(
            [
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 3,
                    name="展开性别选项",
                    action="click",
                    locator=_id_locator("customer_edit_sex_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 4,
                    name=f"选择性别{expected_gender}",
                    action="click",
                    locator=_id_locator(option_id),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 5,
                    name="保存性别",
                    action="click",
                    locator=_id_locator("customer_edit_save_tv"),
                    timeout=timeout,
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 6,
                    name="重新打开详情编辑",
                    action="click",
                    locator=_id_locator("customer_detail_edit_tv"),
                    timeout=max(timeout, 15.0),
                ),
                _step(
                    row=row,
                    case_id=case_id,
                    order=detail_order + 7,
                    name="校验保存后的性别",
                    action="assert",
                    locator=_id_locator("customer_edit_sex_tv"),
                    assertion="text_equals",
                    expected=expected_gender,
                    timeout=max(timeout, 10.0),
                ),
            ]
        )
        return steps

    if case_id == "TC-DETAIL-001":
        field_name = "customer_edit_username_et"
        expected_message = "姓名不能为空"
        step_name = "清空姓名"
        mutation_action = "clear"
        mutation_input = None
    elif case_id == "TC-DETAIL-003":
        field_name = "customer_edit_username_et"
        expected_message = "客户名称长度不合法"
        step_name = "输入超长姓名"
        mutation_action = "input"
        mutation_input = (input_data or "").strip().strip("|")
        if not mutation_input:
            raise SingleSheetMigrationError("TC-DETAIL-003 缺少超长姓名输入值")
    elif case_id == "TC-DETAIL-011":
        field_name = "customer_edit_unique_et"
        expected_message = "手机号不能为空"
        step_name = "清空手机号"
        mutation_action = "clear"
        mutation_input = None
    else:
        raise SingleSheetMigrationError(f"不支持的 Android 详情用例：{case_id}")

    steps.extend(
        [
            _step(
                row=row,
                case_id=case_id,
                order=detail_order + 3,
                name=step_name,
                action=mutation_action,
                locator=_id_locator(field_name),
                input_value=mutation_input,
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=detail_order + 4,
                name="点击保存",
                action="click",
                locator=_id_locator("customer_edit_save_tv"),
                timeout=timeout,
            ),
            _step(
                row=row,
                case_id=case_id,
                order=detail_order + 5,
                name="校验资料校验提示",
                action="assert",
                assertion="page_source_contains",
                expected=expected_message,
                timeout=max(timeout, 5.0),
            ),
        ]
    )
    return steps


def _step(
    *,
    row: int,
    case_id: str,
    order: int,
    name: str,
    action: str,
    timeout: float,
    locator: str | None = None,
    input_value: str | None = None,
    assertion: str | None = None,
    expected: str | None = None,
    note: str | None = None,
) -> ExcelStep:
    return ExcelStep(
        source_row=row,
        case_id=case_id,
        order=order,
        name=name,
        action=action,
        locator=locator,
        input_value=input_value,
        assertion=assertion,
        expected=expected,
        timeout=timeout,
        element_index=0,
        continue_on_failure=False,
        enabled=True,
        note=note or "由更新后的单表用例迁移为 Android resource-id 步骤",
    )


def _input_pair(input_data: str | None) -> tuple[str, str]:
    parts = [part.strip() for part in (input_data or "").split("|")]
    username = parts[0] if parts and parts[0] else "${TEST_USERNAME}"
    password = parts[1] if len(parts) > 1 and parts[1] else "${TEST_PASSWORD}"
    return username, password


def _id_locator(name: str) -> str:
    return f"id={PACKAGE}:id/{name}"


def _case_tag_locator(tag: str) -> str:
    resource = _xpath_literal(f"{PACKAGE}:id/a_case_tag_tv")
    text = _xpath_literal(tag)
    return f"xpath=//*[@resource-id={resource} and @text={text}]"


def _case_card_tag_locator(tag: str) -> str:
    resource = _xpath_literal(f"{PACKAGE}:id/a_case_image_cl")
    text = _xpath_literal(tag)
    return f"xpath=//*[@resource-id={resource}]//*[@text={text}]"


def _xpath_literal(value: str) -> str:
    if "'" not in value:
        return f"'{value}'"
    if '"' not in value:
        return f'"{value}"'
    parts = value.split("'")
    return "concat(" + ', "\'", '.join(f"'{part}'" for part in parts) + ")"


def _normalize_date_range(value: str) -> str:
    raw = value.split("|")[-1].strip()
    match = re.fullmatch(
        r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})\s*(?:-|~)\s*"
        r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})",
        raw,
    )
    if match is None:
        raise SingleSheetMigrationError(f"日期范围格式错误：{value!r}")
    start_year, start_month, start_day, end_year, end_month, end_day = (
        int(part) for part in match.groups()
    )
    try:
        start = date(start_year, start_month, start_day)
        end = date(end_year, end_month, end_day)
    except ValueError as error:
        raise SingleSheetMigrationError(f"日期范围不是有效日期：{value!r}") from error
    if start > end:
        raise SingleSheetMigrationError(f"日期范围起始日期晚于结束日期：{value!r}")
    return f"{start:%Y-%m-%d}~{end:%Y-%m-%d}"


def _value(sheet: Worksheet, row: int, headers: dict[str, int], name: str) -> Any:
    column = headers.get(name)
    return sheet.cell(row, column).value if column is not None else None


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _optional_text(value: Any) -> str | None:
    text = _text(value)
    return text or None


def _number(value: Any, *, default: float) -> float:
    if value is None or _text(value) == "":
        return default
    try:
        result = float(value)
    except (TypeError, ValueError) as error:
        raise SingleSheetMigrationError(f"无法识别的超时：{value!r}") from error
    if result <= 0:
        raise SingleSheetMigrationError(f"超时必须大于0：{value!r}")
    return result


def _boolean(value: Any, *, default: bool) -> bool:
    text = _text(value).lower()
    if not text:
        return default
    if text in {"1", "true", "yes", "y", "on", "是", "启用", "执行"}:
        return True
    if text in {"0", "false", "no", "n", "off", "否", "禁用", "不执行"}:
        return False
    raise SingleSheetMigrationError(f"无法识别的是/否值：{value!r}")
