from __future__ import annotations

from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "test_case.xlsx"
PACKAGE = "com.xiaofutech.yanjia_ai"

# node, locator, action, verification, assertion
AUTOMATED = {
    "TC-LOGIN-007": (
        "tests/smoke/test_read_only_smoke.py::test_login_and_home_contract",
        "login_username_et,login_pwd_et,login_tv,login_store_rv,main_fbl",
        "input,input,click,click,verify",
        "MainActivity 且首页 main_fbl 可见",
        "activity_endswith+element_visible",
    ),
    "TC-HOME-001": (
        "tests/smoke/test_read_only_smoke.py::test_login_and_home_contract",
        "main_records_ll,main_meiji_ll,main_case_ll,main_set_cl",
        "verify",
        "首页四个入口均可见",
        "elements_visible",
    ),
    "TC-HOME-008": (
        "tests/smoke/test_read_only_smoke.py::test_customer_list_contract",
        "main_records_ll,customer_records_rv",
        "click,verify",
        "CustomerRecordsActivity 且列表可见",
        "activity_endswith+element_visible",
    ),
    "TC-CUSTOMER-005": (
        "tests/smoke/test_read_only_smoke.py::test_customer_list_contract",
        "a_records_cl,a_records_name_tv,a_records_age_tv,a_records_unique_tv,a_records_time_tv,a_records_count_tv",
        "verify",
        "至少一张顾客卡片且关键字段控件存在",
        "element_count+elements_visible",
    ),
    "TC-CUSTOMER-006": (
        "tests/smoke/test_read_only_smoke.py::test_first_customer_detail_contract",
        "a_records_cl,customer_detail_name_tv,customer_detail_info_tv,customer_detail_edit_tv",
        "click,verify",
        "CustomerDetailActivity 且详情唯一元素可见",
        "activity_endswith+elements_visible",
    ),
    "TC-HOME-010": (
        "tests/smoke/test_read_only_smoke.py::test_case_library_contract",
        "main_case_ll,case_rv",
        "click,verify",
        "CaseActivity 且案例列表可见",
        "activity_endswith+element_visible",
    ),
    "TC-CASE-001": (
        "tests/smoke/test_read_only_smoke.py::test_case_library_contract",
        "case_search_et,a_case_tag_tv,a_case_category_tv,case_rv",
        "verify",
        "案例搜索、标签、类别和列表控件存在",
        "elements_visible",
    ),
    "TC-HOME-011": (
        "tests/smoke/test_read_only_smoke.py::test_profile_contract",
        "main_set_cl,set_logout_tv",
        "click,verify",
        "SetActivity 且设置页根控件可见",
        "activity_endswith+element_visible",
    ),
    "TC-PROFILE-001": (
        "tests/smoke/test_read_only_smoke.py::test_profile_contract",
        "set_personal_ll,personal_account_tv",
        "click,verify",
        "账号字段非空；实际值不写入报告",
        "text_not_empty",
    ),
    "TC-PROFILE-002": (
        "tests/smoke/test_read_only_smoke.py::test_profile_contract",
        "personal_name_tv",
        "verify",
        "姓名字段非空；实际值不写入报告",
        "text_not_empty",
    ),
    "TC-PROFILE-003": (
        "tests/smoke/test_read_only_smoke.py::test_profile_contract",
        "personal_register_time_tv",
        "verify",
        "注册时间字段非空；实际值不写入报告",
        "text_not_empty",
    ),
    "TC-PROFILE-004": (
        "tests/smoke/test_read_only_smoke.py::test_profile_contract",
        "personal_phone_tv",
        "verify",
        "手机号字段存在，值按测试账号规则校验",
        "element_visible",
    ),
    "TC-IMAGE-001": (
        "tests/regression/test_seeded_image.py::test_existing_image_opens_viewer",
        "a_records_detail_all_head_ifv,skin_result_back_ll,skin_result_vp2",
        "click,verify",
        "SkinResultActivity 且影像 ViewPager 可见",
        "activity_endswith+element_visible",
    ),
}

INPUTS = {
    "TC-LOGIN-001": "${INVALID_USERNAME}|${YANJIA_PASSWORD}",
    "TC-LOGIN-002": "${YANJIA_USERNAME}|${INVALID_PASSWORD}",
    "TC-LOGIN-004": "${YANJIA_USERNAME}|<EMPTY>",
    "TC-LOGIN-005": "${YANJIA_USERNAME}|${YANJIA_PASSWORD}|<NO_STORE>",
    "TC-LOGIN-007": "${YANJIA_USERNAME}|${YANJIA_PASSWORD}|${YANJIA_STORE_OR_FIRST}",
    "TC-HOME-002": "${NON_EXISTENT_CUSTOMER_QUERY}",
    "TC-HOME-003": "${NON_EXISTENT_PHONE_QUERY}",
    "TC-HOME-004": "${SEEDED_CUSTOMER_NAME_PART}",
    "TC-HOME-005": "${SEEDED_CUSTOMER_NAME}",
    "TC-HOME-006": "${SEEDED_CUSTOMER_PHONE_PART}",
    "TC-HOME-007": "${SEEDED_CUSTOMER_PHONE}",
    "TC-CUSTOMER-001": "${EMPTY_DATE_RANGE}",
    "TC-CUSTOMER-002": "${SEEDED_DATE_RANGE}",
    "TC-CUSTOMER-004": "${SEEDED_CUSTOMER_UNDER_18}",
    "TC-IMAGE-001": "${SEEDED_CUSTOMER_WITH_IMAGE}",
    "TC-IMAGE-002": "${SEEDED_CUSTOMER_WITH_IMAGE}",
    "TC-IMAGE-003": "${SEEDED_CUSTOMER_WITH_IMAGE}",
    "TC-IMAGE-004": "${SEEDED_CUSTOMER_WITH_HISTORY_IMAGES}",
    "TC-IMAGE-005": "${SEEDED_CUSTOMER_WITH_HISTORY_IMAGES}",
    "TC-CASE-002": "${NON_EXISTENT_CASE_TAG}",
    "TC-CASE-003": "${SEEDED_CASE_TAG}",
    "TC-CASE-004": "${SEEDED_CASE_TAG}",
}

DESTRUCTIVE = {"TC-DETAIL-009", "TC-CASE-005"}
MUTATING = {"TC-DETAIL-002", "TC-DETAIL-003", "TC-DETAIL-008"}
AUTH_NEGATIVE = {f"TC-LOGIN-{number:03d}" for number in range(1, 6)}
MODULES = {
    "账号登录": "login",
    "首页": "home",
    "首页搜索": "home",
    "首页跳转": "home",
    "顾客列表": "customer",
    "顾客详情": "detail",
    "影像阅览": "image",
    "案例库": "case",
    "个人中心": "profile",
}


def main() -> None:
    workbook = load_workbook(WORKBOOK)
    sheet = workbook["自动化测试用例"]
    headers: dict[str, int] = {}
    for cell in sheet[1]:
        if isinstance(cell.value, str) and isinstance(cell.column, int):
            headers[cell.value] = cell.column
    for name in ("自动化状态", "pytest节点", "标签", "移动端备注"):
        if name not in headers:
            headers[name] = sheet.max_column + 1
            target = sheet.cell(1, headers[name], name)
            _copy_style(sheet.cell(1, headers["实际结果"]), target)
            sheet.column_dimensions[target.column_letter].width = 28

    row_by_id: dict[str, int] = {}
    for row in range(2, sheet.max_row + 1):
        case_id = str(_get(sheet, row, headers, "用例ID") or "").strip()
        if not case_id:
            continue
        row_by_id[case_id] = row
        module = str(_get(sheet, row, headers, "模块") or "").strip()
        priority = str(_get(sheet, row, headers, "优先级") or "").upper()
        _set(sheet, row, headers, "优先级", priority)
        _set(sheet, row, headers, "实际结果", "NOT_RUN")
        _normalize_timeout(sheet, row, headers)

        if case_id in INPUTS:
            _set(sheet, row, headers, "输入数据", INPUTS[case_id])
        if case_id.startswith("TC-PROFILE-") or case_id == "TC-CASE-001":
            _set(sheet, row, headers, "输入数据", "")
        if case_id in {"TC-HOME-006", "TC-HOME-007"}:
            _set(sheet, row, headers, "数据类型", "string")
        if case_id in {"TC-LOGIN-001", "TC-LOGIN-002", "TC-LOGIN-007"}:
            _set(sheet, row, headers, "数据类型", "string|string")

        status, node, note = "BACKLOG", "", "待按真实页面补充 Appium 定位与断言"
        if case_id in AUTOMATED:
            node, locators, actions, verification, assertion = AUTOMATED[case_id]
            status = "AUTOMATED_EXTENDED" if case_id == "TC-IMAGE-001" else "AUTOMATED"
            _set(sheet, row, headers, "元素定位器", _ids(locators))
            _set(sheet, row, headers, "操作类型", actions)
            _set(sheet, row, headers, "验证点", verification)
            _set(sheet, row, headers, "断言类型", assertion)
            note = "已根据 Android 16 平板真实 resource-id 实现"
        elif case_id in DESTRUCTIVE:
            status, note = "DISABLED_DESTRUCTIVE", "只能删除本轮创建且可精确恢复的数据"
        elif case_id == "TC-DETAIL-002":
            status, note = "NEEDS_PRODUCT_RULE", "特殊字符应保存还是拦截需要产品确认"

        _set(sheet, row, headers, "自动化状态", status)
        _set(sheet, row, headers, "pytest节点", node)
        _set(sheet, row, headers, "标签", ",".join(_tags(case_id, module, priority)))
        _set(sheet, row, headers, "移动端备注", note)

    _fix_cases(sheet, headers, row_by_id)
    _update_references(workbook)
    _write_notes(workbook)
    workbook.save(WORKBOOK)
    print("Normalized test_case.xlsx without reading env.txt or copying credential values.")


def _fix_cases(sheet, headers: dict[str, int], rows: dict[str, int]) -> None:
    if "TC-HOME-009" in rows:
        row = rows["TC-HOME-009"]
        _set(sheet, row, headers, "测试场景", "空关键字搜索")
        _set(sheet, row, headers, "测试点", "验证空关键字搜索的产品行为")
        _set(sheet, row, headers, "期望结果", "按产品规则展示默认结果或输入提示")
    image_points = {
        "TC-IMAGE-002": "验证分屏对比功能",
        "TC-IMAGE-003": "验证镜像对比功能",
        "TC-IMAGE-004": "验证历史分屏对比功能",
        "TC-IMAGE-005": "验证历史镜像对比功能",
    }
    for case_id, point in image_points.items():
        _set(sheet, rows[case_id], headers, "测试点", point)
    _set(
        sheet,
        rows["TC-CUSTOMER-006"],
        headers,
        "期望结果",
        "跳转至顾客详情页，详情唯一元素可见",
    )
    _set(sheet, rows["TC-CASE-001"], headers, "期望结果", "进入案例库页面，case_rv 可见")
    for case_id, noun in (("TC-DETAIL-009", "影像"), ("TC-CASE-005", "案例")):
        row = rows[case_id]
        _set(sheet, row, headers, "测试场景", f"删除本轮创建的指定测试{noun}")
        _set(
            sheet,
            row,
            headers,
            "前置条件",
            f"隔离环境；{noun}由本轮创建且有唯一标识；允许 destructive",
        )
        _set(sheet, row, headers, "期望结果", f"仅指定测试{noun}被删除，其他数据不变")


def _update_references(workbook) -> None:
    fields = workbook["字段说明"]
    for row in range(2, fields.max_row + 1):
        if fields.cell(row, 1).value == "元素定位器":
            fields.cell(row, 2, "Appium定位器；优先resource-id，其次accessibility id/UiSelector")
            fields.cell(row, 3, f"id={PACKAGE}:id/main_records_ll")
    assertions = workbook["断言类型说明"]
    assertions.cell(1, 3, "Appium/pytest实现示例")
    for row in range(2, assertions.max_row + 1):
        key = assertions.cell(row, 1).value
        if key in {"url_contains", "url_matches"}:
            assertions.cell(row, 3, "仅WebView；原生页改用Activity与唯一元素")
        elif key == "element_visible":
            assertions.cell(row, 3, "assert element.is_displayed()")
        elif key == "text_not_empty":
            assertions.cell(row, 3, "assert element.text.strip()")
    actions = workbook["操作类型说明"]
    actions.cell(1, 3, "Appium实现示例")
    examples = {
        "click": "driver.find_element(AppiumBy.ID, '...').click()",
        "input": "element.clear(); element.send_keys(value)",
        "verify": "WebDriverWait显式等待后assert",
        "hover": "Android不适用；改为long_click或删除",
        "wait": "WebDriverWait显式等待；禁止固定sleep",
        "nav": "点击页面入口；原生应用不使用page.goto",
        "screenshot": "driver.save_screenshot(path)；注意PII门禁",
    }
    for row in range(2, actions.max_row + 1):
        key = actions.cell(row, 1).value
        if key in examples:
            actions.cell(row, 3, examples[key])


def _write_notes(workbook) -> None:
    notes = (
        workbook["移动端执行说明"]
        if "移动端执行说明" in workbook.sheetnames
        else workbook.create_sheet("移动端执行说明")
    )
    if notes.max_row:
        notes.delete_rows(1, notes.max_row)
    values = [
        ("项目", "内容"),
        ("包名", PACKAGE),
        ("已验证版本", "2.1.9(Build 8) / versionCode 20109"),
        ("执行框架", "Appium 3 + UiAutomator2 + Python + pytest"),
        ("定位优先级", "resource-id > accessibility id > UiSelector > XPath"),
        ("实际结果", "由pytest/JUnit/Allure产生；本表保持NOT_RUN"),
        ("凭据", "只引用环境变量，不保存真实值"),
        ("破坏性", "默认禁用；只能作用于本轮创建且可恢复的数据"),
        ("隐私", "顾客、影像、页面树和截图可能含PII，附件默认关闭"),
    ]
    for row_index, row_values in enumerate(values, 1):
        for column_index, value in enumerate(row_values, 1):
            notes.cell(row_index, column_index, value)
    notes.column_dimensions["A"].width = 18
    notes.column_dimensions["B"].width = 90
    for cell in notes[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    for row in notes.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def _tags(case_id: str, module: str, priority: str) -> list[str]:
    tags = [priority.lower(), MODULES.get(module, "other"), "tablet"]
    if case_id in DESTRUCTIVE:
        tags += ["destructive", "serial"]
    elif case_id in MUTATING:
        tags += ["mutating", "serial"]
    else:
        tags.append("readonly")
    tags += ["auth_negative", "no_retry"] if case_id in AUTH_NEGATIVE else ["requires_auth"]
    if case_id.startswith(("TC-CUSTOMER", "TC-IMAGE")):
        tags.append("requires_seed")
    if case_id in AUTOMATED:
        tags.append("extended" if case_id == "TC-IMAGE-001" else "smoke")
    return list(dict.fromkeys(filter(None, tags)))


def _ids(names: str) -> str:
    return ",".join(f"id={PACKAGE}:id/{name}" for name in names.split(","))


def _normalize_timeout(sheet, row: int, headers: dict[str, int]) -> None:
    value = _get(sheet, row, headers, "超时(秒)")
    if value not in (None, ""):
        try:
            _set(sheet, row, headers, "超时(秒)", int(value))
        except (TypeError, ValueError):
            pass


def _get(sheet, row: int, headers: dict[str, int], name: str):
    return sheet.cell(row, headers[name]).value


def _set(sheet, row: int, headers: dict[str, int], name: str, value) -> None:
    sheet.cell(row, headers[name], value)


def _copy_style(source, target) -> None:
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.alignment = copy(source.alignment)
    target.number_format = source.number_format


if __name__ == "__main__":
    main()
