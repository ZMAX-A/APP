from __future__ import annotations

from fnmatch import fnmatchcase
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from yanjia_automation.excel.models import ExcelCase, ExcelStep
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

ANDROID_LOGIN_CASES = frozenset(
    {
        "TC-LOGIN-001",
        "TC-LOGIN-002",
        "TC-LOGIN-003",
        "TC-LOGIN-007",
    }
)

ANDROID_ERROR_EXPECTATIONS = {
    "TC-LOGIN-002": "账号密码错误，请重新输入",
}


class SingleSheetMigrationError(ValueError):
    """Raised when a single-sheet Web case cannot be safely mapped to Android."""


def load_android_login_cases(
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
        if not case_id or not any(
            fnmatchcase(case_id, pattern) for pattern in case_patterns
        ):
            continue
        if case_id in seen_ids:
            raise SingleSheetMigrationError(f"主表第{row}行用例ID重复：{case_id}")
        seen_ids.add(case_id)

        if not _boolean(_value(sheet, row, headers, "是否执行"), default=True):
            continue
        if case_id not in ANDROID_LOGIN_CASES:
            raise SingleSheetMigrationError(
                f"{case_id} 尚未迁移为 Android Appium 用例；"
                "当前表中的定位器是 Web/Playwright 语法。"
            )

        cases.append(_build_login_case(sheet, headers, row, case_id))

    return cases


def _build_login_case(
    sheet: Worksheet,
    headers: dict[str, int],
    row: int,
    case_id: str,
) -> ExcelCase:
    input_data = _optional_text(_value(sheet, row, headers, "输入数据"))
    expected = _optional_text(_value(sheet, row, headers, "期望结果"))
    timeout = _number(_value(sheet, row, headers, "超时(秒)"), default=10.0)
    steps = _login_steps(
        row=row,
        case_id=case_id,
        input_data=input_data,
        expected=expected,
        timeout=timeout,
    )
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
        tags=("tablet", "readonly", "login"),
        enabled=True,
        automation_status="AUTOMATED_ANDROID_MIGRATED",
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

    raise SingleSheetMigrationError(f"不支持的 Android 登录用例：{case_id}")


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
        note="由更新后的单表用例迁移为 Android resource-id 步骤",
    )


def _input_pair(input_data: str | None) -> tuple[str, str]:
    parts = [part.strip() for part in (input_data or "").split("|")]
    username = parts[0] if parts and parts[0] else "${TEST_USERNAME}"
    password = parts[1] if len(parts) > 1 and parts[1] else "${TEST_PASSWORD}"
    return username, password


def _id_locator(name: str) -> str:
    return f"id={PACKAGE}:id/{name}"


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
