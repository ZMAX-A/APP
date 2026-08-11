from __future__ import annotations

from appium.webdriver.common.appiumby import AppiumBy

Locator = tuple[str, str]
LocatorCandidates = tuple[Locator, ...]

LOCATOR_STRATEGIES = {
    "id": AppiumBy.ID,
    "resource_id": AppiumBy.ID,
    "accessibility_id": AppiumBy.ACCESSIBILITY_ID,
    "accessibility": AppiumBy.ACCESSIBILITY_ID,
    "aid": AppiumBy.ACCESSIBILITY_ID,
    "xpath": AppiumBy.XPATH,
    "uiautomator": AppiumBy.ANDROID_UIAUTOMATOR,
    "android_uiautomator": AppiumBy.ANDROID_UIAUTOMATOR,
    "class_name": AppiumBy.CLASS_NAME,
    "android_viewtag": AppiumBy.ANDROID_VIEWTAG,
}


class LocatorFormatError(ValueError):
    """Raised when an Excel locator does not use a supported strategy."""


def parse_locator(value: str | None) -> Locator:
    if not value or not value.strip():
        raise LocatorFormatError("当前操作需要元素定位器，但Excel单元格为空")
    if "=" not in value:
        raise LocatorFormatError(
            f"定位器格式错误：{value!r}；正确示例为 id=com.example:id/button"
        )
    strategy_name, locator_value = value.split("=", 1)
    strategy = LOCATOR_STRATEGIES.get(strategy_name.strip().lower())
    if strategy is None:
        supported = ", ".join(sorted(LOCATOR_STRATEGIES))
        raise LocatorFormatError(f"不支持的定位方式：{strategy_name!r}；支持：{supported}")
    locator_value = locator_value.strip()
    if not locator_value:
        raise LocatorFormatError(f"定位器值为空：{value!r}")
    return strategy, locator_value


def parse_locator_candidates(value: str | None) -> LocatorCandidates:
    """Parse one or more fallback locators separated by ``||``.

    Excel example: ``id=com.example:id/save || accessibility_id=保存``.
    Candidates are evaluated from left to right at runtime.
    """
    if not value or not value.strip():
        raise LocatorFormatError("当前操作需要元素定位器，但Excel单元格为空")
    raw_candidates = value.split("||")
    if any(not candidate.strip() for candidate in raw_candidates):
        raise LocatorFormatError(f"候选定位器中存在空项：{value!r}")
    return tuple(parse_locator(candidate.strip()) for candidate in raw_candidates)
