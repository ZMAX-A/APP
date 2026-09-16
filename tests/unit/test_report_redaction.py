from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory
from uuid import uuid4

import pytest


@pytest.mark.parametrize("suffix", ["", "\\'\"\n中文"])
def test_real_pytest_and_allure_scrub_nested_exceptions_and_captured_output(
    tmp_path: Path, suffix: str
) -> None:
    root = Path(__file__).resolve().parents[2]
    secret = f"fake-sensitive-{uuid4().hex}{suffix}"
    case_source = '''import logging
import os
import allure
import pytest
from selenium.common.exceptions import WebDriverException

def test_nested_assertion():
    secret = os.environ["YANJIA_AUDIT_SECRET"]
    with allure.step("outer " + secret):
        with allure.step("inner"):
            print(secret)
            logging.warning(secret)
            try:
                raise ValueError(secret)
            except ValueError as cause:
                raise AssertionError(secret) from cause

@pytest.fixture
def failed_setup():
    raise WebDriverException(os.environ["YANJIA_AUDIT_SECRET"])

def test_setup(failed_setup):
    pass

@pytest.fixture
def failed_teardown():
    yield
    raise RuntimeError(os.environ["YANJIA_AUDIT_SECRET"])

def test_teardown(failed_teardown):
    pass
'''
    results = tmp_path / "allure"
    junit = tmp_path / "junit.xml"
    environment = os.environ.copy()
    environment["YANJIA_AUDIT_SECRET"] = secret
    environment["PYTHONUTF8"] = "1"
    environment["PYTHONPATH"] = os.pathsep.join((str(root), str(root / "src")))
    case_root = root / ".tmp" / "redaction-subprocess"
    case_root.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix="case-", dir=case_root) as case_directory:
        case_file = Path(case_directory) / "test_private_failures.py"
        case_file.write_text(case_source, encoding="utf-8")
        completed = subprocess.run(
            [
                sys.executable, "-m", "pytest", str(case_file),
                "-p", "tests.excel.conftest", "-p", "no:cacheprovider",
                "--rootdir", str(root),
                "--alluredir", str(results), "--junitxml", str(junit), "--showlocals", "-q",
            ],
            cwd=root, env=environment, capture_output=True, text=True, encoding="utf-8", timeout=60,
        )
    assert completed.returncode == 1, completed.stdout + completed.stderr
    files = list(results.iterdir())
    assert any(path.name.endswith("-attachment.txt") for path in files)
    components = {
        "pytest stdout": completed.stdout,
        "pytest stderr": completed.stderr,
        "JUnit XML": junit.read_text(encoding="utf-8"),
        **{
            f"Allure {path.name}": path.read_text(encoding="utf-8")
            for path in files
        },
    }
    representations = (
        secret, repr(secret)[1:-1], json.dumps(secret)[1:-1],
        json.dumps(secret, ensure_ascii=False)[1:-1],
    )
    for component_name, content in components.items():
        for representation in representations:
            if representation in content:
                pytest.fail(f"sensitive representation leaked through {component_name}")
    records = [
        json.loads(path.read_text(encoding="utf-8")) for path in results.glob("*-result.json")
    ]
    assert sorted(record["status"] for record in records) == ["broken", "broken", "failed"]
    assert any("AssertionError" in record["statusDetails"]["message"] for record in records)
    assert any("WebDriverException" in record["statusDetails"]["message"] for record in records)
