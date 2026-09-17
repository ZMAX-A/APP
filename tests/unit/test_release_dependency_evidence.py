from __future__ import annotations

import copy
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

from scripts.runtime_dependencies import (
    RuntimeDependencyError,
    lock_digest,
    read_lock,
    validate_inventory,
)
from scripts.verify_release_evidence import EvidenceError, verify_evidence

REVISION = "a" * 40


@pytest.fixture
def evidence(tmp_path: Path) -> tuple:
    lock = tmp_path / "runtime.lock"
    lock.write_text(f"demo==1.2.3 --hash=sha256:{'0' * 64}\n", encoding="utf-8")
    inventory = {
        "schema_version": "1.0",
        "source_revision": REVISION,
        "lock_digest": lock_digest(lock),
        "runtime": {
            "platform": "win32",
            "python_minor": "3.12",
            "implementation": "cpython",
            "machine": "AMD64",
        },
        "packages": [{"name": "demo", "version": "1.2.3"}],
    }
    sbom = {
        "spdxVersion": "SPDX-2.3",
        "documentNamespace": "https://example.test/sbom",
        "packages": [
            {
                "name": "demo",
                "versionInfo": "1.2.3",
                "externalRefs": [
                    {"referenceType": "purl", "referenceLocator": "pkg:pypi/demo@1.2.3"}
                ],
            }
        ],
    }
    scan = {
        "SchemaVersion": 2,
        "ArtifactType": "spdx",
        "Trivy": {"Version": "0.70.0"},
        "Results": [{"Type": "python-pkg", "Packages": [{"Name": "demo", "Version": "1.2.3"}]}],
    }
    return lock, inventory, sbom, scan, REVISION


def test_zero_vulnerabilities_with_complete_scan_coverage_passes(evidence: tuple) -> None:
    report = verify_evidence(*evidence)
    assert report["status"] == "VERIFIED"
    assert report["vulnerabilities_by_severity"] == {}


def test_directory_only_sbom_from_failed_release_is_rejected(evidence: tuple) -> None:
    evidence[2]["packages"] = [{"name": "bundle", "primaryPackagePurpose": "FILE"}]
    with pytest.raises(EvidenceError, match="zero Python"):
        verify_evidence(*evidence)


@pytest.mark.parametrize("results", [None, [], [{}], [{"Type": "python-pkg"}]])
def test_empty_scan_cannot_pass_as_zero_vulnerabilities(evidence: tuple, results) -> None:
    evidence[3]["Results"] = results
    with pytest.raises(EvidenceError):
        verify_evidence(*evidence)


def test_sbom_must_cover_exact_locked_version(evidence: tuple) -> None:
    package = evidence[2]["packages"][0]
    package["versionInfo"] = "1.2.2"
    package["externalRefs"][0]["referenceLocator"] = "pkg:pypi/demo@1.2.2"
    with pytest.raises(EvidenceError, match="missing 1 locked"):
        verify_evidence(*evidence)


def test_trivy_must_cover_exact_sbom_version(evidence: tuple) -> None:
    evidence[3]["Results"][0]["Packages"][0]["Version"] = "1.2.2"
    with pytest.raises(EvidenceError, match="did not scan"):
        verify_evidence(*evidence)


def test_python_purl_cannot_disagree_with_sbom_version(evidence: tuple) -> None:
    evidence[2]["packages"][0]["versionInfo"] = "9.9"
    with pytest.raises(EvidenceError, match="conflicts"):
        verify_evidence(*evidence)


def test_vendored_sbom_components_also_require_scan_coverage(evidence: tuple) -> None:
    vendored = copy.deepcopy(evidence[2]["packages"][0])
    vendored["versionInfo"] = "1.0"
    vendored["externalRefs"][0]["referenceLocator"] = "pkg:pypi/demo@1.0"
    evidence[2]["packages"].append(vendored)
    with pytest.raises(EvidenceError, match="did not scan"):
        verify_evidence(*evidence)


def test_source_revision_mismatch_is_rejected(evidence: tuple) -> None:
    evidence[1]["source_revision"] = "b" * 40
    with pytest.raises(EvidenceError, match="revision"):
        verify_evidence(*evidence)


@pytest.mark.parametrize("severity", ["CRITICAL", "HIGH", "MEDIUM", "LOW", "UNKNOWN"])
def test_all_severities_are_recorded_but_critical_blocks(evidence: tuple, severity: str) -> None:
    evidence[3]["Results"][0]["Vulnerabilities"] = [{"Severity": severity}]
    if severity == "CRITICAL":
        with pytest.raises(EvidenceError, match="critical"):
            verify_evidence(*evidence)
    else:
        assert verify_evidence(*evidence)["vulnerabilities_by_severity"] == {severity: 1}


@pytest.mark.parametrize(
    "change", ["missing", "extra", "version", "duplicate", "platform", "python", "lock"]
)
def test_worker_rejects_runtime_drift_before_execution(evidence: tuple, change: str) -> None:
    lock, inventory, *_ = evidence
    if change == "missing":
        inventory["packages"] = []
    elif change == "extra":
        inventory["packages"].append({"name": "unscanned-plugin", "version": "1.0"})
    elif change == "version":
        inventory["packages"][0]["version"] = "9.0"
    elif change == "duplicate":
        inventory["packages"].append(inventory["packages"][0])
    elif change == "platform":
        inventory["runtime"]["platform"] = "linux"
    elif change == "python":
        inventory["runtime"]["python_minor"] = "3.13"
    else:
        inventory["lock_digest"] = "sha256:" + "1" * 64
    with pytest.raises(RuntimeDependencyError):
        validate_inventory(inventory, read_lock(lock), lock_digest(lock))


@pytest.mark.parametrize(
    "content", ["", "demo>=1.0", "demo==1.0", "--index-url https://example.test"]
)
def test_unlocked_or_option_injected_requirements_are_rejected(
    tmp_path: Path, content: str
) -> None:
    path = tmp_path / "invalid.lock"
    path.write_text(content, encoding="utf-8")
    with pytest.raises(RuntimeDependencyError):
        read_lock(path)


def test_cli_rejects_empty_evidence_without_writing_pass_report(
    evidence: tuple, tmp_path: Path
) -> None:
    lock, inventory, sbom, scan, revision = evidence
    scan.pop("Results")
    paths = {}
    for name, value in (("inventory", inventory), ("sbom", sbom), ("scan", scan)):
        paths[name] = tmp_path / f"{name}.json"
        paths[name].write_text(json.dumps(value), encoding="utf-8")
    output = tmp_path / "coverage.json"
    script = Path(__file__).resolve().parents[2] / "scripts/verify_release_evidence.py"
    result = subprocess.run(
        [
            sys.executable,
            str(script),
            "--lock",
            str(lock),
            "--inventory",
            str(paths["inventory"]),
            "--sbom",
            str(paths["sbom"]),
            "--scan",
            str(paths["scan"]),
            "--source-revision",
            revision,
            "--output",
            str(output),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode != 0
    assert not output.exists()


@pytest.mark.skipif(sys.platform != "win32", reason="Windows launcher integration")
def test_top_launcher_cannot_skip_dependency_gate(tmp_path: Path) -> None:
    shell = shutil.which("pwsh") or shutil.which("powershell")
    if not shell:
        pytest.skip("PowerShell is unavailable")
    root = Path(__file__).resolve().parents[2]
    scripts = tmp_path / "scripts"
    scripts.mkdir()
    source = (root / "scripts/run-excel.ps1").read_text("utf-8-sig")
    source = source.replace(
        "$python = Join-Path $projectRoot '.venv\\Scripts\\python.exe'",
        "$python = '" + sys.executable.replace("'", "''") + "'",
    )
    launcher = scripts / "run-excel.ps1"
    launcher.write_text(source, encoding="utf-8-sig")
    (scripts / "runtime_dependencies.py").write_text("raise SystemExit(2)\n", encoding="utf-8")
    result = subprocess.run(
        [
            shell,
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(launcher),
            "-IgnoreDotEnv",
            "-SkipPreflight",
            "-SkipExcelValidation",
            "-NoWriteBack",
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    assert result.returncode != 0
    assert "TOP runtime dependency verification failed" in result.stdout + result.stderr
    assert not (tmp_path / "reports").exists()
