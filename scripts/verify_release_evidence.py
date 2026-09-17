"""Reject empty, incomplete or mismatched dependency SBOM/Trivy evidence."""

from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from pathlib import Path
from urllib.parse import unquote

try:
    from .runtime_dependencies import lock_digest, normalize, read_lock, validate_inventory
except ImportError:
    from runtime_dependencies import lock_digest, normalize, read_lock, validate_inventory


class EvidenceError(ValueError):
    """Release evidence does not establish dependency scan coverage."""


def sbom_packages(sbom: dict) -> set[tuple[str, str]]:
    if sbom.get("spdxVersion") != "SPDX-2.3" or not sbom.get("documentNamespace"):
        raise EvidenceError("expected an SPDX 2.3 document with a namespace")
    packages: set[tuple[str, str]] = set()
    for package in sbom.get("packages", []):
        for ref in package.get("externalRefs", []):
            if ref.get("referenceType") != "purl":
                continue
            match = re.fullmatch(
                r"pkg:pypi/([^/@?#]+)@([^?#]+)(?:[?#].*)?", ref.get("referenceLocator", "")
            )
            if not match:
                continue
            name, version = normalize(unquote(match[1])), unquote(match[2])
            if normalize(package.get("name", "")) != name or package.get("versionInfo") != version:
                raise EvidenceError("SBOM package name/version conflicts with its Python PURL")
            packages.add((name, version))
    if not packages:
        raise EvidenceError("SBOM contains zero Python dependency components")
    return packages


def scan_packages(scan: dict) -> set[tuple[str, str]]:
    if scan.get("SchemaVersion") != 2 or scan.get("ArtifactType") != "spdx":
        raise EvidenceError("expected a Trivy SPDX scan report")
    if not scan.get("Trivy", {}).get("Version"):
        raise EvidenceError("Trivy scanner version is missing")
    results = scan.get("Results")
    if not isinstance(results, list) or not results:
        raise EvidenceError("Trivy report has no scan results")
    packages: set[tuple[str, str]] = set()
    for result in results:
        if result.get("Type") not in {"python-pkg", "pip"}:
            continue
        for package in result.get("Packages", []):
            name, version = package.get("Name"), package.get("Version")
            if not isinstance(name, str) or not name or not isinstance(version, str) or not version:
                raise EvidenceError("Trivy package name/version is missing")
            packages.add((normalize(name), version))
    if not packages:
        raise EvidenceError("Trivy contains zero Python packages; enable --list-all-pkgs")
    return packages


def verify_evidence(lock: Path, inventory: dict, sbom: dict, scan: dict, revision: str) -> dict:
    if not re.fullmatch(r"[0-9a-f]{40}", revision) or inventory.get("source_revision") != revision:
        raise EvidenceError("runtime inventory does not belong to the release revision")
    expected = read_lock(lock)
    validate_inventory(inventory, expected, lock_digest(lock))
    components = sbom_packages(sbom)
    missing = set(expected.items()) - components
    if missing:
        raise EvidenceError(f"SBOM is missing {len(missing)} locked package versions")
    scanned = scan_packages(scan)
    missing = components - scanned
    if missing:
        raise EvidenceError(f"Trivy did not scan {len(missing)} SBOM package versions")
    severities = Counter(
        vuln.get("Severity", "UNKNOWN")
        for result in scan["Results"]
        for vuln in result.get("Vulnerabilities", [])
    )
    if severities["CRITICAL"]:
        raise EvidenceError("critical vulnerabilities block release")
    return {
        "schema_version": "1.0",
        "status": "VERIFIED",
        "source_revision": revision,
        "lock_digest": lock_digest(lock),
        "locked_packages": len(expected),
        "sbom_python_packages": len(components),
        "scanned_python_packages": len(scanned),
        "vulnerabilities_by_severity": dict(sorted(severities.items())),
        "blocking_severities": ["CRITICAL"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("lock", "inventory", "sbom", "scan", "output"):
        parser.add_argument(f"--{name}", required=True, type=Path)
    parser.add_argument("--source-revision", required=True)
    args = parser.parse_args()
    report = verify_evidence(
        args.lock,
        json.loads(args.inventory.read_text("utf-8")),
        json.loads(args.sbom.read_text("utf-8")),
        json.loads(args.scan.read_text("utf-8")),
        args.source_revision,
    )
    report["evidence_digests"] = {
        name: lock_digest(getattr(args, name)) for name in ("inventory", "sbom", "scan")
    }
    args.output.write_text(json.dumps(report, sort_keys=True, indent=2) + "\n", "utf-8")
    print(json.dumps(report, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (ValueError, OSError, TypeError, KeyError) as error:
        print(f"Release dependency evidence rejected: {error}")
        raise SystemExit(2) from None
