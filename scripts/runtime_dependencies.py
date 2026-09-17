"""Verify the signed Windows Python runtime lock before touching a device."""

from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import json
import platform
import re
import subprocess
import sys
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
LOCK = ROOT / "config/runtime-requirements.lock"
PIN = re.compile(r"([a-z0-9][a-z0-9.-]*)==([a-zA-Z0-9.!+_-]+) --hash=sha256:([a-f0-9]{64})")


class RuntimeDependencyError(ValueError):
    """The installed environment cannot be matched to the signed lock."""


def normalize(name: str) -> str:
    return re.sub(r"[-_.]+", "-", name).lower()


def lock_digest(path: Path) -> str:
    return "sha256:" + hashlib.sha256(path.read_bytes()).hexdigest()


def read_lock(path: Path) -> dict[str, str]:
    packages: dict[str, str] = {}
    for line in path.read_text("utf-8").splitlines():
        if not line.strip() or line.startswith("#"):
            continue
        match = PIN.fullmatch(line)
        if match is None:
            raise RuntimeDependencyError(
                "lock must contain exact versions and SHA-256 wheel hashes"
            )
        name, version, _ = match.groups()
        if normalize(name) != name or name in packages:
            raise RuntimeDependencyError("lock contains a duplicate or non-canonical package name")
        packages[name] = version
    if not packages:
        raise RuntimeDependencyError("runtime lock is empty")
    return packages


def validate_inventory(inventory: dict, expected: dict[str, str], digest: str) -> None:
    if inventory.get("schema_version") != "1.0" or inventory.get("lock_digest") != digest:
        raise RuntimeDependencyError("runtime inventory does not match the signed lock")
    runtime = inventory.get("runtime", {})
    if (
        runtime.get("platform") != "win32"
        or runtime.get("python_minor") != "3.12"
        or runtime.get("implementation") != "cpython"
        or str(runtime.get("machine", "")).lower() not in {"amd64", "x86_64"}
    ):
        raise RuntimeDependencyError("runtime requires Windows x64 CPython 3.12")
    actual: dict[str, str] = {}
    for item in inventory.get("packages", []):
        name, version = item.get("name"), item.get("version")
        if not isinstance(name, str) or not isinstance(version, str):
            raise RuntimeDependencyError("invalid runtime package inventory")
        name = normalize(name)
        if name in actual:
            raise RuntimeDependencyError("duplicate installed distribution")
        actual[name] = version
    if actual != expected:
        missing = len(expected.keys() - actual.keys())
        extra = len(actual.keys() - expected.keys())
        drift = sum(
            actual.get(name) != version for name, version in expected.items() if name in actual
        )
        raise RuntimeDependencyError(
            f"runtime dependency drift: missing={missing}, unexpected={extra}, "
            f"version_mismatch={drift}"
        )


def verify_environment(lock: Path, project_file: Path) -> dict:
    expected = read_lock(lock)
    inventory = {
        "schema_version": "1.0",
        "lock_digest": lock_digest(lock),
        "runtime": {
            "platform": sys.platform,
            "python_minor": f"{sys.version_info.major}.{sys.version_info.minor}",
            "python_version": platform.python_version(),
            "implementation": sys.implementation.name,
            "machine": platform.machine(),
        },
        "packages": sorted(
            (
                {"name": normalize(d.metadata["Name"]), "version": d.version}
                for d in importlib.metadata.distributions()
            ),
            key=lambda item: item["name"],
        ),
    }
    validate_inventory(inventory, expected, lock_digest(lock))
    # Import only after exact distribution/version verification.
    from packaging.requirements import Requirement

    project = tomllib.loads(project_file.read_text("utf-8"))["project"]
    for value in project["dependencies"]:
        requirement = Requirement(value)
        if requirement.marker is not None and not requirement.marker.evaluate():
            continue
        pinned = expected.get(normalize(requirement.name))
        if requirement.url or pinned is None or not requirement.specifier.contains(pinned):
            raise RuntimeDependencyError("project runtime requirement does not match the lock")
    check = subprocess.run(
        [sys.executable, "-I", "-m", "pip", "check"],
        capture_output=True,
        timeout=60,
        check=False,
    )
    if check.returncode:
        raise RuntimeDependencyError("installed dependency graph failed pip check")
    return inventory


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--lock", type=Path, default=LOCK)
    parser.add_argument("--project-file", type=Path, default=ROOT / "pyproject.toml")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--source-revision")
    args = parser.parse_args()
    if args.source_revision and not re.fullmatch(r"[0-9a-f]{40}", args.source_revision):
        raise RuntimeDependencyError("source revision must be a full immutable Git SHA")
    inventory = verify_environment(args.lock, args.project_file)
    if args.source_revision:
        inventory["source_revision"] = args.source_revision
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(inventory, sort_keys=True, indent=2) + "\n", "utf-8")
    print(
        json.dumps(
            {
                "runtime_dependencies": "VERIFIED",
                "package_count": len(inventory["packages"]),
                "lock_digest": inventory["lock_digest"],
            }
        )
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (RuntimeDependencyError, OSError, ValueError, subprocess.SubprocessError) as error:
        # Files, pip output and environment values must never be echoed here.
        detail = str(error) if isinstance(error, RuntimeDependencyError) else type(error).__name__
        print(f"Runtime dependency verification failed: {detail}", file=sys.stderr)
        raise SystemExit(2) from None
