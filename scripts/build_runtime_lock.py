"""Create a reviewed Windows runtime lock from a resolved wheel-only download."""

from __future__ import annotations

import argparse
import hashlib
import re
import zipfile
from email.parser import BytesParser
from pathlib import Path


def build_lock(wheelhouse: Path, output: Path) -> int:
    lines: dict[str, str] = {}
    for wheel in sorted(wheelhouse.glob("*.whl")):
        with zipfile.ZipFile(wheel) as archive:
            entries = [name for name in archive.namelist() if name.endswith(".dist-info/METADATA")]
            if len(entries) != 1:
                raise ValueError("wheel must contain one distribution metadata file")
            metadata = BytesParser().parsebytes(archive.read(entries[0]))
        name = re.sub(r"[-_.]+", "-", str(metadata["Name"])).lower()
        version = str(metadata["Version"])
        if (
            not re.fullmatch(r"[a-z0-9][a-z0-9.-]*", name)
            or not re.fullmatch(r"[a-zA-Z0-9.!+_-]+", version)
            or name in lines
        ):
            raise ValueError("invalid or duplicate wheel distribution")
        digest = hashlib.sha256(wheel.read_bytes()).hexdigest()
        lines[name] = f"{name}=={version} --hash=sha256:{digest}"
    if not lines:
        raise ValueError("wheelhouse is empty")
    output.write_text(
        "# Windows x64 CPython 3.12 runtime, including pip; wheel SHA-256 hashes.\n"
        "# Generated from a clean resolved wheelhouse; review and commit dependency changes.\n"
        + "\n".join(lines[name] for name in sorted(lines))
        + "\n",
        encoding="utf-8",
        newline="\n",
    )
    return len(lines)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--wheelhouse", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    print(f"Locked {build_lock(args.wheelhouse, args.output)} wheel distributions.")
