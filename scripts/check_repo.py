#!/usr/bin/env python3
"""Run repo-level validation for the Slide Aid monorepo."""
from __future__ import annotations

import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parent.parent
POWERPOINT = ROOT / "apps" / "powerpoint"
GOOGLE = ROOT / "apps" / "google-slides"
ICONS = ROOT / "shared" / "icons"


def run(cmd: list[str], cwd: Path = ROOT, allow_pytest_no_tests: bool = False) -> None:
    print(f"$ {' '.join(cmd)}", flush=True)
    proc = subprocess.run(cmd, cwd=cwd)
    if proc.returncode == 0:
        return
    if allow_pytest_no_tests and proc.returncode == 5:
        print("pytest collected no tests; continuing.", flush=True)
        return
    raise SystemExit(proc.returncode)


def check_json_specs() -> None:
    for path in sorted((ROOT / "shared" / "specs").glob("*.json")):
        json.loads(path.read_text())
    print("shared specs: ok")


def check_ribbon_icons() -> None:
    ribbon_xml = POWERPOINT / "ribbon" / "customUI14.xml"
    ET.fromstring(ribbon_xml.read_bytes())
    refs = set(re.findall(r'\bimage="([^"]+)"', ribbon_xml.read_text()))
    missing = sorted(ref for ref in refs if not (ICONS / f"{ref}.png").exists())
    if missing:
        raise SystemExit("missing ribbon icons: " + ", ".join(missing))
    print(f"ribbon icons: {len(refs)} refs ok")


def check_docs_images() -> None:
    refs: list[tuple[Path, str]] = []
    for path in sorted(ROOT.rglob("*.md")):
        if ".git" in path.parts or "node_modules" in path.parts:
            continue
        text = path.read_text()
        for ref in re.findall(r'!\[[^\]]*\]\(([^)]+)\)', text):
            refs.append((path, ref))
        for ref in re.findall(r'<img\s+[^>]*src="([^"]+)"', text):
            refs.append((path, ref))
    missing = []
    for path, ref in refs:
        if re.match(r"^[a-z]+://", ref):
            continue
        if not (path.parent / ref).resolve().exists():
            missing.append(f"{path.relative_to(ROOT)} -> {ref}")
    missing = sorted(missing)
    if missing:
        raise SystemExit("missing documentation images: " + ", ".join(missing))
    print(f"documentation images: {len(refs)} refs ok")


def check_gifs() -> None:
    try:
        from PIL import Image
    except ImportError:
        print("Pillow not installed; skipping GIF decode check.")
        return
    count = 0
    frames = 0
    for gif in sorted((ROOT / "docs" / "img").glob("*.gif")):
        image = Image.open(gif)
        n_frames = getattr(image, "n_frames", 1)
        for frame in range(n_frames):
            image.seek(frame)
            image.convert("RGB").load()
        count += 1
        frames += n_frames
    print(f"gifs: {count} files, {frames} frames ok")


def check_powerpoint_release() -> None:
    addin = POWERPOINT / "dist" / "Slide Aid.ppam"
    package = POWERPOINT / "dist" / "Slide Aid.zip"
    with zipfile.ZipFile(addin) as archive:
        names = set(archive.namelist())
        required = {"ppt/vbaProject.bin", "customUI/customUI14.xml"}
        if not required.issubset(names):
            raise SystemExit("PowerPoint add-in is missing VBA or ribbon content")
        icon_count = sum(name.startswith("customUI/images/") for name in names)
        expected_icon_count = sum(1 for _ in ICONS.glob("*.png"))
        if icon_count != expected_icon_count:
            raise SystemExit(f"PowerPoint add-in has {icon_count} icons; expected {expected_icon_count}")

    with zipfile.ZipFile(package) as archive:
        names = archive.namelist()
        if any(Path(name).name.startswith("._") or name.startswith("__MACOSX/") for name in names):
            raise SystemExit("PowerPoint installer contains macOS metadata sidecars")
        packaged_addin = archive.read("Slide Aid/Slide Aid.ppam")
    if packaged_addin != addin.read_bytes():
        raise SystemExit("PowerPoint installer add-in differs from dist/Slide Aid.ppam")
    print(f"PowerPoint release: VBA, ribbon, {icon_count} icons and installer copy ok")


def main() -> None:
    run([sys.executable, "-m", "py_compile", "scripts/make_icons.py", "scripts/build_icon_catalog.py", "scripts/workflow_contracts.py", "scripts/render_doc_gifs.py", "scripts/render_examples.py", "scripts/check_repo.py"])
    run([sys.executable, "scripts/build_icon_catalog.py", "--check"])
    run([sys.executable, "-m", "py_compile", "apps/powerpoint/tools/inject_ribbon.py"])
    run(["bash", "-n", "apps/powerpoint/tools/build.sh", "apps/powerpoint/tools/make_dist.sh"])
    check_json_specs()
    check_ribbon_icons()
    check_docs_images()
    check_gifs()
    check_powerpoint_release()
    run(["npm", "run", "check"], cwd=GOOGLE)
    run(["npm", "test"], cwd=GOOGLE)
    run(["npm", "run", "build"], cwd=GOOGLE)
    run(["pytest"], allow_pytest_no_tests=True)


if __name__ == "__main__":
    main()
