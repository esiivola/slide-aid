from __future__ import annotations

import ast
import importlib.util
import math
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "workflow_contracts",
    ROOT / "scripts" / "workflow_contracts.py",
)
assert SPEC and SPEC.loader
workflow_contracts = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(workflow_contracts)


def test_all_documentation_workflows_have_valid_contracts() -> None:
    tree = ast.parse((ROOT / "scripts" / "render_doc_gifs.py").read_text())
    demos = [
        node
        for node in ast.walk(tree)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "Demo"
    ]
    names = [ast.literal_eval(node.args[0]) for node in demos]
    assert len(names) == 42
    assert len(names) == len(set(names))
    assert all(any(keyword.arg == "contract" for keyword in node.keywords) for node in demos)


def test_golden_canon_uses_two_to_one_vertical_margins() -> None:
    top = workflow_contracts.golden_top(165, 265, 78)
    top_margin = top - 165
    bottom_margin = 165 + 265 - (top + 78)
    assert math.isclose(bottom_margin, 2 * top_margin)


def test_magic_resizer_scales_position_and_size_about_selection_center() -> None:
    scaled = workflow_contracts.scale_box_about((270, 223, 375, 281), (420, 288), 1.2)
    assert scaled == (240.0, 210.0, 366.0, 279.6)
    assert scaled[2] - scaled[0] == 126.0
    assert math.isclose(scaled[3] - scaled[1], 69.6)


def test_magic_resizer_vba_moves_centers_relative_to_selection_center() -> None:
    source = (ROOT / "apps" / "powerpoint" / "src" / "modSizeAngle.bas").read_text()
    assert "blockCx = sr.Left + sr.Width / 2" in source
    assert "newCx = blockCx + (cx - blockCx) * f" in source
    assert "s.Left = newCx - s.Width / 2" in source


def test_powerpoint_release_scripts_preserve_the_dist_addin() -> None:
    helper = (ROOT / "apps" / "powerpoint" / "tools" / "SlideAidUI.applescript").read_text()
    importer = (ROOT / "apps" / "powerpoint" / "tools" / "import_helper.bas").read_text()
    builder = (ROOT / "apps" / "powerpoint" / "tools" / "build.sh").read_text()
    packager = (ROOT / "apps" / "powerpoint" / "tools" / "make_dist.sh").read_text()

    assert 'repoPath & "/dist/Slide Aid.ppam"' in helper
    assert 'BuildDir = Environ("HOME") & "/SlideAid/build"' in importer
    assert 'sourceDir = BuildDir() & "/src"' in importer
    assert 'repo & vbLf & BuildPptmPath()' in importer
    assert "/Users/" not in importer
    assert 'cp src/*.bas "$SOURCE_CACHE/"' in builder
    assert 'cp tools/import_helper.bas "$SOURCE_CACHE/modImportHelper.bas"' in builder
    assert 'ADDIN="dist/Slide Aid.ppam"' in packager
    assert 'rm -rf dist\n' not in packager
    assert "--norsrc" in packager
    assert 'cp "$ADDIN" hammerspoon/slideaid.lua "$DIST/"' in packager
