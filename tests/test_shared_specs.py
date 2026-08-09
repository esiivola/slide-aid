"""The shared/specs files are the reviewable contract between the PowerPoint
add-in and the Google Slides companion.

Nothing enforced them, so drift was silent: the two products could disagree about
a chart parameter's default or a palette's colors and the only symptom would be
the same deck rendering differently on each platform. These tests pin the spec to
the VBA on one side; the Google Slides suite pins the same files to the
TypeScript on the other (apps/google-slides/tests/core.test.ts).
"""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SPECS = ROOT / "shared" / "specs"
CHART_STYLE_BAS = ROOT / "apps" / "powerpoint" / "src" / "modChartStyle.bas"


def _vba_block(source: str, function_name: str) -> str:
    match = re.search(rf"(Private|Public) Function {function_name}\(.*?End Function", source, re.S)
    assert match, f"{function_name} not found in modChartStyle.bas"
    return match.group(0)


def test_chart_style_spec_matches_powerpoint_keydefs() -> None:
    block = _vba_block(CHART_STYLE_BAS.read_text(), "KeyDefs")
    entries = [
        {"key": key, "default": default, "description": description}
        for key, default, description in re.findall(r'Array\("([^"]+)",\s*"([^"]*)",\s*"([^"]*)"\)', block)
    ]
    spec = json.loads((SPECS / "chart-style.json").read_text())["keys"]
    assert entries, "no KeyDefs entries parsed"
    assert entries == spec, "modChartStyle.bas KeyDefs and shared/specs/chart-style.json disagree"


def test_palette_spec_matches_powerpoint_themes() -> None:
    block = _vba_block(CHART_STYLE_BAS.read_text(), "ThemeDef")
    themes = {
        name: [f"#{value.upper()}" for value in colors.split('", "')]
        for name, colors in re.findall(r'ThemeDef = Array\("([^"]+)",\s*"([^)]+?)"\)', block)
    }
    spec = json.loads((SPECS / "palettes.json").read_text())
    assert len(themes) == 9, f"expected 9 color themes, parsed {len(themes)}"
    assert themes == spec, "modChartStyle.bas ThemeDef and shared/specs/palettes.json disagree"
    for name, colors in spec.items():
        assert len(colors) == 6, f"{name} must have six colors"
        for color in colors:
            assert re.fullmatch(r"#[0-9A-F]{6}", color), f"{name} has a malformed color {color}"


def test_chart_kinds_spec_matches_the_ribbon() -> None:
    ribbon = (ROOT / "apps" / "powerpoint" / "ribbon" / "customUI14.xml").read_text()
    # Every Ch:<KIND> button on the Chart Aid tab must be a declared kind.
    ribbon_kinds = set(re.findall(r'tag="Ch:([A-Z]+)"', ribbon))
    spec = set(json.loads((SPECS / "chart-kinds.json").read_text())["chartKinds"])
    assert ribbon_kinds, "no chart buttons found in the ribbon"
    missing = ribbon_kinds - spec
    assert not missing, f"ribbon builds chart kinds missing from the spec: {sorted(missing)}"
    # BUB has no ribbon button of its own (Scatter builds it from a fourth
    # column), so the spec is allowed to be the larger set.
    assert spec - ribbon_kinds <= {"BUB"}, f"spec declares kinds the ribbon cannot build: {sorted(spec - ribbon_kinds - {'BUB'})}"


def test_specs_readme_lists_every_spec_file() -> None:
    readme = (SPECS / "README.md").read_text()
    for path in sorted(SPECS.glob("*.json")):
        assert path.name in readme, f"{path.name} is not documented in shared/specs/README.md"
