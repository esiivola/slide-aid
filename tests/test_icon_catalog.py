from __future__ import annotations

import copy
import collections
import importlib.util
import json
import subprocess
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "build_icon_catalog",
    ROOT / "scripts" / "build_icon_catalog.py",
)
assert SPEC and SPEC.loader
build_icon_catalog = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(build_icon_catalog)


def test_generated_icon_catalog_is_current_and_searchable() -> None:
    catalog = json.loads((ROOT / "shared" / "iconaid" / "catalog.json").read_text())
    assert catalog["schema"] == 2
    assert catalog["viewBox"] == 24
    assert len(catalog["icons"]) == 425
    assert {icon["category"] for icon in catalog["icons"]} == {
        "Business", "Communication", "ESG", "Finance", "Operations", "People", "Security", "Technology",
    }
    category_counts = collections.Counter(icon["category"] for icon in catalog["icons"])
    assert min(category_counts.values()) >= 17
    variant_ids = {
        f"{family_id}-{badge_id}"
        for family_id in build_icon_catalog.FAMILY_IDS
        for badge_id, _name, _terms in build_icon_catalog.VARIANTS
    }
    assert len({icon["id"] for icon in catalog["icons"]} - variant_ids) == 153
    assert len(catalog["benchmarks"]) == 5
    assert catalog["license"].startswith("MIT")
    assert all(len(icon["tags"]) >= 8 and len(icon["aliases"]) >= 2 for icon in catalog["icons"])
    assert all(len(icon["primitives"]) >= 2 for icon in catalog["icons"])
    assert len({icon["id"] for icon in catalog["icons"]}) == len(catalog["icons"])
    geometry = [json.dumps(icon["primitives"], sort_keys=True) for icon in catalog["icons"]]
    assert len(set(geometry)) == len(geometry)
    for path, expected in build_icon_catalog.outputs().items():
        assert path.read_text() == expected
    vba_paths = sorted((ROOT / "apps" / "powerpoint" / "src").glob("modIconAid*.bas"))
    assert len(vba_paths) == 9
    assert max(len(line) for path in vba_paths for line in path.read_text().splitlines()) < 1024
    controller = (ROOT / "apps" / "powerpoint" / "src" / "modIconAid.bas").read_text()
    assert "Open IconAid from Home > Add-ins." in controller
    assert "IA_DrawPreview" not in controller
    assert "Slides.Add" not in controller


def test_consulting_search_vocabulary_finds_expected_concepts() -> None:
    icons = json.loads((ROOT / "shared" / "iconaid" / "catalog.json").read_text())["icons"]

    def ids_matching(query: str) -> set[str]:
        return {
            icon["id"]
            for icon in icons
            if query in " ".join([icon["name"], icon["category"], *icon["aliases"], *icon["tags"]]).lower()
        }

    expectations = {
        "machine learning": "ai-brain",
        "artificial intelligence": "ai-brain",
        "accounts payable": "invoice",
        "org chart": "org-chart",
        "go to market": "rocket",
        "robotic process automation": "automation",
        "cyber security": "firewall",
        "distribution center": "warehouse",
        "pitch deck": "presentation",
        "circular economy": "recycle",
        "balanced scorecard": "strategy-map",
        "initiative portfolio": "portfolio",
        "2x2 matrix": "matrix",
        "business continuity": "disaster-recovery",
        "source to pay": "procurement",
        "talent acquisition": "recruitment",
        "software as a service": "saas",
        "internet of things": "iot",
        "carbon footprint": "emissions",
        "income statement": "profit-loss",
    }
    for query, icon_id in expectations.items():
        assert icon_id in ids_matching(query)


def test_catalog_validation_rejects_duplicate_ids() -> None:
    icons = copy.deepcopy(build_icon_catalog.ICONS)
    icons.append(copy.deepcopy(icons[0]))
    with pytest.raises(ValueError, match="Duplicate icon id"):
        build_icon_catalog.validate_catalog(icons)


def test_catalog_validation_rejects_out_of_bounds_geometry() -> None:
    icons = copy.deepcopy(build_icon_catalog.ICONS)
    icons[0]["primitives"][0]["x1"] = -1
    with pytest.raises(ValueError, match="invalid x1"):
        build_icon_catalog.validate_catalog(icons)


def test_powerpoint_taskpane_manifest_and_vector_adapter() -> None:
    app = ROOT / "apps" / "powerpoint-iconaid"
    manifest = ET.parse(app / "manifest.xml").getroot()
    namespace = {"office": "http://schemas.microsoft.com/office/appforoffice/1.1"}
    host = manifest.find("office:Hosts/office:Host", namespace)
    source = manifest.find("office:DefaultSettings/office:SourceLocation", namespace)
    assert host is not None and host.attrib["Name"] == "Presentation"
    assert source is not None and source.attrib["DefaultValue"].split("?")[0].endswith("/taskpane.html")
    custom_tabs = [element for element in manifest.iter() if element.tag.endswith("CustomTab")]
    tab_labels = [
        element.attrib.get("DefaultValue")
        for element in manifest.iter()
        if element.tag.endswith("String") and element.attrib.get("id") == "IconAid.TabLabel"
    ]
    assert len(custom_tabs) == 1 and custom_tabs[0].attrib["id"] == "IconAid.Tab"
    assert tab_labels == ["IconAid"]

    script = f"""
import {{ createRequire }} from "node:module";
const require = createRequire(import.meta.url);
const assert = require("node:assert/strict");
const api = require({json.dumps(str(app / "taskpane.js"))});
const catalog = require({json.dumps(str(ROOT / "shared" / "iconaid" / "catalog.json"))});
const icon = catalog.icons.find((entry) => entry.id === "disaster-recovery");
assert.equal(api.matchesIcon(icon, "business continuity", ""), true);
assert.equal(api.matchesIcon(icon, "business continuity", "Technology"), false);
const instructions = api.shapeInstructions(icon, "#1f497d");
assert.equal(instructions.length, icon.primitives.length);
assert.ok(instructions.some((item) => item.kind === "line"));
assert.ok(instructions.every((item) => Number.isFinite(item.left) && Number.isFinite(item.top)));
assert.ok(instructions.filter((item) => item.kind === "line").every((item) => item.width > 0 && item.height > 0));
assert.ok(instructions.filter((item) => item.kind === "line").every((item) => Number.isFinite(item.rotation)));
const calls = [];
let shapeNumber = 0;
function shape(type, options) {{
  return {{
    id: `shape-${{++shapeNumber}}`,
    fill: {{
      clear: () => calls.push(["clear", type]),
      setSolidColor: (color) => calls.push(["fill", type, color]),
    }},
    lineFormat: {{}},
    type,
    options,
  }};
}}
const group = {{}};
const shapes = {{
  addGeometricShape: (type, options) => {{
    const item = shape(type, options);
    calls.push(["shape", type, options, item]);
    return item;
  }},
  addGroup: (items) => (calls.push(["group", items]), group),
}};
global.PowerPoint = {{
  GeometricShapeType: {{ rectangle: "Rectangle", ellipse: "Ellipse" }},
  run: async (callback) => callback({{
    presentation: {{ getSelectedSlides: () => ({{ getItemAt: () => ({{ shapes }}) }}) }},
    sync: async () => calls.push(["sync"]),
  }}),
}};
await api.insertVectorIcon(icon, "#1f497d", catalog.viewBox);
const firstSync = calls.findIndex((entry) => entry[0] === "sync");
const groupIndex = calls.findIndex((entry) => entry[0] === "group");
const shapeCalls = calls.filter((entry) => entry[0] === "shape");
assert.equal(shapeCalls.length, instructions.length);
assert.ok(firstSync >= 0 && firstSync < groupIndex);
assert.equal(calls.filter((entry) => entry[0] === "sync").length, 2);
assert.equal(calls.filter((entry) => entry[0] === "group").length, 1);
assert.ok(shapeCalls.some((entry) => Number.isFinite(entry[3].rotation)));
assert.equal(group.name, "IconAid - Disaster Recovery");
assert.equal(group.altTextDescription, "IconAid vector icon: Disaster Recovery");
"""
    subprocess.run(["node", "--input-type=module", "-e", script], check=True, cwd=ROOT)
