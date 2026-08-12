from __future__ import annotations

import copy
import importlib.util
import json
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
    assert catalog["schema"] == 3
    assert catalog["viewBox"] == 24
    pilot_ids = {
        "strategy-target", "analytics", "transformation", "growth", "cost-reduction",
        "organization-team", "customer", "market", "cloud", "database", "ai",
        "cybersecurity", "process", "supply-chain", "factory", "finance", "risk",
        "sustainability", "document", "communication", "roadmap", "portfolio",
        "matrix", "decision-tree", "milestone", "performance-gauge",
        "market-outlook", "ambition", "value", "priority", "capital",
        "cash-flow", "budget", "forecast", "profit-loss", "investment",
        "tax", "treasury-security", "pricing", "margin", "saas",
        "microservices", "data-pipeline", "ai-agent", "model", "code-branch",
        "web-app", "sensor", "digital-twin", "integration", "firewall",
        "key", "certificate", "compliance", "access-control", "resilience",
        "incident", "privacy", "audit", "security-control", "route",
        "location", "logistics", "warehouse", "quality", "maintenance",
        "inventory", "procurement", "service-operations", "capacity",
    }
    assert {icon["id"] for icon in catalog["icons"]} == pilot_ids
    assert len(catalog["benchmarks"]) >= 10
    assert catalog["reviewPolicy"]["mechanicalVariants"] == "disabled"
    assert catalog["license"].startswith("MIT")
    assert all(len(icon["tags"]) >= 8 and len(icon["aliases"]) >= 2 for icon in catalog["icons"])
    assert all(len(icon["primitives"]) >= 2 for icon in catalog["icons"])
    assert all(icon["reviewStatus"] == "pilot" for icon in catalog["icons"])
    assert all(icon.get("elements") for icon in catalog["icons"])
    assert any(any(element["kind"] == "path" for element in icon["elements"]) for icon in catalog["icons"])
    assert len({icon["id"] for icon in catalog["icons"]}) == len(catalog["icons"])
    geometry = [json.dumps(icon["elements"], sort_keys=True) for icon in catalog["icons"]]
    assert len(set(geometry)) == len(geometry)
    variant_suffixes = {f"-{badge_id}" for badge_id, _name, _terms in build_icon_catalog.VARIANTS}
    assert not any(icon["id"].endswith(suffix) for icon in catalog["icons"] for suffix in variant_suffixes)
    for doc in ("DESIGN_SYSTEM.md", "BENCHMARKS.md", "LICENSES.md"):
        text = (ROOT / "shared" / "iconaid" / doc).read_text()
        assert "copy" in text.lower() or "grid" in text.lower() or "license" in text.lower()
    contact_sheet = (ROOT / "shared" / "iconaid" / "contact-sheets" / "pilot.svg").read_text()
    latest_sheet = (ROOT / "shared" / "iconaid" / "contact-sheets" / "latest-batch.svg").read_text()
    assert "Current" in contact_sheet and "Benchmark cues" in contact_sheet and "72pt dark" in contact_sheet
    assert "IconAid operations batch contact sheet" in latest_sheet and "Service Operations" in latest_sheet
    for path, expected in build_icon_catalog.outputs().items():
        assert path.read_text() == expected
    # The legacy per-icon VBA modules (modIconAid*.bas) were replaced by the
    # file-backed loader + external icons.dat, so the generator no longer emits
    # them and they must stay off disk.
    assert not list((ROOT / "apps" / "powerpoint" / "src").glob("modIconAid[!C]*.bas"))
    # IconAid VBA is now a single file-backed loader (modIconAidCurated.bas); the
    # per-icon data modules (modIconAidData*.bas) were replaced by icons.dat.
    src_dir = ROOT / "apps" / "powerpoint" / "src"
    assert not list(src_dir.glob("modIconAidData*.bas"))
    loader = (src_dir / "modIconAidCurated.bas").read_text()
    assert "icons.dat" in loader
    assert "MakeIconsEditable" in loader
    assert "LockAspectRatio = msoTrue" in loader   # editable icons keep proportion
    # Mac builds can expose Merge Shapes only through the native command, not
    # ShapeRange.MergeShapes. Without this fallback, filled counters/holes become
    # opaque because the contours are merely grouped on top of one another.
    assert 'Application.CommandBars.ExecuteMso "ShapesCombine"' in loader
    assert "beforeMergeCount - grp.Count + 1" in loader


def test_consulting_search_vocabulary_finds_expected_concepts() -> None:
    icons = json.loads((ROOT / "shared" / "iconaid" / "catalog.json").read_text())["icons"]

    def ids_matching(query: str) -> set[str]:
        return {
            icon["id"]
            for icon in icons
            if query in " ".join([icon["name"], icon["category"], *icon["aliases"], *icon["tags"]]).lower()
        }

    expectations = {
        "machine learning": "ai",
        "artificial intelligence": "ai",
        "strategic objective": "strategy-target",
        "north star": "strategy-target",
        "business transformation": "transformation",
        "operating model": "transformation",
        "revenue growth": "growth",
        "cost out": "cost-reduction",
        "margin improvement": "cost-reduction",
        "org": "organization-team",
        "voice of customer": "customer",
        "market opportunity": "market",
        "software as a service": "cloud",
        "lakehouse": "database",
        "cyber security": "cybersecurity",
        "value chain": "supply-chain",
        "manufacturing plant": "factory",
        "strategic roadmap": "roadmap",
        "initiative portfolio": "portfolio",
        "quadrant chart": "matrix",
        "decision logic": "decision-tree",
        "stage gate": "milestone",
        "kpi gauge": "performance-gauge",
        "opportunity horizon": "market-outlook",
        "bold goal": "ambition",
        "value proposition": "value",
        "critical focus": "priority",
        "financial capital": "capital",
        "working capital movement": "cash-flow",
        "spending plan": "budget",
        "financial projection": "forecast",
        "income statement": "profit-loss",
        "capital investment": "investment",
        "tax rate": "tax",
        "treasury protection": "treasury-security",
        "pricing strategy": "pricing",
        "profit margin": "margin",
        "subscription software": "saas",
        "service architecture": "microservices",
        "etl": "data-pipeline",
        "automation agent": "ai-agent",
        "predictive model": "model",
        "version control": "code-branch",
        "browser app": "web-app",
        "telemetry device": "sensor",
        "virtual replica": "digital-twin",
        "system integration": "integration",
        "perimeter defense": "firewall",
        "credential key": "key",
        "verified certificate": "certificate",
        "policy check": "compliance",
        "role access": "access-control",
        "cyber resilience": "resilience",
        "breach event": "incident",
        "data privacy": "privacy",
        "assurance review": "audit",
        "control framework": "security-control",
        "delivery route": "route",
        "facility site": "location",
        "freight logistics": "logistics",
        "fulfillment center": "warehouse",
        "inspection standard": "quality",
        "asset maintenance": "maintenance",
        "stock management": "inventory",
        "strategic sourcing": "procurement",
        "service desk": "service-operations",
        "capacity planning": "capacity",
        "financial performance": "finance",
        "risk exposure": "risk",
        "decarbonization": "sustainability",
        "stakeholder communication": "communication",
    }
    for query, icon_id in expectations.items():
        assert icon_id in ids_matching(query)


def test_catalog_validation_rejects_duplicate_ids() -> None:
    icons = copy.deepcopy(build_icon_catalog.ACTIVE_ICONS)
    icons.append(copy.deepcopy(icons[0]))
    with pytest.raises(ValueError, match="Duplicate icon id"):
        build_icon_catalog.validate_catalog(icons)


def test_catalog_validation_rejects_out_of_bounds_geometry() -> None:
    icons = copy.deepcopy(build_icon_catalog.ACTIVE_ICONS)
    icons[0]["primitives"][0]["x1"] = -1
    with pytest.raises(ValueError, match="invalid x1"):
        build_icon_catalog.validate_catalog(icons)


def test_catalog_validation_rejects_out_of_bounds_path_geometry() -> None:
    icons = copy.deepcopy(build_icon_catalog.ACTIVE_ICONS)
    icons[0]["elements"] = [{"kind": "path", "d": "M0 0 L25 1"}]
    with pytest.raises(ValueError, match="path coordinate"):
        build_icon_catalog.validate_catalog(icons)


def test_powerpoint_taskpane_manifest_targets_the_insert_tab() -> None:
    app = ROOT / "apps" / "powerpoint-iconaid"
    manifest = ET.parse(app / "manifest.xml").getroot()
    namespace = {"office": "http://schemas.microsoft.com/office/appforoffice/1.1"}
    host = manifest.find("office:Hosts/office:Host", namespace)
    source = manifest.find("office:DefaultSettings/office:SourceLocation", namespace)
    assert host is not None and host.attrib["Name"] == "Presentation"
    assert source is not None and source.attrib["DefaultValue"].split("?")[0].endswith("/taskpane.html")
    manifest_text = (app / "manifest.xml").read_text()
    # Current design: the icon button lives on the built-in Insert tab (OfficeTab),
    # not a CustomTab, and reads "Insert Icons" so it pairs with the VBA "Make
    # Editable" button sitting beside it. (Replaces the old schema3-pilot CustomTab
    # + insertVectorIcon adapter, which the file-backed loader and hosted sidebar
    # made obsolete.)
    assert '<OfficeTab id="TabInsert">' in manifest_text
    assert "CustomTab" not in manifest_text
    assert "Insert Icons" in manifest_text
    assert "Make Editable" in manifest_text
    assert "schema3-pilot" not in manifest_text
    assert not [element for element in manifest.iter() if element.tag.endswith("CustomTab")]
