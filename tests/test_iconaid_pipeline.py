"""Exhaustive tests for scripts/build_iconaid_web.py - the generator that turns
the unified catalog into the sidebar's catalog.json and the add-in's icons.dat.

Covers the variant filter, per-source viewBox scaling, and a full synthetic
end-to-end run (variant drop, empty-path skip, pipe escaping, ordering, both
outputs). The checks against the *real* generated files run only when those
gitignored artifacts are present.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

import build_iconaid_web as B
import fetch_all_icons_fast as F
import normalize_all_icons as U
import svg_normalize as N

ROOT = Path(__file__).resolve().parents[1]
CATALOG = ROOT / "apps" / "powerpoint-iconaid" / "catalog.json"
DAT = ROOT / "apps" / "powerpoint" / "data" / "icons.dat"
SLIDES_INDEX = ROOT / "shared" / "iconaid" / "slides" / "index.json"


# --- pure pieces -----------------------------------------------------------

def test_variant_filter_drops_style_variants():
    for vid in ("phosphor-heart-thin", "phosphor-heart-light", "phosphor-heart-bold"):
        assert B.is_variant(vid), f"{vid} should be treated as a style variant"


def test_variant_filter_keeps_real_names():
    # Suffix-anchored, phosphor-weight scoped: legitimate names survive.
    for keep in ("bootstrap-heart", "lucide-home", "tabler-traffic-light",
                 "phosphor-house", "bootstrap-solidarity", "lucide-fill-color",
                 "tabler-arrow-up", "tabler-home-fill", "bootstrap-star-solid",
                 "hero-y-duotone"):
        assert not B.is_variant(keep), f"{keep} should be kept"


def test_reviewed_source_manifest_is_permissive_and_excludes_remix():
    manifest = json.loads(F.SOURCE_MANIFEST.read_text())
    allowed = {"MIT", "ISC", "Apache-2.0", "CC0-1.0"}
    assert manifest["schema"] == 1
    assert all(source["license"] in allowed for source in manifest["sources"] if source.get("enabled"))
    assert all(source.get("version") for source in manifest["sources"] if source.get("enabled"))
    assert {
        "iconoir", "hugeicons", "icon-park-outline", "mingcute", "carbon",
        "material-symbols", "fluent", "simple-icons", "healthicons",
    } <= {
        source["id"] for source in manifest["sources"] if source.get("enabled")
    }
    assert any(source["id"] == "remixicon" for source in manifest["excluded"])


def test_iconify_import_preserves_categories_aliases_and_render_mode(tmp_path):
    package = tmp_path / "package"
    package.mkdir()
    (package / "icons.json").write_text(json.dumps({
        "width": 48,
        "height": 48,
        "icons": {
            "brain-research": {
                "body": '<g fill="none" stroke="currentColor"><path d="M4 4 L44 44"/><circle cx="24" cy="24" r="4"/></g>'
            },
            "diamond-fill": {
                "body": '<path fill="currentColor" d="M4 8 L44 8 L24 44 Z"/>'
            },
        },
        "aliases": {"ai-lab": {"parent": "brain-research"}},
        "categories": {"Science & Technology": ["brain-research"], "Design": ["diamond-fill"]},
    }))
    source = {
        "id": "sample", "name": "Sample", "license": "MIT",
        "renderMode": "auto", "viewBox": 48,
    }
    icons = F.process_iconify_json(tmp_path, source)
    by_id = {icon["id"]: icon for icon in icons}
    assert by_id["brain-research"]["renderMode"] == "stroke"
    assert by_id["diamond-fill"]["renderMode"] == "fill"
    assert by_id["brain-research"]["viewBox"] == 48
    assert "science & technology" in by_id["brain-research"]["tags"]
    assert "ai-lab" in by_id["brain-research"]["tags"]
    assert len(by_id["brain-research"]["svg_paths"]) == 2


def test_iconify_import_applies_reviewed_variant_filter_and_source_terms(tmp_path):
    package = tmp_path / "package"
    package.mkdir()
    (package / "icons.json").write_text(json.dumps({
        "width": 24,
        "height": 24,
        "icons": {
            "heart-20-regular": {"body": '<path fill="currentColor" d="M1 1 L2 2"/>'},
            "heart-24-regular": {"body": '<path fill="currentColor" d="M1 1 L2 2"/>'},
            "heart-24-filled": {"body": '<path fill="currentColor" d="M1 1 L2 2"/>'},
        },
    }))
    source = {
        "id": "sample", "name": "Sample", "license": "MIT",
        "renderMode": "auto", "viewBox": 24,
        "includePattern": r"-24-(?:regular|filled)$",
        "categoryOverride": "Health",
        "sourceTags": ["healthcare", "patient care"],
    }

    icons = F.process_iconify_json(tmp_path, source)

    assert {icon["id"] for icon in icons} == {"heart-24-regular", "heart-24-filled"}
    assert all(icon["category"] == "Health" for icon in icons)
    assert all({"healthcare", "patient care"} <= set(icon["tags"]) for icon in icons)


def test_category_and_tag_matching_use_tokens_not_substrings():
    assert F.categorize_by_name("clock-24-hours") == "Business"
    assert F.categorize_by_name("door-lock") == "Security"
    assert "delete" not in F.generate_tags("taxi")
    assert "delete" in F.generate_tags("x")


def test_unified_icon_keeps_rich_search_terms_and_fill_contract():
    source = {
        "id": "sample", "name": "Sample", "license": "MIT",
        "upstream": "https://example.test", "viewBox": 48, "renderMode": "auto",
    }
    raw_tags = [f"upstream-term-{index}" for index in range(35)]
    icon = U.create_unified_icon({
        "id": "brain-research-artificial-intelligence-laboratory-with-a-very-long-name",
        "name": "Brain Research",
        "category": "Technology",
        "tags": raw_tags,
        "viewBox": 48,
        "renderMode": "fill",
        "svg_paths": ["M4 4 L44 44"],
    }, source, "1.2.3")
    assert len(icon["id"]) <= 64 and re.search(r"-[0-9a-f]{8}-solid$", icon["id"])
    assert icon["style"]["renderMode"] == "fill"
    assert icon["style"]["viewBox"] == 48
    assert "upstream-term-34" in icon["tags"]
    assert {"ai", "machine learning", "computing", "innovation"} & set(icon["tags"])
    assert "upstream-term-34" in icon["searchable"]

    outline = U.create_unified_icon({
        "id": "mail-out-solid", "name": "Mail Out Solid", "category": "Communication",
        "renderMode": "stroke", "svg_paths": ["M1 1 L2 2"],
    }, source)
    assert outline["id"].endswith("-solid-outline")


def test_normalized_id_collisions_keep_every_filled_icon_findable():
    icons = [
        {"id": "carbon-skip-back-outline-solid", "name": "Skip Back Outline", "source": "carbon",
         "category": "Media", "tags": ["previous"], "searchable": "stale"},
        {"id": "carbon-skip-back-outline-solid", "name": "Skip Back Outline Solid", "source": "carbon",
         "category": "Media", "tags": ["previous", "filled"], "searchable": "stale"},
    ]

    unique = U.deduplicate_icons(icons)

    assert len(unique) == 2
    assert len({icon["id"] for icon in unique}) == 2
    assert all(icon["id"].endswith("-solid") for icon in unique)
    renamed = next(icon for icon in unique if icon["name"].endswith("Solid"))
    assert renamed["id"] in renamed["searchable"]
    assert "previous" in renamed["searchable"]


def test_viewbox_scale_factors():
    assert B.VIEWBOX["bootstrap"] == 16
    assert B.VIEWBOX["phosphor"] == 256
    assert B.VIEWBOX["tabler"] == B.VIEWBOX["lucide"] == B.VIEWBOX["heroicons"] == 24
    # A bootstrap (16-grid) coordinate lands on the common 24-grid after scaling.
    assert N.normalize_field("M0 0 L16 16", 24.0 / B.VIEWBOX["bootstrap"]) == ["M0 0 L24 24"]


# --- full synthetic end-to-end (no dependency on the real 10k catalog) -----

def test_build_end_to_end(tmp_path, monkeypatch):
    src = {"icons": [
        {"id": "tabler-home", "name": "Home", "category": "General",
         "source": "tabler", "tags": ["house", "building"],
         "paths": ["M4 12 L12 4 L20 12"]},
        {"id": "bootstrap-circle", "name": "Circle", "category": "Shapes",
         "source": "bootstrap", "tags": [],
         "paths": ["M2 0 A2 2 0 1 0 2 4 A2 2 0 1 0 2 0 Z"]},          # arc -> cubics
        {"id": "tabler-home-fill", "name": "Home Fill", "category": "General",
         "source": "tabler", "tags": [], "paths": ["M0 0 L1 1 L2 0 Z"]},
        {"id": "tabler-empty", "name": "Empty", "category": "General",
         "source": "tabler", "tags": [], "paths": ["M5 5"]},              # no segment -> skipped
        {"id": "tabler-pipe", "name": "A|B", "category": "General",       # pipe must be escaped
         "source": "tabler", "tags": ["x|y"], "paths": ["M0 0 L3 3"]},
        {"id": "mingcute-diamond-fill-solid", "name": "Diamond", "category": "Design",
         "source": "mingcute", "tags": [f"term-{i}" for i in range(40)],
         "style": {"viewBox": 48, "renderMode": "fill"},
         "paths": ["M4 8 L44 8 L24 44 Z"]},
    ]}
    srcf = tmp_path / "unified-catalog.json"
    srcf.write_text(json.dumps(src))
    webf = tmp_path / "catalog.json"
    datf = tmp_path / "icons.dat"
    slidesd = tmp_path / "slides"
    monkeypatch.setattr(B, "SRC", srcf)
    monkeypatch.setattr(B, "WEB_OUT", webf)
    monkeypatch.setattr(B, "DAT_OUT", datf)
    # Every output path must be redirected, or the build writes its fixture over
    # the real generated catalog.
    monkeypatch.setattr(B, "SLIDES_DIR", slidesd)

    B.main()

    cat = json.loads(webf.read_text())
    ids = [ic["id"] for ic in cat]
    assert ids == ["tabler-home", "bootstrap-circle", "tabler-home-fill", "tabler-pipe", "mingcute-diamond-fill-solid"]

    circle = next(ic for ic in cat if ic["id"] == "bootstrap-circle")
    assert circle["d"][0].startswith("M") and "C" in circle["d"][0]   # arc flattened to cubics
    assert next(icon for icon in cat if icon["id"] == "tabler-pipe")["n"] == "A|B"  # raw name kept in JSON

    lines = [ln for ln in datf.read_text().split("\n") if ln]
    assert len(lines) == len(cat)
    assert [ln.split("|", 1)[0] for ln in lines] == ids               # same order
    pipe_line = next(ln for ln in lines if ln.startswith("tabler-pipe|"))
    fields = pipe_line.split("|")
    assert fields[1] == "A/B"                                         # '|' escaped so fields stay aligned
    assert len(fields) >= 5                                           # id|name|cat|tags|>=1 path

    # Google Slides target: metadata index plus sharded path data, same icons.
    index = json.loads((slidesd / "index.json").read_text())
    assert sorted(ic["id"] for ic in index["icons"]) == sorted(ids)
    # Sorted by id so each shard is a contiguous range both sides can binary-search.
    assert [ic["id"] for ic in index["icons"]] == sorted(ids)
    assert index["shards"][0]["firstId"] == sorted(ids)[0]
    assert index["shards"][-1]["lastId"] == sorted(ids)[-1]
    # Solid sources are flagged so the renderers fill them instead of stroking.
    flags = {ic["id"]: ic.get("f", 0) for ic in index["icons"]}
    assert flags["bootstrap-circle"] == 1 and flags["tabler-home"] == 0
    assert flags["mingcute-diamond-fill-solid"] == 1
    expanded = next(icon for icon in cat if icon["id"] == "mingcute-diamond-fill-solid")
    assert "term-39" in expanded["t"]
    assert expanded["d"] == ["M2 4 L22 4 L12 22 Z"]

    shard_paths = {}
    for entry in index["shards"]:
        shard_paths.update(json.loads((slidesd / f"paths-{entry['shard']:02d}.json").read_text()))
    assert sorted(shard_paths) == sorted(ids)
    # The Slides shards carry exactly the paths the task pane catalog carries.
    for icon in cat:
        assert shard_paths[icon["id"]] == icon["d"]


# --- consistency of the real generated files (when present) ----------------

@pytest.mark.skipif(not CATALOG.exists(), reason="generated; run scripts/build_iconaid_web.py")
def test_catalog_shape_and_variants_filtered():
    cat = json.loads(CATALOG.read_text())
    assert isinstance(cat, list) and len(cat) > 1000
    for ic in cat:
        assert {"id", "n", "c", "s", "t", "d"} <= set(ic)
        assert isinstance(ic["d"], list) and ic["d"]
        assert all(isinstance(s, str) and s.startswith("M") for s in ic["d"])
    assert not any(B.is_variant(ic["id"]) for ic in cat)              # filter actually ran


@pytest.mark.skipif(not (CATALOG.exists() and DAT.exists()),
                    reason="generated; run scripts/build_iconaid_web.py")
def test_icons_dat_lines_up_with_catalog():
    cat = json.loads(CATALOG.read_text())
    lines = [ln for ln in DAT.read_text().split("\n") if ln]
    assert len(lines) == len(cat)
    assert [ln.split("|", 1)[0] for ln in lines] == [ic["id"] for ic in cat]
    assert all(len(ln.split("|")) >= 5 for ln in lines)


@pytest.mark.skipif(not CATALOG.exists(), reason="generated; run scripts/build_iconaid_web.py")
def test_filled_flags_are_legacy_compatible():
    cat = json.loads(CATALOG.read_text())
    filled = [ic for ic in cat if ic.get("f")]
    assert filled
    assert all(ic["s"] == "bootstrap" or re.search(r"-(solid|mini)$", ic["id"]) for ic in filled)


@pytest.mark.skipif(not CATALOG.exists(), reason="generated; run scripts/build_iconaid_web.py")
def test_sidebar_and_editable_fill_classification_agree():
    cat = json.loads(CATALOG.read_text())
    for ic in cat:
        assert bool(ic.get("f")) == bool(
            ic["s"] == "bootstrap" or re.search(r"-(solid|mini)$", ic["id"])
        )


def test_committed_slides_catalog_contains_the_expanded_library():
    index = json.loads(SLIDES_INDEX.read_text())
    expected_sources = {
        "tabler", "lucide", "heroicons", "phosphor", "bootstrap", "iconoir",
        "hugeicons", "icon-park-outline", "mingcute", "carbon",
        "material-symbols", "fluent", "simple-icons", "healthicons",
    }
    assert index["schema"] == 5
    assert len(index["icons"]) >= 54_000
    assert {icon["s"] for icon in index["icons"]} == expected_sources
    assert len(index["shards"]) < 200
    for shard in index["shards"]:
        assert (SLIDES_INDEX.parent / f"paths-{shard['shard']:02d}.json").is_file()
