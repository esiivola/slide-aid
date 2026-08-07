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
import svg_normalize as N

ROOT = Path(__file__).resolve().parents[1]
CATALOG = ROOT / "apps" / "powerpoint-iconaid" / "catalog.json"
DAT = ROOT / "apps" / "powerpoint" / "data" / "icons.dat"


# --- pure pieces -----------------------------------------------------------

def test_variant_filter_drops_style_variants():
    for vid in ("tabler-home-fill", "lucide-x-filled", "bootstrap-star-solid",
                "hero-y-duotone", "phosphor-heart-thin", "phosphor-heart-light",
                "phosphor-heart-bold"):
        assert B.is_variant(vid), f"{vid} should be treated as a style variant"


def test_variant_filter_keeps_real_names():
    # Suffix-anchored, phosphor-weight scoped: legitimate names survive.
    for keep in ("bootstrap-heart", "lucide-home", "tabler-traffic-light",
                 "phosphor-house", "bootstrap-solidarity", "lucide-fill-color",
                 "tabler-arrow-up"):
        assert not B.is_variant(keep), f"{keep} should be kept"


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
         "source": "tabler", "tags": [], "paths": ["M0 0 L1 1 L2 0 Z"]},  # variant -> dropped
        {"id": "tabler-empty", "name": "Empty", "category": "General",
         "source": "tabler", "tags": [], "paths": ["M5 5"]},              # no segment -> skipped
        {"id": "tabler-pipe", "name": "A|B", "category": "General",       # pipe must be escaped
         "source": "tabler", "tags": ["x|y"], "paths": ["M0 0 L3 3"]},
    ]}
    srcf = tmp_path / "unified-catalog.json"
    srcf.write_text(json.dumps(src))
    webf = tmp_path / "catalog.json"
    datf = tmp_path / "icons.dat"
    monkeypatch.setattr(B, "SRC", srcf)
    monkeypatch.setattr(B, "WEB_OUT", webf)
    monkeypatch.setattr(B, "DAT_OUT", datf)

    B.main()

    cat = json.loads(webf.read_text())
    ids = [ic["id"] for ic in cat]
    assert ids == ["tabler-home", "bootstrap-circle", "tabler-pipe"]  # variant + empty dropped, order kept

    circle = next(ic for ic in cat if ic["id"] == "bootstrap-circle")
    assert circle["d"][0].startswith("M") and "C" in circle["d"][0]   # arc flattened to cubics
    assert cat[-1]["n"] == "A|B"                                      # raw name kept in JSON

    lines = [ln for ln in datf.read_text().split("\n") if ln]
    assert len(lines) == len(cat)
    assert [ln.split("|", 1)[0] for ln in lines] == ids               # same order
    pipe_line = next(ln for ln in lines if ln.startswith("tabler-pipe|"))
    fields = pipe_line.split("|")
    assert fields[1] == "A/B"                                         # '|' escaped so fields stay aligned
    assert len(fields) >= 5                                           # id|name|cat|tags|>=1 path


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
def test_filled_set_is_exactly_the_bootstrap_source():
    # Sidebar + add-in classify s=='bootstrap' OR id ending -solid/-mini as
    # "filled"; after variant filtering that set is exactly the bootstrap icons.
    cat = json.loads(CATALOG.read_text())
    filled = [ic for ic in cat
              if ic["s"] == "bootstrap" or re.search(r"-(solid|mini)$", ic["id"])]
    assert filled
    assert all(ic["s"] == "bootstrap" for ic in filled)


@pytest.mark.skipif(not CATALOG.exists(), reason="generated; run scripts/build_iconaid_web.py")
def test_sidebar_and_vba_fill_classification_agree():
    # taskpane.js keys "filled" off the source ('bootstrap'); the VBA keys it off
    # the id prefix ('bootstrap-'). They must agree for every real icon.
    cat = json.loads(CATALOG.read_text())
    for ic in cat:
        assert (ic["s"] == "bootstrap") == ic["id"].startswith("bootstrap-")
