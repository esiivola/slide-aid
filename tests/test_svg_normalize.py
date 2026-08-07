"""Exhaustive unit tests for scripts/svg_normalize.py.

The normalizer converts arbitrary SVG path data to absolute M/L/C/Z subpaths.
Both the web sidebar previews and the editable PowerPoint freeforms are built
from its output, so a bug here distorts every icon - hence the wide coverage of
command forms, number formats, and real-world quirks seen in icon-set paths.
"""
from __future__ import annotations

import svg_normalize as N


def only(d: str, scale: float = 1.0) -> str:
    """Normalize a path expected to yield exactly one subpath; return it."""
    subs = N.normalize_field(d, scale)
    assert len(subs) == 1, subs
    return subs[0]


# --- absolute / relative basics -------------------------------------------

def test_absolute_line_passthrough():
    assert only("M0 0 L10 0") == "M0 0 L10 0"


def test_relative_line_chain_becomes_absolute():
    assert only("M10 10 l5 0 l0 5") == "M10 10 L15 10 L15 15"


def test_h_and_v_become_absolute_lines():
    assert only("M0 0 H10 V5") == "M0 0 L10 0 L10 5"
    assert only("M2 2 h3 v-1") == "M2 2 L5 2 L5 1"


def test_close_path_emits_Z():
    assert only("M0 0 L10 0 L10 10 Z") == "M0 0 L10 0 L10 10 Z"


def test_cubic_absolute_and_relative_passthrough():
    assert only("M0 0 C1 2 3 4 5 6") == "M0 0 C1 2 3 4 5 6"
    assert only("M10 10 c1 2 3 4 5 6") == "M10 10 C11 12 13 14 15 16"


def test_relative_cubic_chain():
    assert only("M10 10 c0 5 5 5 5 0 c0 -5 5 -5 5 0") == \
        "M10 10 C10 15 15 15 15 10 C15 5 20 5 20 10"


# --- implicit repeated commands (SVG shorthand) ---------------------------

def test_implicit_moveto_pairs_become_lines():
    assert only("M0 0 1 1 2 2") == "M0 0 L1 1 L2 2"


def test_implicit_repeated_lineto():
    assert only("M0 0 L1 0 2 0 3 0") == "M0 0 L1 0 L2 0 L3 0"


def test_implicit_repeated_curveto():
    assert only("M0 0 C0 0 1 1 2 2 2 2 3 3 4 4") == "M0 0 C0 0 1 1 2 2 C2 2 3 3 4 4"


# --- curve conversions -----------------------------------------------------

def test_quadratic_becomes_cubic():
    # Both cubic controls sit 2/3 of the way to the quadratic control point.
    assert only("M0 0 Q6 12 12 0") == "M0 0 C4 8 8 8 12 0"


def test_smooth_cubic_reflects_previous_control():
    assert only("M0 0 C0 4 4 4 8 4 S12 0 16 4") == "M0 0 C0 4 4 4 8 4 C12 4 12 0 16 4"


def test_smooth_cubic_after_line_uses_current_point():
    assert only("M0 0 L5 0 S10 5 10 0") == "M0 0 L5 0 C5 0 10 5 10 0"


def test_smooth_quadratic_after_line_uses_current_point():
    assert only("M0 0 L6 0 T12 0") == "M0 0 L6 0 C6 0 8 0 12 0"


def test_chained_smooth_quadratic_reflects():
    assert only("M0 0 Q3 3 6 0 T12 0") == "M0 0 C2 2 4 2 6 0 C8 -2 10 -2 12 0"


def test_arc_flattens_to_cubics_with_exact_endpoint():
    out = only("M0 0 A5 5 0 0 1 10 0")
    assert out.startswith("M0 0 C")
    assert "A" not in out and "L" not in out
    assert out.endswith("10 0")            # endpoint preserved exactly


def test_arc_zero_radius_degrades_to_line_segment():
    assert only("M0 0 A0 5 0 0 1 10 0") == "M0 0 C0 0 10 0 10 0"


def test_concatenated_arc_flags_parse():
    # Flags glued to the endpoint ("0 1" written inside "013 4") must still parse.
    out = only("M0 0 a5 5 0 013 4")
    assert out.startswith("M0 0 C")
    assert out.endswith("3 4")             # relative (3,4) from the origin


# --- number formats seen in real icon paths -------------------------------

def test_number_formats_leading_dot_glued_negative_and_exponent():
    assert only("M0 0 L.5 -.5") == "M0 0 L0.5 -0.5"
    assert only("M0 0L10-5") == "M0 0 L10 -5"       # glued command + negative
    assert only("M0 0 L1e1 0") == "M0 0 L10 0"      # scientific notation


def test_commas_and_extra_whitespace_are_tolerated():
    assert only("M0,0 L10,0") == "M0 0 L10 0"
    assert only("   M 0 0   L 10 0  ") == "M0 0 L10 0"


# --- subpaths, scaling, precision -----------------------------------------

def test_multiple_subpaths_split():
    assert N.normalize_field("M0 0 L1 0 Z M5 5 L6 5 Z") == ["M0 0 L1 0 Z", "M5 5 L6 5 Z"]


def test_scale_multiplies_line_coordinates():
    assert only("M1 1 L2 2", scale=3.0) == "M3 3 L6 6"


def test_scale_applies_to_curve_control_points():
    # bootstrap-style 16 -> 24 grid (scale 1.5).
    assert only("M0 0 C0 8 8 16 16 16", scale=1.5) == "M0 0 C0 12 12 24 24 24"


def test_precision_two_decimals_and_trailing_zeros_stripped():
    assert only("M0 0 L3.14159 0") == "M0 0 L3.14 0"
    assert only("M0 0 L10.0 0") == "M0 0 L10 0"


def test_degenerate_subpaths_dropped():
    assert N.normalize_field("") == []
    assert N.normalize_field("M5 5") == []          # single point, no drawable segment
    assert N.normalize_field("   ") == []
