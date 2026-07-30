"""Pure geometry helpers shared by documentation rendering and tests."""
from __future__ import annotations


def golden_top(master_top: float, master_height: float, target_height: float) -> float:
    return master_top + (master_height - target_height) / 3


def scale_box_about(
    box: tuple[float, float, float, float],
    center: tuple[float, float],
    factor: float,
) -> tuple[float, float, float, float]:
    cx, cy = center
    return (
        cx + (box[0] - cx) * factor,
        cy + (box[1] - cy) * factor,
        cx + (box[2] - cx) * factor,
        cy + (box[3] - cy) * factor,
    )
