"""Shared pytest setup: put scripts/ on sys.path so tests can import the
icon-pipeline modules (svg_normalize, build_iconaid_web) directly."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
