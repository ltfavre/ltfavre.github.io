"""Standard SCOPE preparation operation - delegates to the proven SOE Part A."""
from __future__ import annotations

from .engine import run_legacy_operation


def run(config, profile, progress=None):
    return run_legacy_operation("A", config, profile, progress)
