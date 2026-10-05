"""Standard SCOPE optimization operation - delegates to the proven SOE Part B."""
from __future__ import annotations

from .engine import run_legacy_operation


def run(config, profile, progress=None):
    return run_legacy_operation("B", config, profile, progress)
