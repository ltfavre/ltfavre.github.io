"""SCOPE - Sample Configuration, Optimization, and Profiling Engine.

The public workflow functions are loaded only when requested.  This keeps
lightweight administrative tools, such as the Test Profile Manager, from
loading the pandas/optimizer runtime merely to import ``scope.profiles``.
"""
from __future__ import annotations

from typing import Any


__all__ = ["run", "validate", "create_template", "run_scope"]


def __getattr__(name: str) -> Any:
    if name in {"run", "validate"}:
        from .engine import run, validate
        return {"run": run, "validate": validate}[name]
    if name == "create_template":
        from .recruitment import create_template
        return create_template
    if name == "run_scope":
        from .runner import run_scope
        return run_scope
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
