"""Standalone entry point for the SCOPE Sample Planning application."""
from __future__ import annotations

import os
from pathlib import Path
import sys

os.environ["SCOPE_APP_MODE"] = "sample_planning"
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from gui.app import SCOPEApp


if __name__ == "__main__":
    SCOPEApp()
