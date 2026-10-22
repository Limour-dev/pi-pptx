#!/usr/bin/env python3
"""CLI entry point for the pi-pptx backend.

The implementation lives in the ``pi_pptx`` package next to this file:

    python/pi_pptx/utils.py    constants, errors, atomic load/save, arg parsing
    python/pi_pptx/shapes.py   shape tree, resolution, inspection, notes lint
    python/pi_pptx/text.py     runs/paragraphs/letter-spacing
    python/pi_pptx/tables.py   table styling (fills, borders, rich cells)
    python/pi_pptx/slides.py   clone / layout lookup / deletion helpers
    python/pi_pptx/render.py   LibreOffice render, cache, montage, heuristics
    python/pi_pptx/ops.py      the mutation layer shared with `apply`
    python/pi_pptx/cli.py      argparse commands and main()
    python/pi_pptx/__init__.py importable `Deck` API (from pi_pptx import Deck)

This shim keeps ``python3 python/pptkit.py <cmd>`` working unchanged.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pi_pptx.cli import main  # noqa: E402

if __name__ == "__main__":
    main()
