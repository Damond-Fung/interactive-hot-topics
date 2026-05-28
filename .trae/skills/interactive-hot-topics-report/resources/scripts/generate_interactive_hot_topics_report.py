#!/usr/bin/env python3
"""Thin wrapper for the interactive hot topics report script."""

from __future__ import annotations

import runpy
import sys
from pathlib import Path


def locate_repo_script() -> Path:
    current = Path(__file__).resolve()
    for parent in current.parents:
        candidate = parent / "interactive_hot_topics_report.py"
        if candidate.exists():
            return candidate
    raise FileNotFoundError("Unable to locate interactive_hot_topics_report.py from the Skill wrapper.")


def main() -> None:
    script_path = locate_repo_script()
    repo_root = script_path.parent
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))
    sys.argv = [str(script_path), *sys.argv[1:]]
    runpy.run_path(str(script_path), run_name="__main__")


if __name__ == "__main__":
    main()
