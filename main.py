"""Entry point for Excel -> DBF Converter v3 (GUI only)."""

from __future__ import annotations

import sys

VERSION = "1.0.0"


def main() -> int:
    from gui import run_gui

    # Support passing a workbook path as a command-line argument
    initial_file = sys.argv[1] if len(sys.argv) > 1 else None
    return run_gui(initial_file)


if __name__ == "__main__":
    raise SystemExit(main())
