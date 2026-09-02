from __future__ import annotations

import argparse
from collections.abc import Sequence

from . import __version__


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-accountant",
        description="本地 XLSX 精确金额组合工具",
    )
    parser.add_argument("--version", action="version", version=__version__)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    parser.parse_args(argv)
    print("ExcelAccountant 图形界面将在当前开发分支中启用。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
