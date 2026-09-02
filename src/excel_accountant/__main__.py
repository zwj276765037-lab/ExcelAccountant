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
    parser.add_argument("workbook", nargs="?", help="启动时预先选中的 XLSX 文件")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    arguments = parser.parse_args(argv)
    from .main_window import MainWindow, create_application

    application = create_application(["excel-accountant"])
    window = MainWindow(arguments.workbook)
    window.show()
    return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())
