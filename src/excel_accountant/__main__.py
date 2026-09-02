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
    parser.add_argument("--self-test", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("workbook", nargs="?", help="启动时预先选中的 XLSX 文件")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    arguments = parser.parse_args(argv)
    if arguments.self_test:
        return run_self_test()
    from .main_window import MainWindow, create_application

    application = create_application(["excel-accountant"])
    window = MainWindow(arguments.workbook)
    window.show()
    return application.exec()


def run_self_test() -> int:
    """Exercise XLSX IO, the solver, and verified output in packaged builds."""

    from pathlib import Path
    from tempfile import TemporaryDirectory

    from openpyxl import Workbook

    from .service import SearchRequest, run_search

    with TemporaryDirectory(prefix="excel_accountant_self_test_") as directory:
        root = Path(directory)
        source = root / "self-test.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "流水"
        worksheet["E1"] = "金额"
        for row, value in enumerate(("1.10", "2.20", "3.30", "4.40"), start=2):
            worksheet.cell(row, 5, value)
        workbook.save(source)
        workbook.close()
        report = run_search(
            SearchRequest(
                source_path=source,
                sheet="流水",
                range_text="E",
                target_values=("3.30", "7.70"),
                output_directory=root / "output",
                max_exact_solutions=1,
                exact_time_limit_seconds=10,
                max_approximate_solutions=1,
                approximate_time_limit_seconds=5,
            )
        )
        if not report.artifacts or not report.artifacts[0].path.is_file():
            raise RuntimeError("打包自检未生成可验证的精确方案")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
