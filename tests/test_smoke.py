from excel_accountant import __version__
from excel_accountant.__main__ import build_parser, run_self_test


def test_version_is_defined() -> None:
    assert __version__ == "0.1.0"


def test_cli_help_entrypoint(capsys) -> None:
    build_parser().print_help()
    assert "excel-accountant" in capsys.readouterr().out


def test_packaged_self_test_path() -> None:
    assert run_self_test() == 0
