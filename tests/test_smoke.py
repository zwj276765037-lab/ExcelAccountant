from excel_accountant import __version__
from excel_accountant.__main__ import main


def test_version_is_defined() -> None:
    assert __version__ == "0.1.0"


def test_cli_help_entrypoint(capsys) -> None:
    assert main([]) == 0
    assert "ExcelAccountant" in capsys.readouterr().out
