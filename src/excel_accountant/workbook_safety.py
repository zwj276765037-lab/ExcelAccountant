from __future__ import annotations

from pathlib import Path
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook

from .models import WorkbookSafetyReport


class WorkbookSafetyError(ValueError):
    """Raised when a workbook cannot be inspected safely."""


_UNSAFE_PARTS: tuple[tuple[str, str], ...] = (
    ("xl/vbaProject.bin", "包含 VBA 宏"),
    ("xl/drawings/", "包含绘图、图片或形状"),
    ("xl/charts/", "包含图表"),
    ("xl/embeddings/", "包含嵌入对象"),
    ("xl/externalLinks/", "包含外部链接"),
    ("xl/connections.xml", "包含外部数据连接"),
)


def inspect_workbook_safety(path: str | Path) -> WorkbookSafetyReport:
    workbook_path = Path(path)
    reasons: list[str] = []
    if workbook_path.suffix.lower() != ".xlsx":
        reasons.append("第一版只支持 .xlsx 文件")
    try:
        with ZipFile(workbook_path) as archive:
            names = set(archive.namelist())
    except (BadZipFile, OSError) as exc:
        raise WorkbookSafetyError(f"无法读取工作簿结构：{exc}") from exc

    for prefix, message in _UNSAFE_PARTS:
        if prefix.endswith("/"):
            found = any(name.startswith(prefix) for name in names)
        else:
            found = prefix in names
        if found and message not in reasons:
            reasons.append(message)

    try:
        workbook = load_workbook(
            workbook_path,
            read_only=False,
            data_only=False,
            keep_links=True,
        )
    except Exception as exc:
        raise WorkbookSafetyError(f"无法解析工作簿：{exc}") from exc

    try:
        security = workbook.security
        if security is not None and (
            security.lockStructure or security.lockWindows
        ):
            reasons.append("工作簿结构受到保护")
        for worksheet in workbook.worksheets:
            if worksheet.protection.sheet:
                reasons.append(f"工作表受到保护：{worksheet.title}")
    finally:
        workbook.close()

    return WorkbookSafetyReport(
        safe_to_write=not reasons,
        reasons=tuple(dict.fromkeys(reasons)),
    )
