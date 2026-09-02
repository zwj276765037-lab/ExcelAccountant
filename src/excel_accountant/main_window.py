from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import QThread, Qt, QUrl
from PySide6.QtGui import QCloseEvent, QDesktopServices, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .decimal_codec import format_decimal
from .models import ApproximateSolution
from .service import SearchReport, SearchRequest
from .worker import SearchWorker
from .xlsx_reader import list_sheet_names, read_workbook_preview


class MainWindow(QMainWindow):
    def __init__(self, initial_file: str | None = None) -> None:
        super().__init__()
        self.setWindowTitle("ExcelAccountant - Excel 金额精确凑数")
        self.resize(1120, 760)
        self.setMinimumSize(900, 620)
        self._thread: QThread | None = None
        self._worker: SearchWorker | None = None
        self._last_output_directory: Path | None = None
        self._build_ui()
        self._apply_style()
        if initial_file:
            self.file_edit.setText(initial_file)
            self._load_sheets()

    def _build_ui(self) -> None:
        root = QWidget()
        outer = QVBoxLayout(root)
        outer.setContentsMargins(18, 16, 18, 16)
        outer.setSpacing(12)

        title = QLabel("金额精确凑数")
        title.setObjectName("title")
        title.setFont(QFont("Microsoft YaHei UI", 18, QFont.Weight.Bold))
        subtitle = QLabel(
            "本地处理 XLSX，按原始十进制精度计算；同一方案内的每条金额绝不重复使用。"
        )
        subtitle.setObjectName("subtitle")
        outer.addWidget(title)
        outer.addWidget(subtitle)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self._build_input_panel())
        splitter.addWidget(self._build_data_panel())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([380, 700])
        outer.addWidget(splitter, 1)

        status_row = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(8)
        self.status_label = QLabel("就绪")
        self.status_label.setMinimumWidth(360)
        status_row.addWidget(self.progress_bar, 1)
        status_row.addWidget(self.status_label)
        outer.addLayout(status_row)
        self.setCentralWidget(root)

    def _build_input_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 12, 0)
        layout.setSpacing(10)

        source_group = QGroupBox("1. 选择数据")
        source_layout = QGridLayout(source_group)
        self.file_edit = QLineEdit()
        self.file_edit.setPlaceholderText("选择 .xlsx 文件")
        browse_file = QPushButton("浏览…")
        browse_file.clicked.connect(self._browse_file)
        self.sheet_combo = QComboBox()
        self.range_edit = QLineEdit("E")
        self.range_edit.setPlaceholderText("例如 E、5 或 E2:E500")
        source_layout.addWidget(QLabel("文件"), 0, 0)
        source_layout.addWidget(self.file_edit, 0, 1)
        source_layout.addWidget(browse_file, 0, 2)
        source_layout.addWidget(QLabel("工作表"), 1, 0)
        source_layout.addWidget(self.sheet_combo, 1, 1, 1, 2)
        source_layout.addWidget(QLabel("金额列"), 2, 0)
        source_layout.addWidget(self.range_edit, 2, 1, 1, 2)
        self.preview_button = QPushButton("预览读取结果")
        self.preview_button.clicked.connect(self._preview)
        source_layout.addWidget(self.preview_button, 3, 1, 1, 2)
        layout.addWidget(source_group)

        target_group = QGroupBox("2. 输入目标金额")
        target_layout = QVBoxLayout(target_group)
        target_hint = QLabel("每行一个目标；可一次输入多个，各目标使用不同颜色。")
        target_hint.setWordWrap(True)
        self.targets_edit = QPlainTextEdit()
        self.targets_edit.setPlaceholderText("1250.35\n980.00\n-120.5")
        self.targets_edit.setMinimumHeight(115)
        target_layout.addWidget(target_hint)
        target_layout.addWidget(self.targets_edit)
        layout.addWidget(target_group)

        settings_group = QGroupBox("3. 搜索设置")
        settings_layout = QFormLayout(settings_group)
        self.solution_count = QSpinBox()
        self.solution_count.setRange(1, 100)
        self.solution_count.setValue(20)
        self.solution_count.setSuffix(" 个文件")
        self.timeout_seconds = QSpinBox()
        self.timeout_seconds.setRange(1, 3600)
        self.timeout_seconds.setValue(60)
        self.timeout_seconds.setSuffix(" 秒")
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("默认：源文件旁的 ExcelAccountant输出")
        output_row = QWidget()
        output_layout = QHBoxLayout(output_row)
        output_layout.setContentsMargins(0, 0, 0, 0)
        output_layout.addWidget(self.output_edit, 1)
        browse_output = QPushButton("浏览…")
        browse_output.clicked.connect(self._browse_output)
        output_layout.addWidget(browse_output)
        settings_layout.addRow("最多精确方案", self.solution_count)
        settings_layout.addRow("精确搜索时限", self.timeout_seconds)
        settings_layout.addRow("输出目录", output_row)
        layout.addWidget(settings_group)

        action_row = QHBoxLayout()
        self.search_button = QPushButton("开始精确搜索")
        self.search_button.setObjectName("primary")
        self.search_button.clicked.connect(self._start_search)
        self.cancel_button = QPushButton("取消")
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self._cancel_search)
        action_row.addWidget(self.search_button, 1)
        action_row.addWidget(self.cancel_button)
        layout.addLayout(action_row)

        self.open_output_button = QPushButton("打开输出文件夹")
        self.open_output_button.setEnabled(False)
        self.open_output_button.clicked.connect(self._open_output_directory)
        layout.addWidget(self.open_output_button)
        layout.addStretch(1)
        return panel

    def _build_data_panel(self) -> QWidget:
        self.tabs = QTabWidget()
        self.preview_table = QTableWidget(0, 6)
        self.preview_table.setHorizontalHeaderLabels(
            ("位置", "原始值", "精确金额", "类型", "隐藏行", "状态")
        )
        self._configure_table(self.preview_table)
        preview_page = QWidget()
        preview_layout = QVBoxLayout(preview_page)
        self.preview_summary = QLabel("请先选择文件并预览。")
        self.preview_summary.setWordWrap(True)
        preview_layout.addWidget(self.preview_summary)
        preview_layout.addWidget(self.preview_table, 1)

        self.result_table = QTableWidget(0, 7)
        self.result_table.setHorizontalHeaderLabels(
            ("方案", "目标", "目标金额", "实际合计", "差额", "单元格位置", "输出文件")
        )
        self._configure_table(self.result_table)
        result_page = QWidget()
        result_layout = QVBoxLayout(result_page)
        self.result_summary = QLabel("搜索结果将显示在这里。")
        self.result_summary.setWordWrap(True)
        result_layout.addWidget(self.result_summary)
        result_layout.addWidget(self.result_table, 1)

        self.tabs.addTab(preview_page, "数据预览")
        self.tabs.addTab(result_page, "搜索结果")
        return self.tabs

    @staticmethod
    def _configure_table(table: QTableWidget) -> None:
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setAlternatingRowColors(True)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow, QWidget { background: #f6f8fb; color: #172033; }
            QLabel#title { color: #123b64; }
            QLabel#subtitle { color: #53657a; padding-bottom: 4px; }
            QGroupBox { background: white; border: 1px solid #dce3eb; border-radius: 8px;
                        margin-top: 10px; padding: 12px 8px 8px 8px; font-weight: 600; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }
            QLineEdit, QPlainTextEdit, QComboBox, QSpinBox, QTableWidget {
                background: white; border: 1px solid #cbd5e1; border-radius: 5px; padding: 5px;
            }
            QPushButton { background: #e7edf4; border: 1px solid #c4ceda; border-radius: 5px;
                          padding: 7px 12px; }
            QPushButton:hover { background: #dae5f0; }
            QPushButton#primary { background: #1769aa; color: white; border-color: #1769aa;
                                  font-weight: 600; padding: 9px 12px; }
            QPushButton#primary:hover { background: #10588f; }
            QPushButton:disabled { color: #8c98a7; background: #edf1f5; }
            QHeaderView::section { background: #e8eef5; padding: 6px; border: 0;
                                   border-right: 1px solid #d3dce6; font-weight: 600; }
            QTabWidget::pane { border: 1px solid #dce3eb; background: white; }
            QTabBar::tab { padding: 8px 16px; background: #e8eef5; }
            QTabBar::tab:selected { background: white; color: #1769aa; }
            QProgressBar { border: 0; background: #dfe6ee; border-radius: 4px; }
            QProgressBar::chunk { background: #2b83c6; border-radius: 4px; }
            """
        )

    def _browse_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 XLSX 文件",
            str(Path(self.file_edit.text()).parent) if self.file_edit.text() else "",
            "Excel 工作簿 (*.xlsx)",
        )
        if path:
            self.file_edit.setText(path)
            self._load_sheets()

    def _load_sheets(self) -> None:
        try:
            sheets = list_sheet_names(self.file_edit.text().strip())
        except Exception as exc:
            self.sheet_combo.clear()
            self._show_error("无法读取工作表", str(exc))
            return
        self.sheet_combo.clear()
        self.sheet_combo.addItems(sheets)
        self.status_label.setText(f"已读取 {len(sheets)} 个工作表")

    def _browse_output(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if directory:
            self.output_edit.setText(directory)

    def _preview(self) -> None:
        try:
            source, sheet, range_text = self._source_fields()
            preview = read_workbook_preview(source, sheet, range_text)
        except Exception as exc:
            self._show_error("预览失败", str(exc))
            return
        self.preview_table.setRowCount(0)
        for cell in preview.cells:
            row = self.preview_table.rowCount()
            self.preview_table.insertRow(row)
            values = (
                cell.address,
                cell.raw_value,
                format_decimal(cell.amount),
                "文本" if cell.source_type == "text" else "数值",
                "是" if cell.hidden else "否",
                "可用",
            )
            for column, value in enumerate(values):
                self.preview_table.setItem(row, column, QTableWidgetItem(value))
        safe_text = "可安全输出" if preview.safety.safe_to_write else "仅允许分析，禁止输出"
        self.preview_summary.setText(
            f"范围 {preview.range_text}：{len(preview.cells)} 条可用，"
            f"{len(preview.skipped)} 条跳过（公式 {preview.formula_count}，"
            f"零值 {preview.zero_count}，隐藏行 {preview.hidden_count}）；{safe_text}。"
        )
        self.status_label.setText("预览完成")
        self.tabs.setCurrentIndex(0)

    def _source_fields(self) -> tuple[Path, str, str]:
        file_text = self.file_edit.text().strip()
        if not file_text:
            raise ValueError("请先选择 XLSX 文件")
        source = Path(file_text)
        if source.suffix.lower() != ".xlsx":
            raise ValueError("当前版本仅支持 .xlsx 文件")
        if not source.is_file():
            raise ValueError("选择的文件不存在")
        sheet = self.sheet_combo.currentText().strip()
        if not sheet:
            raise ValueError("请选择工作表")
        range_text = self.range_edit.text().strip()
        if not range_text:
            raise ValueError("请输入金额列或单列范围")
        return source, sheet, range_text

    def _start_search(self) -> None:
        if self._thread is not None:
            return
        try:
            source, sheet, range_text = self._source_fields()
            targets = tuple(self.targets_edit.toPlainText().splitlines())
            if not any(value.strip() for value in targets):
                raise ValueError("请至少输入一个目标金额")
            output_text = self.output_edit.text().strip()
            output_directory = (
                Path(output_text)
                if output_text
                else source.parent / "ExcelAccountant输出"
            )
            request = SearchRequest(
                source_path=source,
                sheet=sheet,
                range_text=range_text,
                target_values=targets,
                output_directory=output_directory,
                max_exact_solutions=self.solution_count.value(),
                exact_time_limit_seconds=float(self.timeout_seconds.value()),
                max_approximate_solutions=3,
                approximate_time_limit_seconds=float(
                    min(30, self.timeout_seconds.value())
                ),
            )
        except Exception as exc:
            self._show_error("输入有误", str(exc))
            return

        self.result_table.setRowCount(0)
        self.result_summary.setText("正在搜索…")
        self.tabs.setCurrentIndex(1)
        self._set_running(True)
        thread = QThread(self)
        worker = SearchWorker(request)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress.connect(self._on_progress)
        worker.succeeded.connect(self._on_success)
        worker.failed.connect(self._on_failure)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(self._search_finished)
        thread.finished.connect(thread.deleteLater)
        self._thread = thread
        self._worker = worker
        thread.start()

    def _set_running(self, running: bool) -> None:
        self.search_button.setEnabled(not running)
        self.preview_button.setEnabled(not running)
        self.cancel_button.setEnabled(running)
        self.progress_bar.setRange(0, 0 if running else 1)
        if not running:
            self.progress_bar.setValue(0)

    def _cancel_search(self) -> None:
        if self._worker is not None:
            self._worker.request_cancel()
            self.cancel_button.setEnabled(False)
            self.status_label.setText("正在取消…")

    def _on_progress(self, _stage: str, message: str) -> None:
        self.status_label.setText(message)

    def _on_success(self, report: SearchReport) -> None:
        self._render_report(report)

    def _on_failure(self, message: str, _details: str) -> None:
        self.result_summary.setText(f"执行失败：{message}")
        self.status_label.setText("执行失败")
        self._show_error("搜索失败", message)

    def _search_finished(self) -> None:
        self._thread = None
        self._worker = None
        self._set_running(False)
        if self.status_label.text() == "正在取消…":
            self.status_label.setText("已取消")

    def _render_report(self, report: SearchReport) -> None:
        self.result_table.setRowCount(0)
        self.result_summary.setText("\n".join(report.messages))
        if report.artifacts:
            self._last_output_directory = report.request.output_directory
            self.open_output_button.setEnabled(True)
            for artifact, solution in zip(
                report.artifacts,
                report.exact_outcome.exact_solutions,
                strict=False,
            ):
                for assignment in solution.assignments:
                    target = report.targets[assignment.target_index]
                    cells = [report.preview.cells[index] for index in assignment.cell_indices]
                    self._append_result_row(
                        (
                            f"精确 {artifact.scheme_number:03d}",
                            target.identifier,
                            target.raw_value,
                            format_decimal(sum((cell.amount for cell in cells), target.amount * 0)),
                            "0",
                            ", ".join(cell.address for cell in cells),
                            str(artifact.path),
                        )
                    )
        elif report.approximate_outcome is not None:
            for number, solution in enumerate(
                report.approximate_outcome.approximate_solutions,
                start=1,
            ):
                self._append_approximate_solution(report, solution, number)
        self.status_label.setText("搜索完成" if not report.cancelled else "已取消")
        self.tabs.setCurrentIndex(1)

    def _append_approximate_solution(
        self,
        report: SearchReport,
        solution: ApproximateSolution,
        number: int,
    ) -> None:
        for assignment in solution.assignments:
            target = report.targets[assignment.target_index]
            cells = [report.preview.cells[index] for index in assignment.cell_indices]
            self._append_result_row(
                (
                    f"近似 {number:03d}",
                    target.identifier,
                    target.raw_value,
                    format_decimal(report.scaled_problem.restore(assignment.actual)),
                    format_decimal(report.scaled_problem.restore(assignment.difference)),
                    ", ".join(cell.address for cell in cells) or "（未使用金额）",
                    "不输出文件",
                )
            )

    def _append_result_row(self, values: tuple[str, ...]) -> None:
        row = self.result_table.rowCount()
        self.result_table.insertRow(row)
        for column, value in enumerate(values):
            item = QTableWidgetItem(value)
            item.setToolTip(value)
            self.result_table.setItem(row, column, item)

    def _open_output_directory(self) -> None:
        if self._last_output_directory is not None:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self._last_output_directory)))

    def _show_error(self, title: str, message: str) -> None:
        QMessageBox.warning(self, title, message)

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802 - Qt API
        if self._worker is not None and self._thread is not None:
            self._worker.request_cancel()
            if not self._thread.wait(5000):
                QMessageBox.information(self, "正在取消", "请等待当前搜索停止后再关闭。")
                event.ignore()
                return
        event.accept()


def create_application(argv: list[str]) -> QApplication:
    os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
    application = QApplication.instance() or QApplication(argv)
    application.setApplicationName("ExcelAccountant")
    application.setOrganizationName("ExcelAccountant")
    return application
