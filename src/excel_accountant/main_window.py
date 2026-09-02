from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import QMimeData, QThread, Qt, QUrl
from PySide6.QtGui import QCloseEvent, QDesktopServices, QDragEnterEvent, QDropEvent, QFont
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
    QScrollArea,
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
from .worker import ExportWorker, SearchWorker
from .xlsx_reader import list_sheet_names, read_workbook_preview


class MainWindow(QMainWindow):
    def __init__(self, initial_file: str | None = None) -> None:
        super().__init__()
        self.setWindowTitle("ExcelAccountant - Excel 金额精确凑数")
        self.resize(1120, 760)
        self.setMinimumSize(900, 620)
        self._thread: QThread | None = None
        self._worker: SearchWorker | ExportWorker | None = None
        self._last_output_directory: Path | None = None
        self._last_report: SearchReport | None = None
        self._preview_signature: tuple[str, str, str] | None = None
        self._preview_data = None
        self.setAcceptDrops(True)
        self._build_ui()
        self._apply_style()
        self._enforce_control_sizes()
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
        panel.setMinimumWidth(350)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 12, 0)
        layout.setSpacing(10)

        source_group = QGroupBox("1. 选择数据")
        source_layout = QGridLayout(source_group)
        self.file_edit = QLineEdit()
        self.file_edit.setPlaceholderText("选择 .xlsx 文件")
        self.file_edit.textChanged.connect(self._invalidate_preview)
        browse_file = QPushButton("浏览…")
        browse_file.clicked.connect(self._browse_file)
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentIndexChanged.connect(self._invalidate_preview)
        self.range_edit = QLineEdit("E")
        self.range_edit.setPlaceholderText("例如 E、5 或 E2:E500")
        self.range_edit.textChanged.connect(self._invalidate_preview)
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

        self.export_button = QPushButton("输出勾选方案")
        self.export_button.setEnabled(False)
        self.export_button.clicked.connect(self._start_export)
        layout.addWidget(self.export_button)
        self.open_output_button = QPushButton("打开输出文件夹")
        self.open_output_button.setEnabled(False)
        self.open_output_button.clicked.connect(self._open_output_directory)
        layout.addWidget(self.open_output_button)
        layout.addStretch(1)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        scroll.setWidget(panel)
        scroll.setMinimumWidth(370)
        return scroll

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

        self.result_table = QTableWidget(0, 8)
        self.result_table.setHorizontalHeaderLabels(
            ("选择", "方案", "目标", "目标金额", "实际合计", "差额", "单元格位置", "输出文件")
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
            QLineEdit, QComboBox, QSpinBox {
                background: white; border: 1px solid #cbd5e1; border-radius: 5px; padding: 3px 6px;
            }
            QPlainTextEdit, QTableWidget {
                background: white; border: 1px solid #cbd5e1; border-radius: 5px; padding: 5px;
            }
            QPushButton { background: #e7edf4; border: 1px solid #c4ceda; border-radius: 5px;
                          padding: 3px 10px; }
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

    def _enforce_control_sizes(self) -> None:
        """Keep text visible under Windows high-DPI and enlarged font settings."""

        for button in self.findChildren(QPushButton):
            button.setMinimumHeight(32)
        self.search_button.setMinimumHeight(38)
        for line_edit in self.findChildren(QLineEdit):
            line_edit.setMinimumHeight(30)
        for combo_box in self.findChildren(QComboBox):
            combo_box.setMinimumHeight(30)
        for spin_box in self.findChildren(QSpinBox):
            spin_box.setMinimumHeight(30)

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
        if len(sheets) == 1:
            self.sheet_combo.addItem(sheets[0], sheets[0])
        else:
            self.sheet_combo.addItem("请选择工作表…", "")
            for sheet in sheets:
                self.sheet_combo.addItem(sheet, sheet)
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
        for skipped in preview.skipped:
            row = self.preview_table.rowCount()
            self.preview_table.insertRow(row)
            values = (
                skipped.address,
                skipped.raw_value,
                "",
                "已跳过",
                "",
                skipped.reason,
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setForeground(Qt.GlobalColor.darkRed)
                self.preview_table.setItem(row, column, item)
        safe_text = "可安全输出" if preview.safety.safe_to_write else "仅允许分析，禁止输出"
        self.preview_summary.setText(
            f"范围 {preview.range_text}：{len(preview.cells)} 条可用，"
            f"{len(preview.skipped)} 条跳过（公式 {preview.formula_count}，"
            f"零值 {preview.zero_count}，隐藏行 {preview.hidden_count}）；{safe_text}。"
        )
        self.status_label.setText("预览完成")
        self.tabs.setCurrentIndex(0)
        self._preview_signature = (str(source.resolve()), sheet, range_text.upper())
        self._preview_data = preview

    def _source_fields(self) -> tuple[Path, str, str]:
        file_text = self.file_edit.text().strip()
        if not file_text:
            raise ValueError("请先选择 XLSX 文件")
        source = Path(file_text)
        if source.suffix.lower() != ".xlsx":
            raise ValueError("当前版本仅支持 .xlsx 文件")
        if not source.is_file():
            raise ValueError("选择的文件不存在")
        sheet = (self.sheet_combo.currentData() or self.sheet_combo.currentText()).strip()
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
            signature = (str(source.resolve()), sheet, range_text.upper())
            if self._preview_signature != signature or self._preview_data is None:
                raise ValueError("请先点击“预览读取结果”，确认当前数据后再搜索")
            review_items = [
                item for item in self._preview_data.skipped if item.reason != "表头"
            ]
            if review_items:
                answer = QMessageBox.question(
                    self,
                    "确认忽略异常数据",
                    f"当前有 {len(review_items)} 个非表头单元格将不参与搜索，"
                    "请在预览表格中核对原因。\n\n确定继续吗？",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if answer != QMessageBox.StandardButton.Yes:
                    return
            request = SearchRequest(
                source_path=source,
                sheet=sheet,
                range_text=range_text,
                target_values=targets,
                output_directory=output_directory,
                max_exact_solutions=self.solution_count.value(),
                exact_time_limit_seconds=float(self.timeout_seconds.value()),
                max_approximate_solutions=5,
                approximate_time_limit_seconds=float(
                    min(30, self.timeout_seconds.value())
                ),
            )
        except Exception as exc:
            self._show_error("输入有误", str(exc))
            return

        self.result_table.setRowCount(0)
        self.result_summary.setText("正在搜索…")
        self._last_report = None
        self._last_output_directory = None
        self.export_button.setEnabled(False)
        self.open_output_button.setEnabled(False)
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
        can_export = bool(
            self._last_report is not None
            and self._last_report.exact_outcome.exact_solutions
            and self._last_report.preview.safety.safe_to_write
        )
        self.export_button.setEnabled(not running and can_export)
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
        self._show_error("执行失败", message)

    def _search_finished(self) -> None:
        self._thread = None
        self._worker = None
        self._set_running(False)
        if self.status_label.text() == "正在取消…":
            self.status_label.setText("已取消")

    def _render_report(self, report: SearchReport) -> None:
        self.result_table.setRowCount(0)
        self._last_report = report
        self.result_summary.setText("\n".join(report.messages))
        if report.exact_outcome.exact_solutions:
            for solution_index, solution in enumerate(
                report.exact_outcome.exact_solutions
            ):
                for assignment_position, assignment in enumerate(solution.assignments):
                    target = report.targets[assignment.target_index]
                    cells = [report.preview.cells[index] for index in assignment.cell_indices]
                    row = self._append_result_row(
                        (
                            "",
                            f"精确 {solution_index + 1:03d}",
                            target.identifier,
                            target.raw_value,
                            format_decimal(sum((cell.amount for cell in cells), target.amount * 0)),
                            "0",
                            ", ".join(cell.address for cell in cells),
                            "待勾选输出",
                        )
                    )
                    scheme_item = self.result_table.item(row, 1)
                    scheme_item.setData(Qt.ItemDataRole.UserRole, solution_index)
                    if assignment_position == 0 and report.preview.safety.safe_to_write:
                        check_item = self.result_table.item(row, 0)
                        check_item.setText("勾选")
                        check_item.setFlags(
                            check_item.flags() | Qt.ItemFlag.ItemIsUserCheckable
                        )
                        check_item.setCheckState(Qt.CheckState.Unchecked)
                        check_item.setData(Qt.ItemDataRole.UserRole, solution_index)
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
                    "",
                    f"近似 {number:03d}",
                    target.identifier,
                    target.raw_value,
                    format_decimal(report.scaled_problem.restore(assignment.actual)),
                    format_decimal(report.scaled_problem.restore(assignment.difference)),
                    ", ".join(cell.address for cell in cells) or "（未使用金额）",
                    "不输出文件",
                )
            )

    def _append_result_row(self, values: tuple[str, ...]) -> int:
        row = self.result_table.rowCount()
        self.result_table.insertRow(row)
        for column, value in enumerate(values):
            item = QTableWidgetItem(value)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsUserCheckable)
            item.setToolTip(value)
            self.result_table.setItem(row, column, item)
        return row

    def _start_export(self) -> None:
        if self._thread is not None or self._last_report is None:
            return
        selected: list[int] = []
        for row in range(self.result_table.rowCount()):
            item = self.result_table.item(row, 0)
            if item is not None and item.checkState() == Qt.CheckState.Checked:
                selected.append(int(item.data(Qt.ItemDataRole.UserRole)))
        if not selected:
            self._show_error("未勾选方案", "请在结果表左侧至少勾选一套精确方案。")
            return
        output_text = self.output_edit.text().strip()
        output_directory = (
            Path(output_text)
            if output_text
            else self._last_report.request.source_path.parent / "ExcelAccountant输出"
        )
        self._last_output_directory = output_directory
        self._set_running(True)
        self.status_label.setText("正在输出勾选方案…")
        thread = QThread(self)
        worker = ExportWorker(self._last_report, tuple(selected), output_directory)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress.connect(self._on_progress)
        worker.succeeded.connect(self._on_export_success)
        worker.failed.connect(self._on_failure)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(self._search_finished)
        thread.finished.connect(thread.deleteLater)
        self._thread = thread
        self._worker = worker
        thread.start()

    def _on_export_success(self, artifacts: tuple) -> None:
        by_scheme = {artifact.scheme_number - 1: artifact for artifact in artifacts}
        for row in range(self.result_table.rowCount()):
            scheme_item = self.result_table.item(row, 1)
            if scheme_item is None:
                continue
            solution_index = scheme_item.data(Qt.ItemDataRole.UserRole)
            if solution_index in by_scheme:
                artifact = by_scheme[solution_index]
                output_item = QTableWidgetItem(str(artifact.path))
                output_item.setToolTip(str(artifact.path))
                self.result_table.setItem(row, 7, output_item)
                check_item = self.result_table.item(row, 0)
                if check_item is not None and (
                    check_item.flags() & Qt.ItemFlag.ItemIsUserCheckable
                ):
                    check_item.setCheckState(Qt.CheckState.Unchecked)
        self.open_output_button.setEnabled(bool(artifacts))
        message = f"已输出 {len(artifacts)} 个并通过复核的 XLSX 文件。"
        self.result_summary.setText(self.result_summary.text() + "\n" + message)
        self.status_label.setText(message)

    def _open_output_directory(self) -> None:
        if self._last_output_directory is not None:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self._last_output_directory)))

    def _show_error(self, title: str, message: str) -> None:
        QMessageBox.warning(self, title, message)

    def _invalidate_preview(self, *_args) -> None:
        self._preview_signature = None
        self._preview_data = None

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802 - Qt API
        if self._xlsx_from_mime(event.mimeData()) is not None:
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802 - Qt API
        path = self._xlsx_from_mime(event.mimeData())
        if path is not None:
            self.file_edit.setText(str(path))
            self._load_sheets()
            event.acceptProposedAction()

    @staticmethod
    def _xlsx_from_mime(mime_data: QMimeData) -> Path | None:
        for url in mime_data.urls():
            if url.isLocalFile():
                path = Path(url.toLocalFile())
                if path.suffix.lower() == ".xlsx" and path.is_file():
                    return path
        return None

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
