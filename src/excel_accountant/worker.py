from __future__ import annotations

import threading
import traceback
from pathlib import Path

from PySide6.QtCore import QObject, Signal, Slot

from .service import SearchReport, SearchRequest, export_selected_solutions, run_search


class SearchWorker(QObject):
    progress = Signal(str, str)
    succeeded = Signal(object)
    failed = Signal(str, str)
    finished = Signal()

    def __init__(self, request: SearchRequest) -> None:
        super().__init__()
        self.request = request
        self.cancel_event = threading.Event()

    def request_cancel(self) -> None:
        self.cancel_event.set()

    @Slot()
    def run(self) -> None:
        try:
            report = run_search(
                self.request,
                cancel_event=self.cancel_event,
                progress=self.progress.emit,
            )
            self.succeeded.emit(report)
        except Exception as exc:  # UI boundary: turn all failures into a visible message.
            self.failed.emit(str(exc), traceback.format_exc())
        finally:
            self.finished.emit()


class ExportWorker(QObject):
    progress = Signal(str, str)
    succeeded = Signal(object)
    failed = Signal(str, str)
    finished = Signal()

    def __init__(
        self,
        report: SearchReport,
        solution_indices: tuple[int, ...],
        output_directory: Path,
    ) -> None:
        super().__init__()
        self.report = report
        self.solution_indices = solution_indices
        self.output_directory = output_directory
        self.cancel_event = threading.Event()

    def request_cancel(self) -> None:
        self.cancel_event.set()

    @Slot()
    def run(self) -> None:
        try:
            artifacts = export_selected_solutions(
                self.report,
                self.solution_indices,
                self.output_directory,
                cancel_event=self.cancel_event,
                progress=self.progress.emit,
            )
            self.succeeded.emit(artifacts)
        except Exception as exc:  # UI boundary: turn all failures into a visible message.
            self.failed.emit(str(exc), traceback.format_exc())
        finally:
            self.finished.emit()
