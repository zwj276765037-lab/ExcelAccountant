from __future__ import annotations

import threading
import traceback

from PySide6.QtCore import QObject, Signal, Slot

from .service import SearchRequest, run_search


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
