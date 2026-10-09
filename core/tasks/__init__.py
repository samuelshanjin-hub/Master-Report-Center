import traceback
from PyQt6.QtCore import QObject, pyqtSignal, pyqtSlot
from core.tasks.task_context import TaskContext


class TaskWorker(QObject):
    progress = pyqtSignal(int, str)
    log = pyqtSignal(str)
    completed = pyqtSignal(object)
    failed = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, task_callable):
        super().__init__()
        self.task_callable = task_callable

    @pyqtSlot()
    def run(self):
        context = TaskContext(self.progress.emit, self.log.emit)
        try:
            result = self.task_callable(context)
            self.completed.emit(result)
        except Exception as exc:
            self.log.emit(traceback.format_exc())
            self.failed.emit(str(exc))
        finally:
            self.finished.emit()

