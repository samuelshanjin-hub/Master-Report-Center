from PyQt6.QtCore import QThread
from core.tasks.task_worker import TaskWorker


class TaskManager:
    def __init__(self, logger):
        self.logger = logger
        self._active = []

    def start(self, task_callable, callbacks):
        thread = QThread()
        worker = TaskWorker(task_callable)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)

        worker.progress.connect(callbacks.get("progress", lambda *_: None))
        worker.log.connect(callbacks.get("log", lambda *_: None))
        worker.completed.connect(callbacks.get("completed", lambda *_: None))
        worker.failed.connect(callbacks.get("failed", lambda *_: None))

        def cleanup():
            thread.quit()
            thread.wait()
            try:
                self._active.remove((thread, worker))
            except ValueError:
                pass
            worker.deleteLater()
            thread.deleteLater()

        worker.finished.connect(cleanup)
        self._active.append((thread, worker))
        thread.start()
        return worker
