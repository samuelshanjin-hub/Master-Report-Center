class TaskContext:
    def __init__(self, progress_cb=None, log_cb=None):
        self._progress_cb = progress_cb or (lambda *_: None)
        self._log_cb = log_cb or (lambda *_: None)

    def progress(self, value, message=""):
        self._progress_cb(int(value), str(message))

    def log(self, message):
        self._log_cb(str(message))
