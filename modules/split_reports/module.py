from core.models.module_info import ModuleInfo
from .page import SplitReportsPage


class SplitReportsModule:
    info = ModuleInfo(
        "split_reports",
        "Split Reports",
        "Excel Report Splitter",
        "Split a large Excel report into multiple smaller reports based on a selected column.",
        "SPLIT",
    )

    def create_page(self, task_manager, logger, back_callback):
        return SplitReportsPage(self.info, task_manager, logger, back_callback)
