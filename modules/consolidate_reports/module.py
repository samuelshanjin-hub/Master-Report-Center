from core.models.module_info import ModuleInfo
from .page import ConsolidateReportsPage
class ConsolidateReportsModule:
    info=ModuleInfo("consolidate_reports","Consolidate Reports","Combine multiple reports","Combine .xlsx files from one folder into a new workbook using a common data model.","MERGE")
    def create_page(self,task_manager,logger,back_callback): return ConsolidateReportsPage(self.info,task_manager,logger,back_callback)
