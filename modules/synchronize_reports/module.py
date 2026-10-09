from core.models.module_info import ModuleInfo
from .page import SynchronizeReportsPage
class SynchronizeReportsModule:
    info=ModuleInfo("synchronize_reports","Synchronize Reports","Write updates back to a target","Match source rows to a target workbook using one or more key columns, then update selected fields.","SYNC")
    def create_page(self,task_manager,logger,back_callback): return SynchronizeReportsPage(self.info,task_manager,logger,back_callback)
