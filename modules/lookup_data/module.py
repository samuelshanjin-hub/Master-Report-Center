from core.models.module_info import ModuleInfo
from .page import LookupDataPage
class LookupDataModule:
    info=ModuleInfo("lookup_data","Lookup Data","Fill missing fields from a reference file","Use composite keys to fill blank target fields from a larger reference workbook without overwriting populated cells.","LOOKUP")
    def create_page(self,task_manager,logger,back_callback): return LookupDataPage(self.info,task_manager,logger,back_callback)
