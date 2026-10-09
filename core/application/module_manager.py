class ModuleManager:
    def __init__(self):
        self._modules = {}

    def register(self, module):
        if module.info.module_id in self._modules:
            raise ValueError(f"Duplicate module id: {module.info.module_id}")
        self._modules[module.info.module_id] = module

    def get(self, module_id):
        return self._modules[module_id]

    def all(self):
        return list(self._modules.values())
