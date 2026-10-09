from core.engine.pandas_engine import PandasEngine

class ExcelEngine:
    """Facade for value-oriented Excel operations.

    Production formatting-preservation work can be delegated to ComEngine.
    """
    def __init__(self):
        self.data = PandasEngine()

    def read(self, path): return self.data.read(path)
    def write(self, df, path): return self.data.write(df, path)
