import pandas as pd

class PandasEngine:
    @staticmethod
    def read(path, sheet_name=0):
        return pd.read_excel(path, sheet_name=sheet_name, dtype=object)

    @staticmethod
    def write(df, path):
        df.to_excel(path, index=False)

    @staticmethod
    def list_columns(path):
        return list(PandasEngine.read(path).columns)

