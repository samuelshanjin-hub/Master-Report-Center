import pandas as pd

class MappingEngine:
    @staticmethod
    def composite_key(df, columns):
        if not columns:
            raise ValueError("At least one match column is required")
        values = []
        for col in columns:
            if col not in df.columns:
                raise ValueError(f"Column not found: {col}")
            values.append(df[col].fillna("").astype(str).str.strip())
        key = values[0]
        for series in values[1:]:
            key = key + "" + series
        return key
