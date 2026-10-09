from dataclasses import dataclass

@dataclass(frozen=True)
class FieldMapping:
    source: str
    target: str
