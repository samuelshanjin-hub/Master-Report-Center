from dataclasses import dataclass

@dataclass(frozen=True)
class FileInfo:
    path: str
    name: str
    suffix: str
    size: int
