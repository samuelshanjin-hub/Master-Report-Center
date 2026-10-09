from dataclasses import dataclass, field
from typing import List

@dataclass
class JobResult:
    success: bool
    message: str
    output_files: List[str] = field(default_factory=list)
    processed_count: int = 0
    skipped_count: int = 0
    warning_count: int = 0
