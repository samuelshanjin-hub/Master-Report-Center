from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

@dataclass
class JobConfig:
    source_files: List[str] = field(default_factory=list)
    target_file: Optional[str] = None
    output_folder: Optional[str] = None
    output_file: Optional[str] = None
    options: Dict[str, Any] = field(default_factory=dict)
    mappings: Dict[str, Any] = field(default_factory=dict)
