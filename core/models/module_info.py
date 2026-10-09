from dataclasses import dataclass


@dataclass(frozen=True)
class ModuleInfo:
    module_id: str
    title: str
    subtitle: str
    description: str
    icon_text: str

    @property
    def id(self) -> str:
        """Backward-compatible alias used by the original Publish GUI."""
        return self.module_id
