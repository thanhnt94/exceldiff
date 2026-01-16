from dataclasses import dataclass
from typing import Optional, Any

@dataclass
class CellData:
    row: int
    col: int
    value: Any
    coordinate: str

@dataclass(frozen=True)
class AnchorPoint:
    row: int
    col: int
    row_off: int = 0
    col_off: int = 0

@dataclass(frozen=True)
class ShapeData:
    id: str
    name: str
    type_name: str
    from_anchor: AnchorPoint
    to_anchor: Optional[AnchorPoint] = None
    text: Optional[str] = None
    
    def __repr__(self):
        return f"Shape(id={self.id}, name={self.name}, from={self.from_anchor.row}:{self.from_anchor.col})"

from enum import Enum

class DiffType(Enum):
    MATCH = "match"
    CHANGED = "changed"
    INSERTED = "inserted"
    DELETED = "deleted"
    MOVED = "moved" # For shapes specifically

@dataclass
class DiffItem:
    location: str # Coordinate (e.g. A1) or Shape Name
    item_type: str # "Cell" or "Shape"
    diff_type: DiffType
    old_value: Any = None
    new_value: Any = None
    details: str = ""

@dataclass
class DiffResult:
    items: list[DiffItem]

