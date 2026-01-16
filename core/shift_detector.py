import difflib
from typing import List, Dict, Optional, Any
from .data_types import CellData

class ShiftDetector:
    def __init__(self):
        pass

    def compute_mapping(self, list_a: List[Any], list_b: List[Any]) -> Dict[int, Optional[int]]:
        """
        Computes a mapping from indices in A to indices in B.
        Returns: Dict { index_a: index_b (or None if deleted) }
        If index_b is None, row A was deleted.
        Any index_b not in values means row B was inserted.
        """
        matcher = difflib.SequenceMatcher(None, list_a, list_b)
        mapping: Dict[int, Optional[int]] = {}
        
        # Initialize all as deleted (None) first
        for i in range(len(list_a)):
            mapping[i] = None
            
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                # Block match: A[i1:i2] == B[j1:j2]
                for k in range(i2 - i1):
                    mapping[i1 + k] = j1 + k
            elif tag == 'replace':
                # Block modified: A[i1:i2] became B[j1:j2]
                # Heuristic: Map 1-to-1 for the minimum length of the block.
                # Remaining items in A are deleted, remaining in B are inserted.
                len_a = i2 - i1
                len_b = j2 - j1
                min_len = min(len_a, len_b)
                for k in range(min_len):
                    mapping[i1 + k] = j1 + k
                # If len_a > min_len, indices i1+min_len ... i2 are unmapped (None -> Deleted)
                # If len_b > min_len, indices j1+min_len ... j2 are unmapped (Inserted)
            # 'delete', 'insert' imply no mapping for those indices
            
        return mapping

    def get_row_signatures(self, cells: List[CellData], shapes: List[Any], max_row: int) -> List[str]:
        """
        Generates a signature for each row.
        Signature: "Cell1|Cell2|...||Shape1|Shape2..."
        """
        # Group by row
        rows = {}
        for c in cells:
            if c.row not in rows:
                rows[c.row] = []
            rows[c.row].append(str(c.value))
            
        # Add shapes to row signatures
        # Shapes are usually anchored to top-left (from_anchor)
        for s in shapes:
            r = s.from_anchor.row + 1 # Convert XML 0-indexed to 1-indexed to match CellData
            if r not in rows:
                rows[r] = []
            # Use Shape Name or ID as signature
            rows[r].append(f"SHP:{s.name}")

        signatures = []
        # 1-indexed rows
        for r in range(1, max_row + 1):
            if r in rows:
                # Sort values to ensure order doesn't matter?
                # Cells usually ordered. Shapes might not be.
                # Let's sort the shape entries at least? 
                # Actually, simple join is fine if order is specific. 
                # But safer to just join.
                row_vals = "|".join(rows[r])
                signatures.append(row_vals)
            else:
                signatures.append("") # Empty row
        return signatures
