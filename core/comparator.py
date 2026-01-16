from typing import List, Dict, Optional
from .excel_loader import ExcelLoader
from .shift_detector import ShiftDetector
from .data_types import CellData, ShapeData, DiffResult, DiffItem, DiffType, AnchorPoint

class ExcelComparator:
    def __init__(self, file_a: str, file_b: str, sheet_a: str = None, sheet_b: str = None):
        self.file_a = file_a
        self.file_b = file_b
        self.loader_a = ExcelLoader(file_a, sheet_name=sheet_a)
        self.loader_b = ExcelLoader(file_b, sheet_name=sheet_b)
        self.detector = ShiftDetector()

    def compare(self) -> DiffResult:
        cells_a, shapes_a = self.loader_a.load()
        cells_b, shapes_b = self.loader_b.load()
        
        # DEBUG: Print shapes from each file
        print(f"DEBUG: Shapes in File A (Base): {[s.name for s in shapes_a]}")
        print(f"DEBUG: Shapes in File B (Modified): {[s.name for s in shapes_b]}")
        
        # 1. Compute Mappings
        # Determine max rows/cols
        max_row_a = max([c.row for c in cells_a] + [s.from_anchor.row for s in shapes_a]) if cells_a or shapes_a else 0
        max_row_b = max([c.row for c in cells_b] + [s.from_anchor.row for s in shapes_b]) if cells_b or shapes_b else 0
        
        sigs_a = self.detector.get_row_signatures(cells_a, shapes_a, max_row_a)
        sigs_b = self.detector.get_row_signatures(cells_b, shapes_b, max_row_b)
        
        row_mapping = self.detector.compute_mapping(sigs_a, sigs_b)
        
        # Col mapping (Simplified: assuming column letters match for now, 
        # or we could do transpose signatures. Let's start with just Row mapping)
        
        diff_items = []
        
        # 2. Compare Cells
        # Index cells by (row, col)
        map_a = {(c.row, c.col): c for c in cells_a}
        map_b = {(c.row, c.col): c for c in cells_b}
        
        # Check A against B
        for (r_a, c_a), cell_a in map_a.items():
            # Get expected row index in B (0-indexed in mapping + 1 for Excel)
            # mapping keys are 0-indexed. r_a is 1-indexed.
            idx_a = r_a - 1
            idx_b = row_mapping.get(idx_a)
            
            if idx_b is None:
                # Row deleted
                diff_items.append(DiffItem(
                    location=cell_a.coordinate,
                    item_type="Cell",
                    diff_type=DiffType.DELETED,
                    old_value=cell_a.value
                ))
            else:
                r_b = idx_b + 1
                # Check if this cell exists in B at (r_b, c_a) (assuming cols check out)
                cell_b = map_b.get((r_b, c_a))
                
                if cell_b:
                    if str(cell_a.value) != str(cell_b.value):
                        diff_items.append(DiffItem(
                            location=f"{cell_a.coordinate} -> {cell_b.coordinate}",
                            item_type="Cell",
                            diff_type=DiffType.CHANGED,
                            old_value=cell_a.value,
                            new_value=cell_b.value
                        ))
                else:
                     diff_items.append(DiffItem(
                        location=cell_a.coordinate,
                        item_type="Cell",
                        diff_type=DiffType.DELETED, # Start cell gone (or moved)
                        old_value=cell_a.value,
                        details=f"Mapped to row {r_b} but cell empty"
                    ))

        # Check for Insertions in B (reverse check)
        mapped_rows_b = set(v for v in row_mapping.values() if v is not None)
        for (r_b, c_b), cell_b in map_b.items():
            idx_b = r_b - 1
            if idx_b not in mapped_rows_b:
                diff_items.append(DiffItem(
                    location=cell_b.coordinate,
                    item_type="Cell",
                    diff_type=DiffType.INSERTED,
                    new_value=cell_b.value,
                    details="Row inserted"
                ))
            else:
                # If row mapped, we already checked A vs B. 
                # Unless B has a cell where A was empty.
                # Find reverse mapping? Or just iterate all B cells in mapped rows.
                # Find (r_a) s.t. map[r_a] == r_b
                # Slow but works:
                original_row_idx = next((k for k, v in row_mapping.items() if v == idx_b), None)
                if original_row_idx is not None:
                     r_a = original_row_idx + 1
                     if (r_a, c_b) not in map_a:
                         diff_items.append(DiffItem(
                            location=cell_b.coordinate,
                            item_type="Cell",
                            diff_type=DiffType.INSERTED,
                            new_value=cell_b.value,
                            details="Cell added in existing row"
                        ))

        # 3. Compare Shapes
        # Shape similarity could be based on ID (unreliable?) or Text/Content + Relative Pos
        # Let's try matching by ID first, then fallback to property matching?
        # IDs often change if copy-pasted.
        # Let's match by "Adjusted Position" + "Type".
        # Create lookups for shapes in B
        shapes_b_by_id = {s.id: s for s in shapes_b}
        shapes_b_by_name = {s.name: s for s in shapes_b}
        
        matched_shapes_b = set()
        
        for sa in shapes_a:
            found_match = None
            
            # PRIORITY 1: Match by ID (most reliable)
            if sa.id in shapes_b_by_id and shapes_b_by_id[sa.id] not in matched_shapes_b:
                found_match = shapes_b_by_id[sa.id]
            
            # PRIORITY 2: Match by NAME (fallback if ID not found)
            if found_match is None:
                if sa.name in shapes_b_by_name and shapes_b_by_name[sa.name] not in matched_shapes_b:
                    found_match = shapes_b_by_name[sa.name]
            
            # PRIORITY 3: Match by POSITION (last resort)
            if found_match is None:
                idx_b = row_mapping.get(sa.from_anchor.row)
                if idx_b is not None:
                    for sb in shapes_b:
                        if sb in matched_shapes_b: continue
                        if sb.from_anchor.row == idx_b and sb.from_anchor.col == sa.from_anchor.col:
                            found_match = sb
                            break
            
            if found_match:
                matched_shapes_b.add(found_match)
                # Compare contents
                reason = []
                position_changed = False
                size_changed = False
                
                if sa.text != found_match.text:
                    reason.append("text_changed")
                if sa.type_name != found_match.type_name:
                    reason.append("type_changed")
                    
                # Check Offsets (EMUs tolerance, e.g. 10000 approx 1mm?)
                # 360000 EMUs = 1 cm.
                TOLERANCE = 10000
                if abs(sa.from_anchor.row_off - found_match.from_anchor.row_off) > TOLERANCE:
                    position_changed = True
                if abs(sa.from_anchor.col_off - found_match.from_anchor.col_off) > TOLERANCE:
                    position_changed = True
                    
                # SIZE Comparison: Compare to_anchor if available
                if sa.to_anchor and found_match.to_anchor:
                    # Calculate Size (in cell units, rough approximation)
                    sa_width = sa.to_anchor.col - sa.from_anchor.col
                    sa_height = sa.to_anchor.row - sa.from_anchor.row
                    fb_width = found_match.to_anchor.col - found_match.from_anchor.col
                    fb_height = found_match.to_anchor.row - found_match.from_anchor.row
                    
                    if sa_width != fb_width or sa_height != fb_height:
                        size_changed = True
                    
                    # Also check sub-cell offsets for size
                    # Width in EMU: to_col_off - from_col_off (within anchor cells)
                    # This is more precise
                    sa_emu_width = sa.to_anchor.col_off - sa.from_anchor.col_off
                    fb_emu_width = found_match.to_anchor.col_off - found_match.from_anchor.col_off
                    if abs(sa_emu_width - fb_emu_width) > TOLERANCE:
                        size_changed = True
                        
                    sa_emu_height = sa.to_anchor.row_off - sa.from_anchor.row_off
                    fb_emu_height = found_match.to_anchor.row_off - found_match.from_anchor.row_off
                    if abs(sa_emu_height - fb_emu_height) > TOLERANCE:
                        size_changed = True
                
                if position_changed:
                    reason.append("position_changed")
                if size_changed:
                    reason.append("size_changed")
                    
                if reason:
                    diff_items.append(DiffItem(
                        location=sa.name,
                        item_type="Shape",
                        diff_type=DiffType.CHANGED,
                        old_value=sa.text,
                        new_value=found_match.text,
                        details=",".join(reason)
                    ))
                else:
                    # Match
                    diff_items.append(DiffItem(
                        location=sa.name,
                        item_type="Shape",
                        diff_type=DiffType.MATCH
                    ))
            else:
                # Check if it was moved (exists in B but diff pos)
                # Or deleted
                diff_items.append(DiffItem(
                    location=sa.name,
                    item_type="Shape",
                    diff_type=DiffType.DELETED,
                    details="Or moved significantly"
                ))

        # Check for inserted shapes in B
        inserted_shapes = []
        for sb in shapes_b:
            if sb not in matched_shapes_b:
                inserted_shapes.append(DiffItem(
                    location=sb.name,
                    item_type="Shape",
                    diff_type=DiffType.INSERTED,
                    new_value=sb.text
                ))
        
        # DEBUG
        print(f"DEBUG: matched_shapes_b names: {[s.name for s in matched_shapes_b]}")
        print(f"DEBUG: inserted_shapes: {[s.location for s in inserted_shapes]}")

        # 4. Post-Process: Link DELETED and INSERTED shapes by Name/ID to detect Moves (Anchor changes)
        final_items = []
        deleted_map = {} # name -> item
        
        # Filter out DELETED shapes temporarily
        for item in diff_items:
            if item.item_type == "Shape" and item.diff_type == DiffType.DELETED:
                deleted_map[item.location] = item
            else:
                final_items.append(item)
                
        # Check against INSERTED
        for item in inserted_shapes:
            if item.location in deleted_map:
                # Pair found! It was Deleted (anchor lost) and Inserted (new anchor) -> MOVED
                del_item = deleted_map.pop(item.location)
                
                # Compare text to see if that changed too
                details = "Anchor changed (Moved)"
                if del_item.old_value != item.new_value: # text stored in old/new_value
                    details += ", Text changed"
                    
                final_items.append(DiffItem(
                    location=item.location,
                    item_type="Shape",
                    diff_type=DiffType.MOVED, # Explicitly Moved
                    old_value=del_item.old_value,
                    new_value=item.new_value,
                    details=details
                ))
            else:
                final_items.append(item)
                
        # Add remaining deleted
        for item in deleted_map.values():
            final_items.append(item)
            
        return DiffResult(items=final_items)
