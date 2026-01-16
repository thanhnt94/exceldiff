import xlwings as xw
from core.data_types import DiffResult, DiffType
import os

class VisualReporter:
    def __init__(self, file_a, file_b, output_path, sheet_a=None, sheet_b=None):
        self.file_a = os.path.abspath(file_a)
        self.file_b = os.path.abspath(file_b)
        self.output_path = os.path.abspath(output_path)
        self.sheet_a = sheet_a
        self.sheet_b = sheet_b
        
    def generate(self, result: DiffResult):
        app = xw.App(visible=False)
        try:
            wb_a = app.books.open(self.file_a)
            wb_b = app.books.open(self.file_b)
            
            # Determine Sheets to use
            ws_a = wb_a.sheets[self.sheet_a] if self.sheet_a and self.sheet_a != "Select File First" else wb_a.sheets.active
            ws_b = wb_b.sheets[self.sheet_b] if self.sheet_b and self.sheet_b != "Select File First" else wb_b.sheets.active
            
            # Create new workbook for output
            wb_out = app.books.add()
            
            # Delete default sheet if exists
            while wb_out.sheets.count > 1:
                wb_out.sheets[-1].delete()
            
            # ========== SHEET 1: Original (copy of Base) ==========
            ws_a.api.Copy(Before=wb_out.sheets[0].api)
            ws_original = wb_out.sheets[0]
            ws_original.name = "Original"
            
            # ========== SHEET 2: Modified (copy of Modified) ==========
            ws_b.api.Copy(After=ws_original.api)
            ws_modified = wb_out.sheets[1]
            ws_modified.name = "Modified"
            
            # ========== SHEET 3: Modified Diff (copy of Modified with highlights) ==========
            ws_b.api.Copy(After=ws_modified.api)
            ws_mod_diff = wb_out.sheets[2]
            ws_mod_diff.name = "Modified Diff"
            
            # ========== SHEET 4: Base Diff (copy of Base with highlights for deletions) ==========
            ws_a.api.Copy(After=ws_mod_diff.api)
            ws_base_diff = wb_out.sheets[3]
            ws_base_diff.name = "Base Diff"
            
            # ========== SHEET 5: Unchanged (copy of Base, will remove changed items) ==========
            ws_a.api.Copy(After=ws_base_diff.api)
            ws_unchanged = wb_out.sheets[4]
            ws_unchanged.name = "Unchanged"
            
            # ========== SHEET 6: Legend - Skip for now, create via API later ==========
            # Legend sheet will be added manually after core functionality works
            
            # Delete the default blank sheet if still exists
            for sheet in wb_out.sheets:
                if sheet.name not in ["Original", "Modified", "Modified Diff", "Base Diff", "Unchanged"]:
                    try:
                        sheet.delete()
                    except:
                        pass
            
            # Colors
            color_inserted = (144, 238, 144) # Light Green (rows/cells added)
            color_changed = (255, 255, 224) # Light Yellow (cells changed)
            color_deleted = (255, 182, 193) # Light Red/Pink (rows/cells deleted)
            
            # 4. Iterate Differences
            # Shape Color Codes (VBA RGB values)
            COLOR_RED = 255        # Position changed
            COLOR_BLUE = 16711680  # Size changed
            COLOR_GREEN = 65280    # Deleted shape
            COLOR_ORANGE = 42495   # Inserted shape
            
            import re
            
            for item in result.items:
                if item.item_type == "Cell":
                    # ===== MODIFIED DIFF: Show changes/insertions in Modified =====
                    if item.diff_type == DiffType.INSERTED:
                        if "Row inserted" in item.details:
                            match = re.search(r'\d+', item.location)
                            if match:
                                r = int(match.group())
                                try:
                                    ws_mod_diff.range(f"{r}:{r}").color = color_inserted
                                except: pass
                        else:
                            try:
                                ws_mod_diff.range(item.location).color = color_inserted
                            except: pass

                    elif item.diff_type == DiffType.CHANGED:
                        parts = item.location.split("->")
                        loc_new = parts[-1].strip()
                        loc_old = parts[0].strip() if len(parts) > 1 else loc_new
                        # Highlight in Modified Diff
                        try:
                            ws_mod_diff.range(loc_new).color = color_changed
                            ws_mod_diff.range(loc_new).api.AddComment(f"Was: {item.old_value}")
                        except: pass
                        # Highlight in Base Diff
                        try:
                            ws_base_diff.range(loc_old).color = color_changed
                            ws_base_diff.range(loc_old).api.AddComment(f"Now: {item.new_value}")
                        except: pass
                    
                    # ===== BASE DIFF: Show deletions from Base =====
                    elif item.diff_type == DiffType.DELETED:
                        if "Row deleted" in (item.details or ""):
                            match = re.search(r'\d+', item.location)
                            if match:
                                r = int(match.group())
                                try:
                                    ws_base_diff.range(f"{r}:{r}").color = color_deleted
                                except: pass
                        else:
                            try:
                                ws_base_diff.range(item.location).color = color_deleted
                            except: pass
                            
                elif item.item_type == "Shape":
                    details = item.details if item.details else ""
                    
                    if item.diff_type in [DiffType.CHANGED, DiffType.MOVED]:
                        # Highlight changed/moved shape in Modified Diff
                        position_changed = "position_changed" in details
                        size_changed = "size_changed" in details
                        
                        try:
                            shape_mod = ws_mod_diff.shapes[item.location]
                            if position_changed:
                                shape_mod.api.Line.ForeColor.RGB = COLOR_RED
                            elif size_changed:
                                shape_mod.api.Line.ForeColor.RGB = COLOR_BLUE
                            else:
                                shape_mod.api.Line.ForeColor.RGB = COLOR_RED
                            shape_mod.api.Line.Weight = 2.5
                            shape_mod.api.Line.Visible = True
                        except Exception as e:
                            print(f"DEBUG: Shape style failed: {e}")
                            
                    elif item.diff_type == DiffType.DELETED:
                        # Highlight deleted shape in Base Diff (it exists there)
                        try:
                            shape_del = ws_base_diff.shapes[item.location]
                            shape_del.api.Line.ForeColor.RGB = COLOR_GREEN
                            shape_del.api.Line.Weight = 2.5
                            shape_del.api.Line.Visible = True
                        except Exception as e:
                            print(f"DEBUG: Deleted shape style failed: {e}")
                     
                    elif item.diff_type == DiffType.INSERTED:
                        # Highlight inserted shape in Modified Diff
                        try:
                            shape_new = ws_mod_diff.shapes[item.location]
                            shape_new.api.Line.ForeColor.RGB = COLOR_ORANGE
                            shape_new.api.Line.Weight = 2.5
                            shape_new.api.Line.Visible = True
                        except Exception as e:
                            print(f"DEBUG: Inserted shape style failed: {e}")
                      
                    elif item.diff_type == DiffType.MATCH:
                        # Remove unchanged shapes from BOTH diff sheets
                        try:
                            ws_mod_diff.shapes[item.location].api.Delete()
                        except: pass
                        try:
                            ws_base_diff.shapes[item.location].api.Delete()
                        except: pass
            
            # Save output
            wb_out.save(self.output_path)
            wb_a.close()
            wb_b.close()
            wb_out.close()
            
        finally:
            app.quit()

