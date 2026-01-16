import unittest
import xlsxwriter
import os
from core.comparator import ExcelComparator
from core.data_types import DiffType

class TestExcelComparator(unittest.TestCase):
    def setUp(self):
        self.file_a = 'test_a.xlsx'
        self.file_b = 'test_b.xlsx'
        self.create_files()

    def tearDown(self):
        if os.path.exists(self.file_a):
            os.remove(self.file_a)
        if os.path.exists(self.file_b):
            os.remove(self.file_b)

    def create_files(self):
        # File A: 
        # A1: Title
        # A2: Row1
        # A3: Row2
        # Shape at B2
        wb = xlsxwriter.Workbook(self.file_a)
        ws = wb.add_worksheet()
        ws.write('A1', 'Title')
        ws.write('A2', 'Row1')
        ws.write('A3', 'Row2') # Row 3
        # Add shape at row 1 (0-indexed) = Row 2
        ws.insert_textbox('B2', 'Box1', {'width': 100, 'height': 50})
        wb.close()

        # File B: 
        # A1: Title
        # A2: INSERTED ROW (New)
        # A3: Row1 (Shifted)
        # A4: Row2 Changed (Shifted + Value Change)
        # Shape at B3 (Shifted)
        wb = xlsxwriter.Workbook(self.file_b)
        ws = wb.add_worksheet()
        ws.write('A1', 'Title')
        ws.write('A2', 'INSERTED ROW')
        ws.write('A3', 'Row1')
        ws.write('A4', 'Row2 Changed')
        # Shape moved to B3 (Row 3, 0-indexed = 2)
        # Original was B2 (Row 1). Mapping should be 1 -> 2.
        ws.insert_textbox('B3', 'Box1', {'width': 100, 'height': 50}) 
        wb.close()

    def test_compare_shifts(self):
        comparator = ExcelComparator(self.file_a, self.file_b)
        diff = comparator.compare()
        
        # Analyze results
        # 1. A2 (Row1) should map to A3 (Row1). Values identical. Should be NO DiffItem for this pair.
        # 2. A3 (Row2) should map to A4 (Row2 Changed). Values diff. Should be CHANGED.
        # 3. New row at 2 (A2) in B. Should be INSERTED.
        
        items = diff.items
        print("\nDiff Items:")
        for i in items:
            print(f"{i.diff_type.value}: {i.item_type} {i.location} ({i.details})")

        # Check Insertion
        inserted = [i for i in items if i.diff_type == DiffType.INSERTED and i.item_type == "Cell" and "INSERTED ROW" in str(i.new_value)]
        self.assertTrue(len(inserted) > 0, "Should detect inserted row")

        # Check Shifted Change
        # Old A3 was 'Row2'. New A4 is 'Row2 Changed'.
        # Mapping 2 -> 3.
        # Should detect change for Coordinate A3 -> A4 (or however we log it)
        changed = [i for i in items if i.diff_type == DiffType.CHANGED and i.old_value == "Row2"]
        self.assertTrue(len(changed) > 0, "Should detect change in shifted row")
        self.assertTrue("Row2 Changed" in str(changed[0].new_value))

        # Check Shifted NO Change
        # Old A2 ('Row1') -> New A3 ('Row1').
        # Should NOT be in diff items as changed.
        unchanged_false_pos = [i for i in items if i.diff_type == DiffType.CHANGED and i.old_value == "Row1"]
        self.assertEqual(len(unchanged_false_pos), 0, "Should NOT report unchanged shifted row as changed")

        # Check Shape
        # Shape 'Box1' moved from B2 (row 1) to B3 (row 2).
        # Shift mapping says 1 -> 2.
        # So comparison checks Shape A (row 1) vs Shape B (row 2).
        # They match. Should report MATCH.
        shape_match = [i for i in items if i.item_type == "Shape" and i.diff_type == DiffType.MATCH]
        self.assertTrue(len(shape_match) > 0, "Shape should match despite shift")

if __name__ == '__main__':
    unittest.main()
