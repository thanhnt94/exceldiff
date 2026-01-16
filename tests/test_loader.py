import unittest
import os
from core.excel_loader import ExcelLoader

class TestExcelLoader(unittest.TestCase):
    def setUp(self):
        self.filename = 'test_shapes.xlsx'
        # Ensure the file exists (created by PoC)
        if not os.path.exists(self.filename):
            raise FileNotFoundError("Run poc_shapes.py first or ensure test_shapes.xlsx exists")

    def test_load_cells(self):
        loader = ExcelLoader(self.filename)
        cells, _ = loader.load()
        
        # Check basic cell content
        # B2 has 'This is a textbox' but it's a shape! Wait, create_excel_with_shapes
        # wrote 'A1' -> 'Hello', 'C3' -> 'World'
        val_map = {c.coordinate: c.value for c in cells}
        self.assertEqual(val_map.get('A1'), 'Hello')
        self.assertEqual(val_map.get('C3'), 'World')

    def test_load_shapes(self):
        loader = ExcelLoader(self.filename)
        _, shapes = loader.load()
        
        print(f"\nLoaded {len(shapes)} shapes")
        for s in shapes:
            print(s)
            
        # We expect 2 textboxes
        # B2 is row 1, col 1 (0-indexed) or row 2, col 2 (1-indexed).
        # OpenPyxl 1-indexed. XML often 0-indexed.
        # Let's see what the loader returns.
        
        self.assertTrue(len(shapes) >= 2)
        
        # Check for our known textboxes
        # B2 textbox
        # xdr:col should be 1 (B), xdr:row should be 1 (2)
        shape_b2 = next((s for s in shapes if s.from_anchor.col == 1 and s.from_anchor.row == 1), None)
        self.assertIsNotNone(shape_b2, "Could not find B2 shape (col=1, row=1)")
        self.assertTrue("This is a textbox" in shape_b2.text, f"Text mismatch: {shape_b2.text}")

        # E5 textbox (E=4, 5=4) -> col=4, row=4
        shape_e5 = next((s for s in shapes if s.from_anchor.col == 4 and s.from_anchor.row == 4), None)
        self.assertIsNotNone(shape_e5, "Could not find E5 shape (col=4, row=4)")
        self.assertTrue("Another shape" in shape_e5.text)

if __name__ == '__main__':
    unittest.main()
