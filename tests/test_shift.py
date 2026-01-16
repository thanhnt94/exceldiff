import unittest
from core.shift_detector import ShiftDetector
from core.data_types import CellData

class TestShiftDetector(unittest.TestCase):
    def test_simple_insertion(self):
        detector = ShiftDetector()
        # A: [A, B, C]
        # B: [A, INS, B, C]
        list_a = ['A', 'B', 'C']
        list_b = ['A', 'INS', 'B', 'C']
        
        mapping = detector.compute_mapping(list_a, list_b)
        
        # Expected: 0->0, 1->2, 2->3
        self.assertEqual(mapping[0], 0)
        self.assertEqual(mapping[1], 2)
        self.assertEqual(mapping[2], 3)
        
    def test_deletion(self):
        detector = ShiftDetector()
        # A: [A, B, C]
        # B: [A, C]
        list_a = ['A', 'B', 'C']
        list_b = ['A', 'C']
        
        mapping = detector.compute_mapping(list_a, list_b)
        
        # Expected: 0->0, 1->None, 2->1
        self.assertEqual(mapping[0], 0)
        self.assertIsNone(mapping[1])
        self.assertEqual(mapping[2], 1)

    def test_row_signatures(self):
        detector = ShiftDetector()
        cells = [
            CellData(row=1, col=1, value="Head", coordinate="A1"),
            CellData(row=2, col=1, value="Val1", coordinate="A2"),
            CellData(row=2, col=2, value="Val2", coordinate="B2"),
        ]
        # Max row 3 (implies row 3 is empty)
        sigs = detector.get_row_signatures(cells, 3)
        self.assertEqual(len(sigs), 3)
        self.assertEqual(sigs[0], "Head")
        self.assertEqual(sigs[1], "Val1|Val2")
        self.assertEqual(sigs[2], "")

if __name__ == '__main__':
    unittest.main()
