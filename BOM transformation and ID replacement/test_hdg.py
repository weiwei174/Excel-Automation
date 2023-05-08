import unittest
from copy import search_for_bolt_id, replace_bolt_ids

class TestSearchBoltId(unittest.TestCase):
    
    def test_search_bolt(self):
        row = search_for_bolt_id(['C', 'D', 'E'], 498782)
        self.assertEqual(row, 59)
        row = search_for_bolt_id(['C', 'D', 'E'], 555555)
        self.assertEqual(row, 0)

    def test_replace_bolt_id(self):
        ID = replace_bolt_ids({3: 405784.0,	4: 498782, 5:"406158"}, '')[0]   
        self.assertEqual(ID, "405784 / 406158")

        empty_comment = replace_bolt_ids({3: 405784.0,	4: 498782, 5:"406158"}, '')[1]
        comment = replace_bolt_ids({3: 405773.0, 4:406149.0}, '')[1]
        self.assertEqual(empty_comment, "")
        self.assertEqual(comment, "406149")
 
if __name__ == '__main__':
    unittest.main()