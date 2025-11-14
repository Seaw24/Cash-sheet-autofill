import unittest
from csv_parser import CSVParser

class TestReadFunction(unittest.TestCase): 
    #CC
    csv_report = CSVParser(r"C:\Users\admin\OneDrive\Nam career\Chartwells\Auto-fill casheet\reports 22\Operations Report-Crimson Corner 1-1-(Oct 22 2025).csv")

    def test_read_location(self): 
        self.csv_report.read_csv()
        self.csv_report.parse_location()
        self.assertEqual(self.csv_report.data["location"] ,"Crimson Corner 1-1")