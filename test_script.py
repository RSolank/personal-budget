import unittest
from unittest.mock import patch
import sys
import os
import warnings
import getopt
import datetime
import openpyxl
from tabulate import tabulate
from uuid import uuid4
from personal_budget_main import extract_raw_data_from_excel, process_transactions, consolidate_transactions_month_wise, generate_report_meta, generate_report, show_output

class TestTransactionProcessing(unittest.TestCase):
    
    def setUp(self):
        """Set up a temporary directory and Excel file for testing."""
        self.test_dir = 'test_source_data'
        os.makedirs(self.test_dir, exist_ok=True)
        
        self.test_file = os.path.join(self.test_dir, 'test_transactions.xlsx')
        self.create_test_excel_file()
    
    def tearDown(self):
        """Clean up the temporary test directory and files."""
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        if os.path.exists(self.test_dir):
            os.rmdir(self.test_dir)
    
    def create_test_excel_file(self):
        """Create a dummy Excel file for testing."""
        headers = ["Tran Date", "Chq No", "Particulars", "Category", "Sub-Category", "Debit", "Credit", "Balance", "Init. Br"]
        data = [
            [datetime.datetime(2024, 1, 15), "1234", "Test Detail 1", "Category A", "Sub-1", 100, 50, 50, 1000],
            [datetime.datetime(2024, 2, 20), "5678", "Test Detail 2", "Category B", "Sub-2", 200, 150, 50, 1200]
        ]
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(headers)
        for row in data:
            ws.append(row)
        wb.save(self.test_file)

    def test_extract_raw_data_from_excel(self):
        """Test extracting raw data from an Excel file."""
        raw_data = extract_raw_data_from_excel(self.test_file)
        self.assertIsNotNone(raw_data)
        self.assertEqual(len(raw_data), 2)
        self.assertEqual(raw_data[0][1:], [datetime.datetime(2024, 1, 15), "Test Detail 1", 100, 50])

    def test_process_transactions(self):
        """Test processing transactions into a dictionary."""
        raw_data = [
            ["123456", datetime.datetime(2024, 1, 15), "Test Detail 1", 100, 50],
            ["789012", datetime.datetime(2024, 2, 20), "Test Detail 2", 200, 150]
        ]
        transactions = process_transactions(raw_data)
        self.assertEqual(len(transactions), 2)
        self.assertIn("123456", transactions)
        self.assertEqual(transactions["123456"]["txn_detail"], "Test Detail 1")
    
    def test_consolidate_transactions_month_wise(self):
        """Test consolidating transactions month-wise."""
        transactions = {
            "123456": {"txn_date": "15-Jan-2024", "txn_detail": "Test Detail 1", "debit": 100, "credit": 50},
            "789012": {"txn_date": "20-Feb-2024", "txn_detail": "Test Detail 2", "debit": 200, "credit": 150}
        }
        monthly_summary = consolidate_transactions_month_wise(transactions)
        self.assertEqual(len(monthly_summary), 2)
        self.assertIn("01-2024", monthly_summary)
        self.assertEqual(monthly_summary["01-2024"]["debit"], 100)
    
    @patch('sys.argv', ['personal_budget_main.py', '-t', 'month_top_5', '-s', 'debit', '-o', 'descend', '-c', '5'])
    def test_generate_report_meta(self):
        """Test generating report metadata."""
        report_meta = generate_report_meta()
        self.assertEqual(report_meta["type"], "MT")
        self.assertEqual(report_meta["txn_per_month"], 5)
        self.assertTrue(report_meta["sort_req"])
        self.assertEqual(report_meta["sort_by"], "D")
        self.assertEqual(report_meta["sort_order"], "D")
        self.assertEqual(report_meta["view_count"], 5)
    
    def test_generate_report(self):
        """Test generating a report from transactions and monthly summary."""
        transactions = {
            "123456": {"txn_date": "15-Jan-2024", "txn_detail": "Test Detail 1", "debit": 100, "credit": 50},
            "789012": {"txn_date": "20-Feb-2024", "txn_detail": "Test Detail 2", "debit": 200, "credit": 150}
        }
        monthly_summary = {
            "01-2024": {"month_name": "January, 2024", "txn_keys": ["123456"], "debit": 100, "credit": 50, "surplus_deficit": -50},
            "02-2024": {"month_name": "February, 2024", "txn_keys": ["789012"], "debit": 200, "credit": 150, "surplus_deficit": -50}
        }
        report_meta = {
            "type": "M",
            "sort_req": False,
            "sort_by": "D",
            "sort_order": "A",
            "view_count": 10
        }
        report = generate_report(transactions, monthly_summary, report_meta)
        self.assertEqual(len(report), 2)
        self.assertEqual(report[0][0], "January, 2024")
    
    @patch('builtins.print')
    def test_show_output(self, mock_print):
        """Test showing the output of the report."""
        report = [
            ["January, 2024", 100, 50, -50],
            ["February, 2024", 200, 150, -50]
        ]
        report_meta = {"type": "M", "view_count": 2}
        show_output(report, report_meta)
        mock_print.assert_called_once()

if __name__ == '__main__':
    unittest.main()
