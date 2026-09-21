import unittest

import openpyxl
import pandas as pd

import score_logic as app


class UpdateExcelTests(unittest.TestCase):
    def setUp(self):
        self.workbook = openpyxl.Workbook()
        self.sheet = self.workbook.active
        self.excel_data = pd.DataFrame({'Title': ['Test Song', None]})

    def tearDown(self):
        self.workbook.close()

    def test_updates_failed_result_to_full_combo(self):
        self.sheet['I3'] = 'FL'
        csv_data = pd.DataFrame([{
            'title': 'Test Song',
            'difficulty': 'standard',
            'highScore': 600000,
            'FCCount': 1,
            'APCount': 0,
        }])

        warnings = app.update_excel(self.sheet, self.excel_data, csv_data, 0)

        self.assertEqual(self.sheet['I3'].value, 'FC')
        self.assertEqual(warnings, {})

    def test_does_not_downgrade_existing_result(self):
        self.sheet['J3'] = 'AP'
        csv_data = pd.DataFrame([{
            'title': 'Test Song',
            'difficulty': 'expert',
            'highScore': 0,
            'FCCount': 0,
            'APCount': 0,
        }])

        app.update_excel(self.sheet, self.excel_data, csv_data, 0)

        self.assertEqual(self.sheet['J3'].value, 'AP')

    def test_reports_missing_title(self):
        csv_data = pd.DataFrame([{
            'title': 'Missing Song',
            'difficulty': 'ultimate',
            'highScore': 0,
            'FCCount': 0,
            'APCount': 0,
        }])

        warnings = app.update_excel(self.sheet, self.excel_data, csv_data, 0)

        self.assertEqual(warnings, {'Missing Song': {'ultimate'}})


if __name__ == '__main__':
    unittest.main()
