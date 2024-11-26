import unittest
from unittest.mock import patch, MagicMock
from broadcast.addin import xlAddin
from datetime import datetime

class TestXlAddin(unittest.TestCase):

    @patch("broadcast.addin.xw.App")  # Mocking the xlwings App
    def test_initiate(self, mock_app):
        # Mock the app object and RegisterXLL method
        mock_instance = mock_app.return_value
        mock_instance.api.RegisterXLL = MagicMock()

        # Initialize the xlAddin with a mock path
        addin = xlAddin(addin_path="mock/path/to/addin")
        result = addin.initiate()

        # Assert that the app object is returned
        self.assertEqual(result, mock_instance)
        mock_instance.api.RegisterXLL.assert_called_once_with("mock/path/to/addin")

    @patch("broadcast.addin.time.time")  # Mocking the time module
    @patch("broadcast.addin.xw.Range")  # Mocking the xlwings Range
    def test_wait(self, mock_range, mock_time):
        # Mock a range object
        mock_range.value = "#BCONN"
        mock_time.side_effect = [0, 1, 2, 3, 4,5]  # Mock time progression

        # Pass the addin_path argument to avoid the ValueError
        addin = xlAddin(visible=True, addin_path="mock/path/to/addin")

        # Call the wait method and assert it loops correctly
        addin.wait(mock_range, timeout=5)
        self.assertEqual(mock_time.call_count, 6)

    def test_convert_excel_date(self):
        addin = xlAddin(addin_path="mock/path/to/addin")
        # Test with a valid Excel date
        excel_date = 45239  # Corresponds to 2023-12-15
        result = addin.convert_excel_date(excel_date)
        self.assertEqual(result, datetime(2023, 11, 9).date())

        # Test with None
        result = addin.convert_excel_date(None)
        self.assertIsNone(result)

    def test_datetime_to_excel_date(self):
        addin = xlAddin(addin_path="mock/path/to/addin")
        # Test with a valid datetime string
        result = addin.datetime_to_excel_date("2024-01-05")
        expected = 45296.0  # Excel date for 2024-11-26
        self.assertAlmostEqual(result, expected, places=1)

        # Test with invalid format (should raise ValueError)
        with self.assertRaises(ValueError):
            addin.datetime_to_excel_date("invalid-date")

if __name__ == "__main__":
    unittest.main()
