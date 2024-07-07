import unittest
from unittest.mock import patch, MagicMock
from mail import Mail
from gsheets import GoogleSheets
import os


class TestMail(unittest.TestCase):

    @patch('mail.smtplib.SMTP')
    @patch('mail.os.environ.get')
    def test_send_email_using_mime(self, mock_get_env, mock_smtp):
        mock_get_env.return_value = 'fake_password'
        mock_server = mock_smtp.return_value.__enter__.return_value
        mock_server.sendmail = MagicMock()

        mail = Mail("mariasteliana85@gmail.com")
        mail.send_email_using_mime(
            to_email="maya_dlr03@yahoo.com",
            subject="Acesta e subiectul",
            html_data="<h1>Acesta este body-ul</h1>"
        )

        self.assertTrue(mock_server.sendmail.called, "sendmail method was not called")
        mock_server.sendmail.assert_called_with(
            "mariasteliana85@gmail.com",
            "maya_dlr03@yahoo.com",
            unittest.mock.ANY  # Allow any MIME message
        )


class TestGoogleSheets(unittest.TestCase):

    @patch('gsheets.gspread.authorize')
    @patch('gsheets.Credentials.from_service_account_file')
    def test_get_values(self, mock_credentials, mock_authorize):
        mock_client = mock_authorize.return_value
        mock_excel = mock_client.open_by_key.return_value
        mock_sheet = mock_excel.sheet1
        mock_sheet.get_all_records.return_value = [
            {'Name': 'John Doe', 'Email': 'john@example.com'},
            {'Name': 'Jane Doe', 'Email': 'jane@example.com'}
        ]

        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        credentials_path = "path_to_credentials.json"
        excel_key = "your_excel_key"
        google_sheets = GoogleSheets(scopes, credentials_path, excel_key)

        values = google_sheets.get_values()
        self.assertEqual(values, [
            {'Name': 'John Doe', 'Email': 'john@example.com'},
            {'Name': 'Jane Doe', 'Email': 'jane@example.com'}
        ])


if __name__ == '__main__':
    unittest.main()
