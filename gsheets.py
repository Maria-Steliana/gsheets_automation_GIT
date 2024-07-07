import json
import logging
import gspread
from google.oauth2.service_account import Credentials

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s')
class GoogleSheets:

    def __init__(self, scopes, credentials_path, excel_key):
        self.scopes = scopes
        self.creds = Credentials.from_service_account_file(credentials_path, scopes=self.scopes)
        self.client = gspread.authorize(self.creds)
        self.excel = self.client.open_by_key(excel_key)
        self.working_sheet = self.excel.sheet1

    def insert_row(self, row_values):
        num_existing_rows = len(self.working_sheet.get_all_values()) # imi ia cate randuri sunt in gsheet
        # calculeaza indexul pentru noul rand astfel incat,
        # atunci cand introduc multiple randuri, sa mi le puna unele sub altele
        index = num_existing_rows + 1 if num_existing_rows >= 4 else 4
        self.working_sheet.insert_row(row_values, index)
        logging.info(f"Now rând inserat la pozitia {index} cu valorile {row_values}")

    def update_cell(self, row, col, value):
        self.working_sheet.update_cell(row, col, value)
        logging.info(f"Celula de la randul {row}, coloana {col} are noua valoare: {value}")

    def add_worksheet(self, title, rows, cols):
        new_sheet = self.excel.add_worksheet(title=title, rows=rows, cols=cols)
        logging.info(f"Worksheet '{title}' added with {rows} rows and {cols} columns")
        return new_sheet

    def get_values(self):
        return self.working_sheet.get_all_records()

    def bold_first_row(self):
        self.working_sheet.format('A1:F1', {'textFormat': {'bold': True}})


# scopes = [
#     "https://www.googleapis.com/auth/spreadsheets"
# ]
# creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
#
# client = gspread.authorize(creds)
#
# full_excel = client.open_by_key("1jdbZ30JGxarzDVxOXrpeev6QtP5MecrMaifDIKXVy_A")
# working_sheet = full_excel.sheet1
# values = working_sheet.get_all_records()
# print(values)


if __name__ == '__main__':

    with open("config.json", "r") as f:
        gsheets_config = json.loads(f.read())

    excel = GoogleSheets(scopes=gsheets_config["scopes"],
                         credentials_path="credentials.json",
                         excel_key=gsheets_config["excel_id"])
    employees = excel.get_values()
    excel.bold_first_row()

    while True:
        choice = input("""Choose an option: 
        1. Insert a new row
        2. Update a cell
        3. Add a worksheet
        4. Quit
        """)

        match choice:
            case "1":
                name = input("Name: ")
                mail = input("Email: ")
                department = input("Department: ")
                manager_name = input("Manager Name: ")
                manager_email = input("Manager Email: ")
                password_expiration_date = input("Password Expiration Date (dd/mm/yyyy): ")
                excel.insert_row([name, mail, department, manager_name, manager_email, password_expiration_date])
            case "2":
                row = int(input("Enter Row Number: "))
                col = int(input("Enter Column Number: "))
                value = input("Enter New Value: ")
                excel.update_cell(row, col, value)
                # la functia asta se poate modifica coloana, astfel incat sa fie introdusa ca litera, nu cifra/numar
                # coloanele sunt indexate cu litere mari
                # o sa ma gandesc mai tarziu
            case "3":
                title = input("Enter Worksheet Title: ")
                rows = input("Enter Number of Rows: ")
                cols = input("Enter Number of Columns: ")
                excel.add_worksheet(title, rows, cols)
            case "4":
                print("Exiting...")
                break
            case _:
                print("Invalid option. Please try again.")

    logging.info(f"Lista angajati: {employees}")





