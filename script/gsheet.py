from os import environ

import gspread
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials


class Google_spreadsheet_operation():
    def __init__(self) -> None:
        # Roading Environment variable
        load_dotenv(dotenv_path="../setting/.env")
        # use creds to create a client to interact with the Google Drive API
        self.scope = ['https://spreadsheets.google.com/feeds',
                      'https://www.googleapis.com/auth/drive']
        self.cred = environ['credential_file_path']
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(
            self.cred, self.scope)
        self.client = gspread.authorize(self.creds)
        self.sheet_name = str("高額物品管理システム(DB)")
        # Find a workbook by name and open the first sheet
        self.sheet = self.client.open(self.sheet_name).sheet1

    def sp_insert(self, all_list: list):
        all_list.append("-")
        # Searching empty cell and Return row-number(int)
        index: int = 0      # Initial countor
        while True:
            index += 1
            confirm_cell = self.sheet.row_values(index)
            if len(confirm_cell) == 0:
                self.sheet.insert_row(all_list, index)
                break

    def sp_waste(self, control_num: str):
        # Find row of control_num
        find_cell = self.sheet.find(control_num)
        # Update cell ...
        self.sheet.update_cell(find_cell.row, 8, "済")

    def sp_delete(self, control_num: str):
        # Find row of control_num
        find_cell = self.sheet.find(control_num)
        # Delete row for option(delete_row)
        self.sheet.delete_row(find_cell.row)
