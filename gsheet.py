import gspread
from oauth2client.service_account import ServiceAccountCredentials
from os import environ


class Google_spreadsheet_operation():
    def __init__(self) -> None:
        # use creds to create a client to interact with the Google Drive API
        self.scope = ['https://spreadsheets.google.com/feeds',
                      'https://www.googleapis.com/auth/drive']
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(
            r'C:\Users\yuuto\OneDrive\VS_code\equipment-management-system\gleaming-tube-364710-204f4f6b3eb9.json', self.scope)
        self.client = gspread.authorize(self.creds)
        self.sheet_name = str("高額物品管理システム(DB)")

    def sp_insert(self,all_list: list):
        # Find a workbook by name and open the first sheet
        sheet = self.client.open(self.sheet_name).sheet1

        # Searching empty cell and Return row-number(int)
        index: int = 0      # Initial countor
        while True:
            index += 1
            confirm_cell = sheet.row_values(index)
            if len(confirm_cell) == 0:
                sheet.insert_row(all_list,index)
                break
    
    def sp_delete(self, delete_row: int):
        # Find a workbook by name and open the first sheet
        sheet = self.client.open(self.sheet_name).sheet1
        # Delete row for option(delete_row)
        sheet.delete_row(delete_row)