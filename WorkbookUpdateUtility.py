# Author: Quinn James (qj@quinnjam.es)
# 
# A program for updating an old-format excel workbook into a current-format
# workbook.
# 
# Copies data from a .xls fle into a .xlsx file. Changes numerical times and
# dates into excel-format times and dates
#
# More details about this project can be found in the README file or at:
#   https://github.com/qcjames53/AJM-RouteSummaries

import win32com.client

# CONSTANTS
DEBUG = True    # Whether Excel should be visible when updating document format

class ExcelApplication:
    def __init__(self):
        self.app = win32com.client.Dispatch('Excel.Application')
        self.app.visible = DEBUG

    def __del__(self):
        '''
        Closes the Excel application whenever this object is deleted. Used for
        convenience and in case of program crash.
        '''
        self.app.Quit()

    def convertToXLSX(self, old_filepath, new_filepath):
        # Open the old book
        wb = self.app.WorkBooks.Open(old_filepath)

        # Save the new book
        self.app.DisplayAlerts = False
        wb.SaveAs(new_filepath)
        wb.Close()

def convertWorkbook(old_filepath, new_filepath):
    # Update the document format from .xls to .xlsx
    excelApp = ExcelApplication()
    excelApp.convertToXLSX(old_filepath, new_filepath)

