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
import openpyxl

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
        wb.SaveAs(new_filepath, FileFormat = 51, ConflictResolution=2)
        wb.Close()
        self.app.DisplayAlerts = True

def convertFormat(log, old_filepath, new_filepath) -> None:
    '''
    Updates an excel document from .xls to .xlsx format without modifying
    the document contents.

    @param old_filepath The filepath for the document to open
    @param new_filepath The filepath for the document to save
    '''
    # Update the document format from .xls to .xlsx
    excelApp = ExcelApplication()
    excelApp.convertToXLSX(old_filepath, new_filepath)

def convertValues(log, filepath, save_filepath=None) -> None:
    '''
    Updates the values of a .xlsx workbook to match the required format.
    Numerical dates and times become excel-format dates and times.

    @param filepath The filepath of the excel document.
    @param save_filepath Optional: The filepath to save the document to
    '''
    # Load the file
    try:
        wb = openpyxl.load_workbook(filename=filepath)
    except:
        log.logError("Could not load the workbook")
        return
    ride_checks = wb.active

