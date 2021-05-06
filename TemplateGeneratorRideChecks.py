# Date: 2021-05-06
# 
# A program for generating a template ride checks log workbook
#
# Command line options:
# -o <filepath>
#   Specify the output location for the created template. Defaults to
#   'ridechecks[date].xlsx' if not specified.

import sys
from datetime import date
import openpyxl

# Constants
DEFUALT_OUTPUT_PREFIX = "ridechecks"
SHEET_TITLE = "Ride Checks"

def createTemplateRideChecks(filepath=None):
    # Get default filename
    if filepath == None:
        date_string = str(date.today())
        filepath = DEFUALT_OUTPUT_PREFIX + date_string + ".xlsx"

    # Create output workbook / sheet
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = SHEET_TITLE

    # Create column headers
    sheet["A1"] = "SEQUENCE"
    sheet.column_dimensions["A"].width = 11
    sheet["B1"] = "DATE"
    sheet.column_dimensions["B"].width = 10
    sheet["C1"] = "ROUTE"
    sheet.column_dimensions["C"].width = 8
    sheet["D1"] = "DIRECTION"
    sheet.column_dimensions["D"].width = 11
    sheet["E1"] = "RUN"
    sheet.column_dimensions["E"].width = 6
    sheet["F1"] = "START TIME"
    sheet.column_dimensions["F"].width = 12
    sheet["G1"] = "ONBOARD"
    sheet.column_dimensions["G"].width = 10
    sheet["H1"] = "STOP NUMBER"
    sheet.column_dimensions["H"].width = 15
    sheet["I1"] = "ARRIVAL TIME"
    sheet.column_dimensions["I"].width = 15
    sheet["J1"] = "SCHEDULE TIME"
    sheet.column_dimensions["J"].width = 15
    sheet["K1"] = "OFFS"
    sheet.column_dimensions["K"].width = 6
    sheet["L1"] = "ONS"
    sheet.column_dimensions["L"].width = 6
    sheet["M1"] = "LOADS"
    sheet.column_dimensions["M"].width = 8
    sheet["N1"] = "TIME CHECK"
    sheet.column_dimensions["N"].width = 12

    # Save the file
    wb.save(filepath)