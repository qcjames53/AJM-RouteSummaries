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
COL_WIDTH_PER_CHAR = 1.5


# Get default filename
date = date.today()
export_filepath = DEFUALT_OUTPUT_PREFIX + str(date) + ".xlsx"

# Override if command line arg triggered
i = 1
while i < len(sys.argv):
    arg = sys.argv[i]

    if arg == "-o":
        # Throw error if no filepath provided
        i += 1
        if i == len(sys.argv):
            raise ValueError("No filepath provided. Usage: [-o filepath]")

        # Assumed to be safe, grab filepath
        export_filepath = sys.argv[i]

    # Invalid arguments
    else:
        raise ValueError("'" + str(arg) + 
            "' is not recognized as a valid argument")

    # Increment loop
    i += 1


# Create output workbook / sheet
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = SHEET_TITLE

# Create column headers
sheet["A1"] = "SEQUENCE"
sheet.column_dimensions["A"].width = 10
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
sheet.column_dimensions["N"].width = 11

# Save the file
wb.save(export_filepath)
