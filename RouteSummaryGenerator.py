# Author: Quinn James (qj@quinnjam.es)
#
# A command-line utility to generate a ridership summary for bus routes.
# Functionality of this program's input and output files are documented in the
# main README.md file.
#
# More details about this project can be found in the README file or at:
#   https://github.com/qcjames53/AJM-RouteSummaries

import openpyxl
import datetime

# Constants
LOG_SHEET_TITLE = "Log"

class Log:
    """
    A log used to display important messages to the end user.
    """

    def __init__(self, log_sheet):
        """
        Initialize the log with the proper local variables.

        @param log_sheet The sheet containing the log
        """
        self.log_sheet = log_sheet
        self.log_row = 1
        self.creation_time = datetime.datetime.now()

        # Make headers
        self.log_sheet.column_dimensions["A"].width = 14
        self.log_sheet.column_dimensions["B"].width = 200
        self.log_sheet["A1"] = "Elapsed time"
        self.log_sheet["B1"] = "Message"
    

    def logMessage(self, message):
        """
        Logs a message to the log sheet.

        @param message A string representing the message to log.
        """
        time_elapsed = datetime.datetime.now() - self.creation_time
        self.log_sheet["A" + str(self.log_row)] = str(time_elapsed)
        self.log_sheet["B" + str(self.log_row)] = message
        self.log_row += 1


def generateSummary(ride_checks_filepath, route_info_filepath, 
    output_filepath):
    """
    Creates an output workbook from the provided input workbooks. See the README
    for more information on how this function operates.

    @param ride_checks_filepath The filepath for the ridechecks workbook
    @param route_info_filepath The filepath for the route info workbook
    @param output_filepath The filepath for the output workbook

    @returns An integer representing the status of the output workbook:
        0 - OK, success or minor errors
        1 - Major error. Check the workbook log for details.
        2 - Output workbook could not be created
    """

    # Create output workbook / sheet
    wb = openpyxl.Workbook()

    # Set up debug/message log
    log_sheet = wb.active
    log_sheet.title = LOG_SHEET_TITLE
    log = Log(log_sheet)
    log.logMessage("Document created at " + str(datetime.datetime.now()))

    # Try to open the ride checks file, if can't return major error
    try:
        ride_checks_wb = openpyxl.load_workbook(filename=ride_checks_filepath)
    except Exception as e:
        log.logMessage("[ERROR] Could not open the ride checks workbook '" +\
            ride_checks_filepath + "'")
        
        # Try to save the output file
        try:
            wb.save(output_filepath)
        except Exception as e:
            return 2
        return 1
    ride_checks = ride_checks_wb.active

    # Try to open the route info file, if can't return major error
    try:
        route_info_wb = openpyxl.load_workbook(filename=route_info_filepath).active
    except Exception as e:
        log.logMessage("[ERROR] Could not open the route info workbook '" +\
            route_info_filepath + "'")
        
        # Try to save the output file
        try:
            wb.save(output_filepath)
        except Exception as e:
            return 2
        return 1

    # start parsing the ride checks file
    current_row = 2
    while(ride_checks.cell(row=current_row, column=1).value is not None):
        print(str(current_row) + ": " + str(ride_checks.cell(row=current_row, column=1).value))

        # get data
        sequence = ride_checks.cell(row=current_row, column=1).value
        date = ride_checks.cell(row=current_row, column=2).value
        route = ride_checks.cell(row=current_row, column=3).value
        direction = ride_checks.cell(row=current_row, column=4).value
        run = ride_checks.cell(row=current_row, column=5).value
        start_time = ride_checks.cell(row=current_row, column=6).value
        onboard = ride_checks.cell(row=current_row, column=7).value
        stop_number = ride_checks.cell(row=current_row, column=8).value
        arrival_time = ride_checks.cell(row=current_row, column=9).value
        schedule_time = ride_checks.cell(row=current_row, column=10).value
        offs = ride_checks.cell(row=current_row, column=11).value
        ons = ride_checks.cell(row=current_row, column=12).value
        loads = ride_checks.cell(row=current_row, column=13).value
        time_check = ride_checks.cell(row=current_row, column=14).value

        # increment current row
        current_row += 1

    log.logMessage("Generation complete")

    # Try to save the output file
    try:
        wb.save(output_filepath)
    except Exception as e:
        return 2

    # Return successfully
    return 0
    
