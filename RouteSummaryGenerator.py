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

    # Create output workbook / sheet
    wb = openpyxl.Workbook()

    # Set up debug/message log
    log_sheet = wb.active
    log_sheet.title = LOG_SHEET_TITLE
    log = Log(log_sheet)
    log.logMessage("Document created at " + str(datetime.datetime.now()))

    # delay
    for i in range(0, 1000000):
        i += 1
        i -= 1

    # Save the output file
    log.logMessage("Generation complete")
    wb.save(output_filepath)
    
