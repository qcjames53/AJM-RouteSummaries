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
from enum import Enum

# Constants
LOG_SHEET_TITLE = "Log"


# Route direction enum
class Direction(Enum):
    IB = "IB" # inbound
    OB = "OB" # outbound
    NB = "NB" # northbound
    SB = "SB" # southbound
    EB = "EB" # eastbound
    WB = "WB" # westbound
    UN = "UN" # unknown


# Utility functions
def routeToRoutestring(route_no, direction):
    """
    A utility function to convert between route no, direction and routestring

    @param route_no An int representing the route number
    @param direction A Direction enum
    @returns A string consisting of the format:
        [direction][route number]
    """
    return direction + str(route_no)


def routestringToRoute(routestring):
    """
    A utility function to convert between routestring and route no, direction

    @param routestring A string consisting of the format:
        [direction][route number]
    @returns A list of the format:
        [route number] - An integer representing the route number
        [direction] - A Direction enum
    """
    route_no = int(routestring[2:])
    dir_string = routestring[:2]
    
    # return route number and direction
    return route_no, stringToDirection(dir_string)


def stringToDirection(dir_string):
    """
    Converts a direction string to a direction enum

    @param dir_string The string to convert
    @returns Direction enum for the direction of this string
    """
    direction = Direction.UN
    if dir_string == "IB":
        direction = Direction.IB
    elif dir_string == "OB":
        direction = Direction.OB
    elif dir_string == "NB":
        direction = Direction.NB
    elif dir_string == "SB":
        direction = Direction.SB
    elif dir_string == "EB":
        direction = Direction.EB
    elif dir_string == "WB":
        direction = Direction.WB
    return direction


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

    
    def logWarning(self, message):
        """
        Logs a warning to the log sheet.

        @param message A string representing the message to log.
        """
        self.logMessage("[WARNING] " + message)


    def logError(self, message):
        """
        Logs an error to the log sheet.

        @param message A string representing the message to log.
        """
        self.logMessage("[ERROR] " + message)


class RouteManager:
    """ 
    A class that allows for easy storage of routes and their respective data
    """

    def __init__(self):
        """
        Initilize RouteManager with the appropriate internal variables
        """
        # A map of routes from routestring to route object
        self.routes = {}


    def __str__(self) -> str:
        output = ""
        for key in self.routes:
            output += self.routes[key].__str__() + '\n'
        return output


    def __repr__(self) -> str:
        return self.__str__()


    def addData(self, routestring, datetime, stop_no):
        # if the routestring key does not exist, create the route
        if routestring not in self.routes:
            self.routes[routestring] = Route(routestring)

        # Add the data to the appropriate route
        self.routes[routestring].addData(datetime, stop_no)


class Route:
    """
    A class that allows for easy storage of the stops inside of a route
    """
    def __init__(self, routestring):
        """
        Initialize Route with the appropriate internal variables
        """
        self.routestring = routestring
        self.departures = {}


    def __str__(self) -> str:
        output = ""
        for key1 in self.departures:
            for key2 in self.departures[key1]:
                output += str(self.routestring) + " " + str(key1) + " "
                output += str(key2) + ": " + str(self.departures[key1][key2])
                output += "\n"
        return output

    
    def __repr__(self) -> str:
        return self.__str__()

    
    def addData(self, datetime, stop_no):
        if datetime not in self.departures:
            self.departures[datetime] = {}

        self.departures[datetime][stop_no] = "temp data"


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

    # Create output workbook / sheet. Set date format to ISO 8601
    wb = openpyxl.Workbook()
    wb.iso_dates = True

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
    route_manager = RouteManager()
    current_row = 2
    prev_seq = 0
    while(ride_checks.cell(row=current_row, column=1).value is not None):
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

        # check that all required data is the proper type
        if not isinstance(sequence, int):
            log.logError("Row " + str(current_row) + ": Sequence '" + \
                str(sequence) + "' is not an integer. Skipping row.")
            current_row += 1
            continue
        if not isinstance(date, datetime.date):
            log.logError("Row " + str(current_row) + ": Date '" + \
                str(date) + "' is not an excel-formatted date. " + \
                "Skipping row.")
            current_row += 1
            continue
        if not isinstance(route, int):
            log.logError("Row " + str(current_row) + ": Route '" + \
                str(route) + "' is not an integer. Skipping row.")
            current_row += 1
            continue
        if stringToDirection(direction) == Direction.UN:
            log.logError("Row " + str(current_row) + ": Direction '" + \
                str(direction) + "' is not a valid input. Skipping row.")
            current_row += 1
            continue
        if (not isinstance(run, int)):
            log.logError("Row " + str(current_row) + ": Run '" + \
                str(run) + "' is not an integer. Skipping row.")
            current_row += 1
            continue
        if (not isinstance(start_time, datetime.time)):
            log.logError("Row " + str(current_row) + ": Start time '" + \
                str(start_time) + "' is not an excel-formatted time. " + \
                "Skipping row.")
            current_row += 1
            continue

        # check that all optional data is the correct format if filled in


        # Check that the sequence number is in order, alert if not
        if sequence - 1 != prev_seq:
            log.logWarning("Out-of-order sequence number: Row " +
            str(current_row))
        prev_seq = sequence

        # add data to the route manager object
        # order of placement is route string -> date and time -> stop_number
        routestring = routeToRoutestring(route, direction)
        date_and_time = datetime.datetime.combine(date, start_time)
        route_manager.addData(routestring, date_and_time, stop_number)

        # increment current row
        current_row += 1

    print(route_manager)

    log.logMessage("Generation complete")

    # Try to save the output file
    try:
        wb.save(output_filepath)
    except Exception as e:
        return 2

    # Return successfully
    return 0
    
