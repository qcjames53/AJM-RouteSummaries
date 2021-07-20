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
    def __init__(self, log:Log):
        """
        Initilize RouteManager with the appropriate internal variables
        """
        # A map of routes from routestring to route object
        self.routes = {}
        self.log = log

    def __str__(self) -> str:
        output = ""
        for key in self.routes:
            output += self.routes[key].__str__() + '\n'
        return output

    def __repr__(self) -> str:
        return self.__str__()

    def addStop(self, route, stop_no, street, cross_street):
        # if the route key does not exist, create the route
        if route not in self.routes:
            self.routes[route] = Route(route, self.log)
        
        # Add the stop to the appropriate route
        self.routes[route].addStop(stop_no, street, cross_street)

    def addData(self, route, stop_no, datetime, arrival_time, \
        schedule_time, offs, ons) -> None:
        """
        Adds data to the appropriate internal route object

        @param route A number representing the route
        @param datetime The datetime the route began
        @param stop_no The stop number
        @param arrival_time The arrival time of the stop
        @param schedule_time The scheduled arrival time of the stop
        @param offs The number of passengers departing the bus
        @param ons The number of passengers boarding the bus
        """
        # If the route does not exist, log an error and return
        if route not in self.routes:
            self.log.logError("Tried to add data to nonexistent route: " + str(route))
            return

        # Add the data to the appropriate route
        self.routes[route].addData(stop_no, datetime, arrival_time,\
            schedule_time, offs, ons)

    def buildRouteTotals(self, worksheet) -> None:
        """
        Builds a worksheet of route totals. See the README file for further description regarding this functionality.

        @param worksheet The excel worksheet to operate on
        @param log The log file to output to
        """
        # Display headers
        worksheet["A1"] = "Route No."
        worksheet["B1"] = "Route"
        worksheet["C1"] = "Ons"
        worksheet["D1"] = "Offs"
        worksheet["E1"] = "Total"
        
        # Display values
        current_row = 2
        for route in self.routes:
            offs, ons, total = self.routes[route].getTotalOffsAndOns()
            worksheet.cell(row=current_row, column=1).value = route
            worksheet.cell(row=current_row, column=2).value = \
                self.routes[route].descriptor
            worksheet.cell(row=current_row, column=3).value = offs
            worksheet.cell(row=current_row, column=4).value = ons
            worksheet.cell(row=current_row, column=5).value = total

            current_row += 1

    def buildMaxLoads(self, worksheet) -> None:
        """
        Builds a worksheet of max loads for each route. See the README file for 
        further description regarding this functionality.

        @param worksheet The excel worksheet to operate on
        @param log The log file to output to
        """
        # Display headers
        worksheet["A1"] = "Route No."
        worksheet["B1"] = "Route"
        worksheet["C1"] = "Start Time"
        worksheet["D1"] = "Ons"
        worksheet["E1"] = "Offs"
        worksheet["F1"] = "Max Load"

        # Display values
        current_row = 2
        for route in self.routes:
            current_row = self.routes[route].buildTotalsByTime(worksheet,\
            current_row)


class Route:
    """
    A class that allows for easy storage of the stops inside of a route
    """
    def __init__(self, route, log:Log):
        """
        Initialize Route with the appropriate internal variables
        """
        self.route = route
        self.stops = {}
        self.times = []
        self.descriptor = "Temp Descriptor"
        self.log = log

    def __str__(self) -> str:
        output = ""
        for key in self.stops:
            output += str(self.stops[key])
        return output

    def __repr__(self) -> str:
        return self.__str__()

    def addStop(self, stop_no, street, cross_street):
        # If stop exists, throw error and return
        if stop_no in self.stops:
            self.log.logError("Tried to add stop " + str(stop_no) + " to route " + str(self.route) + " when it already exists.")
            return

        self.stops[stop_no] = Stop(self.route, stop_no, street, cross_street)

    def addData(self, stop_no, datetime, arrival_time, schedule_time, offs, \
        ons):
        # If stop_no does not exist, throw error and return
        if stop_no not in self.stops:
            self.log.logError("Tried to add data to stop " + str(stop_no) + " in route " + str(self.route) + " when stop does not exist.")
            return

        # If datetime does not exist, add it to datetimes and sort
        if datetime not in self.times:
            self.times.append(datetime)
            self.times.sort()

        self.stops[stop_no].addData(datetime, arrival_time, schedule_time, offs, ons)

    def getDescriptor(self) -> str:
        return self.descriptor

    def getTotalOffsAndOns(self):
        """
        @returns the total offs, ons, and total passengers for this route over 
            all stops
        """
        offs = 0
        ons = 0
        for stop_no in self.stops:
            current_off, current_on = self.stops[stop_no].getTotalOffsAndOns()
            offs += current_off
            ons += current_on
        total = offs + ons
        return offs, ons, total

    def buildTotalsByTime(self, worksheet, current_row):
        for datetime in self.times:
            # Generate the totals per datetime
            offs = 0
            ons = 0
            current_load = 0
            max_load = 0
            for stop_no in self.stops:
                current_off, current_on = \
                    self.stops[stop_no].getOffsAndOns(datetime)

                # Add to tallies
                offs += current_off
                ons += current_on
                current_load = current_load + current_on - current_off

                # Save max load if is the current max
                if current_load > max_load:
                    max_load = current_load

            # Write data to sheet
            worksheet.cell(row=current_row, column=1).value = self.route
            worksheet.cell(row=current_row, column=2).value = self.descriptor
            worksheet.cell(row=current_row, column=3).value = \
                datetime.strftime("%H:%M")
            worksheet.cell(row=current_row, column=4).value = ons
            worksheet.cell(row=current_row, column=5).value = offs
            worksheet.cell(row=current_row, column=6).value = max_load

            current_row += 1
        return current_row


class Stop:
    def __init__(self, route, stop_no, street, cross_street) -> None:
        self.route = route
        self.stop_no = stop_no
        self.street = street
        self.cross_street = cross_street
        self.data = {}

    def __str__(self) -> str:
        output = self.route + ": " + str(self.stop_no) + " [" + self.street + \
            "/" + self.cross_street + "]\n"
        for datetime in self.data:
            output += str(datetime) + " " + str(self.data[datetime][0]) + " " + str(self.data[datetime][1]) + " " + str(self.data[datetime][2]) + " " + str(self.data[datetime][3]) + "\n"
        output += "\n"
        return output

    def __repr__(self) -> str:
        return self.__str__()

    def addData(self, datetime, arrival_time, schedule_time, offs, ons):
        self.data[datetime] = [arrival_time, schedule_time, offs, ons]

    def getOffsAndOns(self, datetime):
        # If datetime does not exist, return zeros
        if datetime not in self.data:
            return 0, 0

        return self.data[datetime][2], self.data[datetime][3]

    def getTotalOffsAndOns(self):
        total_offs = 0
        total_ons = 0
        for datetime in self.data:
            total_offs += self.data[datetime][2]
            total_ons += self.data[datetime][3]
        return total_offs, total_ons
        

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

    # parse the route info file to create route names and times
    route_manager = RouteManager(log)

    # start parsing the ride checks file
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
        #onboard = ride_checks.cell(row=current_row, column=7).value
        stop_number = ride_checks.cell(row=current_row, column=8).value
        arrival_time = ride_checks.cell(row=current_row, column=9).value
        schedule_time = ride_checks.cell(row=current_row, column=10).value
        offs = ride_checks.cell(row=current_row, column=11).value
        ons = ride_checks.cell(row=current_row, column=12).value
        #loads = ride_checks.cell(row=current_row, column=13).value
        #time_check = ride_checks.cell(row=current_row, column=14).value

        # Check that the sequence number is in order, alert if not
        if sequence - 1 != prev_seq:
            log.logWarning("Out-of-order sequence number: Row " +
            str(current_row))
        prev_seq = sequence

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
        if (arrival_time is not None) and (not isinstance(arrival_time, \
            datetime.time)):
            log.logError("Row " + str(current_row) + ": Arrival time '" + \
                str(arrival_time) + "' is not an excel-formatted time."\
                + "Skipping row.")
            current_row += 1
            continue
        if (schedule_time is not None) and (not isinstance(schedule_time, \
            datetime.time)):
            log.logError("Row " + str(current_row) + ": Scheduled time '" + \
                str(schedule_time) + "' is not an excel-formatted time."\
                + "Skipping row.")
            current_row += 1
            continue
        if (ons is not None) and (not isinstance(ons, int)):
            log.logError("Row " + str(current_row) + ": Ons value '" + \
            str(ons) + "' is not an integer. Skipping row.")
            current_row += 1
            continue   
        if (offs is not None) and (not isinstance(offs, int)):
            log.logError("Row " + str(current_row) + ": Offs value '" + \
            str(offs) + "' is not an integer. Skipping row.")
            current_row += 1
            continue           

        # add data to the route manager object
        # order of placement is route string -> date and time -> stop_number
        date_and_time = datetime.datetime.combine(date, start_time)
        route_manager.addData(route, date_and_time, stop_number, \
            arrival_time, schedule_time, offs, ons)

        # increment current row
        current_row += 1

    # Generate route totals sheet
    log.logMessage("Generating route totals")
    routeTotalsSheet = wb.create_sheet("Rte Totals")
    route_manager.buildRouteTotals(routeTotalsSheet)

    # Generate max load sheet
    log.logMessage("Generating max load sheet")
    maxLoadSheet = wb.create_sheet("Max Load")
    route_manager.buildMaxLoads(maxLoadSheet)

    log.logMessage("Generation complete")

    # Try to save the output file
    try:
        wb.save(output_filepath)
    except Exception as e:
        return 2

    # Return successfully
    return 0
    
