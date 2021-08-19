# Author: Quinn James (qj@quinnjam.es)
#
# The main UI for the graphical driver application. Interfaces with the
# command-line utility to generate an output spreadsheet.
#
# More details about this project can be found in the README file or at:
#   https://github.com/qcjames53/AJM-RouteSummaries

from tkinter import messagebox
import tkinter
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askopenfilename
from datetime import date, datetime
import webbrowser
import traceback
import os

from TemplateGeneratorRideChecks import createTemplateRideChecks
from TemplateGeneratorRouteInfo import createTemplateRouteSummary
from RouteSummaryGenerator import generateSummary
from WorkbookUpdateUtility import convertFormat, convertValues
from enum import Enum
from inspect import getframeinfo, stack, Traceback
from pathlib import Path

# Constants
DEFAULT_WINDOW_WIDTH = 1200
DEFAULT_WINDOW_HEIGHT = 500
REPO_URL = "https://github.com/qcjames53/AJM-RouteSummaries"
DEFUALT_RIDECHECKS_PREFIX = "ridechecks"
DEFAULT_ROUTE_INFO_PREFIX = "routeinfo"
DEFAULT_ROUTE_SUMMARY_PREFIX = "summary"
LOG_SHEET_TITLE = "Log"
LOG_PRINT_TIMESTAMP = False
LOG_PRINT_SEVERITY = True
LOG_PRINT_MESSAGE = True
LOG_PRINT_LOCATION = False

def getDateString():
    """
    Helper function for getting the date as a string.
    
    @returns An ISO 8601 formatted date as a string.
    """
    return str(date.today())


def getDateTimeString():
    """
    Helper function for getting the datetime as a string.

    @returns An ISO 8601 formatted datetime as a string.
    """
    return str(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

# Log severity enum
class Severity(Enum):
    GENERAL = "G"
    WARNING = "W"
    ERROR = "E"
    FAILURE = "F"

class Log:
    def __init__(self, log_method) -> None:
        self.messages = []
        self.creation_time = datetime.now()
        self.log_method = log_method

    def __str__(self) -> str:
        output = ""
        for message in self.messages:
            output += str(message) + '\n'
        return output

    def __repr__(self) -> str:
        return self.__str__()

    def logMessage(self, severity: Severity, message: str, \
        # Create the message
        location: Traceback= None) -> None:
        temp_message = LogMessage(self.log_method, severity, message, location)
        self.messages.append(temp_message)

        # Output the message
        temp_message.output()

    def logGeneral(self, message: str):
        location = getframeinfo(stack()[1][0])
        self.logMessage(Severity.GENERAL, message, location)

    def logWarning(self, message: str):
        location = getframeinfo(stack()[1][0])
        self.logMessage(Severity.WARNING, message, location)

    def logError(self, message: str):
        location = getframeinfo(stack()[1][0])
        self.logMessage(Severity.ERROR, message, location)

    def logFailure(self, message: str):
        location = getframeinfo(stack()[1][0])
        self.logMessage(Severity.FAILURE, message, location)

class LogMessage:
    def __init__(self, log_method, severity: Severity, message: str, \
        location: Traceback) -> None:
        self.log_method = log_method
        self.creation_time = datetime.now()
        self.severity = severity
        self.message = message
        self.location = location

    def __str__(self) -> str:
        output = ""
        if LOG_PRINT_TIMESTAMP:
            output = str(self.creation_time) + " "
        if LOG_PRINT_SEVERITY:
            if self.severity == Severity.GENERAL:
                output += "[General] "
            elif self.severity == Severity.WARNING:
                output += "[Warning] "
            elif self.severity == Severity.ERROR:
                output += "[Error]   "
            else:
                output += "[Failure] "
        if LOG_PRINT_MESSAGE:
            output += self.message + " "
        if LOG_PRINT_LOCATION:
            output += "[" + self.getLocationShortFormatted() + "]"
        return output

    def __repr__(self) -> str:
        return self.__str__()

    def getLocationShortFormatted(self) -> str:
        file_name = Path(self.location.filename).stem 
        return file_name + ":" + str(self.location.lineno)
        
    def output(self) -> None:
        self.log_method(self.__str__())


class MainWindow(tkinter.Frame):
    """
    A class that defines the main window of the GUI application for the route
    summary generator. Extends tkinter.Frame.
    """

    def __init__(self, master=None):
        """
        Initialize the main window of the GUI application.

        @param master The parent window of the main window. Defaults to None
        """
        # Non-tkinter member variables
        self.ride_checks_filepath = None
        self.route_info_filepath = None

        # Handle instantiation
        tkinter.Frame.__init__(self, master)
        self.master = master
        self.master.option_add('*tearOff', False)

        # Create the options menu bar
        menu = tkinter.Menu(self.master)
        self.master.config(menu=menu)

        # Create the file menu
        fileMenu = tkinter.Menu(menu)
        selectMenu = tkinter.Menu(fileMenu)
        selectMenu.add_command(label="Ride Checks", 
            command=self.setRideChecks)
        selectMenu.add_command(label="Route Information",
            command=self.setRouteInfo)
        fileMenu.add_cascade(label="Select Input File", menu=selectMenu)
        fileMenu.add_command(label="Exit", command=self.exitProgram)
        menu.add_cascade(label="File", menu=fileMenu)

        # Create Template Menu
        templateMenu = tkinter.Menu(menu)
        generateMenu = tkinter.Menu(templateMenu)
        generateMenu.add_command(label="Ride Checks", 
            command=self.createRideChecks)
        generateMenu.add_command(label="Route Information",
            command=self.createRouteInfo)
        templateMenu.add_cascade(label="Generate Template", menu=generateMenu)
        templateMenu.add_command(label="Old Format Conversion", 
            command=self.convertOldFormat)
        menu.add_cascade(label="Utility", menu=templateMenu)

        # Create Run Button
        menu.add_command(label="Run", command=self.runRouteSummary)

        # Create Help Menu
        helpMenu = tkinter.Menu(menu)
        helpMenu.add_command(label="Open Repository", 
            command=self.openRepository)
        menu.add_cascade(label="Help", menu=helpMenu)

        # Create a scrolling debug/message log on the GUI application
        self.log_text = tkinter.Text(self.master)
        scroll_bar = tkinter.Scrollbar(self.master, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scroll_bar.set)
        scroll_bar.pack(side=tkinter.RIGHT, fill=tkinter.Y)
        self.log_text.pack(side=tkinter.LEFT, expand=True, fill=tkinter.BOTH)

        # Create the log
        self.log = Log(self.applicationMessage)
        self.log.logGeneral("Application launched")


    def openRepository(self):
        """
        Opens the repository for the RouteSummaryGenerator project.
        """
        webbrowser.open(REPO_URL, new=1)


    def exitProgram(self):
        """
        Exits the GUI program safely.
        """
        exit()


    def applicationMessage(self, message):
        """
        Alerts the user with a provided message

        @param message A string to display to the user
        """
        self.log_text.insert(tkinter.END, "\n" + message)

    
    def createRideChecks(self):
        """
        Opens a save-as dialog and creates a ride checks template at the
        provided location.
        """
        # Get default filename
        default_name = DEFUALT_RIDECHECKS_PREFIX + getDateString() + ".xlsx"
        
        # Save-As dialog
        save_filepath = asksaveasfilename(
            title="Save Ride Checks Workbook",
            initialfile=default_name, 
            defaultextension=".xlsx", 
            filetypes=[("Excel Workbook", "*.xlsx")
        ])

        # Check for 'cancel', if true return
        if save_filepath is None or save_filepath == "":
            return

        # Write the file, alert the user
        createTemplateRideChecks(self.log, save_filepath)
        self.log.logGeneral("Successfully created the ride check template"\
        + " '" + save_filepath + "'")

    
    def createRouteInfo(self):
        """
        Opens a save-as dialog and creates a route info template at the
        provided location.
        """
        # Get default filename
        default_name = DEFAULT_ROUTE_INFO_PREFIX + getDateString() + ".xlsx"

        # Save-As dialog
        save_filepath = asksaveasfilename(
            title="Save Route Info Workbook",
            initialfile=default_name, 
            defaultextension=".xlsx", 
            filetypes=[("Excel Workbook", "*.xlsx")]
        )

        # Check for 'cancel', if true return
        if save_filepath is None or save_filepath == "":
            return

        # Write the file, alert the user
        createTemplateRouteSummary(self.log, save_filepath)
        self.log.logGeneral("Successfully created the route info template"\
            + " '" + save_filepath + "'")

    
    def convertOldFormat(self):
        """
        Asks the user to select a document to convert. Tries to update this
        document to the new document format (.xls to .xlsx, replaces numberical
        dates and times with Excel dates and times, etc.)
        """
        # Ask the user to select the old workbook
        self.log.logGeneral("Selecting old-format workbook file...")
        filepath = askopenfilename(
            title="Select Old-format Workbook", 
            filetypes=[
                ("Excel 1995-2003 Workbook", "*.xls"), 
                ("Excel Workbook", "*.xlsx")
                ]
        )

        # If cancel was selected, set None and return
        if filepath is None or filepath == "":
            self.log.logGeneral("File selection was cancelled by the user.")
            return
        self.log.logGeneral("Selected old-format file '" + \
            filepath + "'")

        # Ask where to save the generated sheet
        # Get default filename & extension
        basename = os.path.splitext(os.path.basename(filepath))
        filename = basename[0]
        extension = basename[1]
        default_name = filename + ".xlsx"

        # Save-As dialog
        save_filepath = asksaveasfilename(
            title="Save Ride Checks Workbook",
            initialfile=default_name, 
            defaultextension=".xlsx", 
            filetypes=[("Excel Workbook", "*.xlsx")]
        )

        # Check for 'cancel', if true return
        if save_filepath is None or save_filepath == "":
            self.log.logGeneral("File selection was cancelled by the user.")
            return

        # Update to .xlsx if neccesary
        if extension == ".xls":
            self.log.logGeneral("Updating Excel 1995-2003 .xls document...")
            self.update()
            try:
                convertFormat(self.log, filepath, save_filepath)
            except:
                self.log.logFailure("Workbook file format update failed.")
                print(traceback.format_exc())
                return

        # Update the internal values
        self.log.logGeneral("Updating workbook values to current format.")
        try:
            if extension == ".xls":
                convertValues(self.log, save_filepath)
            else:
                convertValues(self.log, filepath, save_filepath=save_filepath)
            self.log.logGeneral("Successfully updated the workbook.")
        except:
            self.log.logFailure("Workbook value update failed.")
            print(traceback.format_exc())
            return


    def setRideChecks(self):
        """
        Asks the user to select a ride checks workbook. Sets the ride checks
        member variable to match. Sets to None if cancelled.
        """
        # Ask the user for the ride checks workbook
        filepath = askopenfilename(
            title="Select Ride Checks Workbook",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )

        # If cancel was selected, set None and return
        if filepath is None or filepath == "":
            self.ride_checks_filepath = None
            return
        
        # Set the member variable to the selected filepath, alert the user
        self.ride_checks_filepath = filepath
        self.log.logGeneral("Selected ride checks file '" + \
            self.ride_checks_filepath + "'")


    def setRouteInfo(self):
        """
        Asks the user to select a route info workbook. Sets the route info
        member variable to match. Sets to None if cancelled.
        """
        # Ask the user for the ride checks workbook
        filepath = askopenfilename(
            title="Select Route Info Workbook",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )

        # If cancel was selected, set None and return
        if filepath is None or filepath == "":
            self.route_info_filepath = None
            return
        
        # Set the member variable to the selected filepath, alert the user
        self.route_info_filepath = filepath
        self.log.logGeneral("Selected route info file '" + \
            self.route_info_filepath + "'")


    def runRouteSummary(self):
        """
        Runs the route summary generator with the selected filepaths. Will
        run setRideCheck and/or setRouteInfo if the respective filepaths have
        not been set.
        """
        # Get the respective sheet filepaths if they haven't been selected
        if self.ride_checks_filepath is None:
            self.setRideChecks()

            # check a valid filepath was submitted. If not, exit.
            if self.ride_checks_filepath is None:
                self.log.logGeneral("Route summary generation was cancelled by the user.")
                return
        if self.route_info_filepath is None:
            self.setRouteInfo()

            # check a valid filepath was submitted. If not, exit.
            if self.route_info_filepath is None:
                self.log.logGeneral("Route summary generation was cancelled by the user.")
                return

        # Ask where to save the generated sheet
        # Get default filename
        default_name = DEFAULT_ROUTE_SUMMARY_PREFIX + getDateString() + ".xlsx"

        # Save-As dialog
        save_filepath = asksaveasfilename(
            title="Save Route Summary Workbook",
            initialfile=default_name, 
            defaultextension=".xlsx", 
            filetypes=[("Excel Workbook", "*.xlsx")]
        )

        # Check for 'cancel', if true return
        if save_filepath is None or save_filepath == "":
            self.log.logGeneral("Route summary generation was cancelled by the user.")
            return

        # Run the RouteSummaryGenerator utility and log
        self.log.logGeneral("Generating route summary...")
        self.update()
        try:
            summary_generation_result = generateSummary(self.log, \
                self.ride_checks_filepath, self.route_info_filepath, \
                save_filepath)
            # Console message for the correct result
            if(summary_generation_result == 0):
                self.log.logGeneral("Successfully created the route summary '"\
                + save_filepath + "'")
            elif(summary_generation_result == 1):
                self.log.logError("Major error enconutered while generating summary.")
            elif(summary_generation_result == 2):
                self.log.logFailure("Output workbook '" + save_filepath +\
                    "' could not be created.")
            else:
                self.log.logError("Unspecified error")

        except:
            self.log.logFailure("Summary generation failed.")
            print(traceback.format_exc())

        # Reset the selected filepaths
        self.route_info_filepath = None
        self.ride_checks_filepath = None




# Main program start
root = tkinter.Tk()
app = MainWindow(root)
root.wm_title("Route Summaries Utility")
root.geometry(str(DEFAULT_WINDOW_WIDTH) + "x" + str(DEFAULT_WINDOW_HEIGHT))
root.mainloop()