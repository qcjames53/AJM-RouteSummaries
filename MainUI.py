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
from TemplateGeneratorBusStop import createTemplateBusStop
from RouteSummaryGenerator import generateSummary
from WorkbookUpdateUtility import convertFormat, convertValues
from Log import Log

# Constants
UPDATE_AFTER_LOG = True
DEFAULT_WINDOW_WIDTH = 1200
DEFAULT_WINDOW_HEIGHT = 500
REPO_URL = "https://github.com/qcjames53/AJM-RouteSummaries"
DEFUALT_RIDECHECKS_PREFIX = "ridechecks"
DEFAULT_BUS_STOP_PREFIX = "busstop"
DEFAULT_ROUTE_SUMMARY_PREFIX = "summary"

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
        self.bus_stop_filepath = None

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
        selectMenu.add_command(label="Bus Stop",
            command=self.setBusStop)
        fileMenu.add_cascade(label="Select Input File", menu=selectMenu)
        fileMenu.add_command(label="Exit", command=self.exitProgram)
        menu.add_cascade(label="File", menu=fileMenu)

        # Create Template Menu
        templateMenu = tkinter.Menu(menu)
        generateMenu = tkinter.Menu(templateMenu)
        generateMenu.add_command(label="Ride Checks", 
            command=self.createRideChecks)
        generateMenu.add_command(label="Bus Stop",
            command=self.createBusStop)
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
        self.log_text.yview_scroll(1, "units")
        if UPDATE_AFTER_LOG:
            self.update()

    
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

    
    def createBusStop(self):
        """
        Opens a save-as dialog and creates a bus stop template at the
        provided location.
        """
        # Get default filename
        default_name = DEFAULT_BUS_STOP_PREFIX + getDateString() + ".xlsx"

        # Save-As dialog
        save_filepath = asksaveasfilename(
            title="Save Bus Stop Workbook",
            initialfile=default_name, 
            defaultextension=".xlsx", 
            filetypes=[("Excel Workbook", "*.xlsx")]
        )

        # Check for 'cancel', if true return
        if save_filepath is None or save_filepath == "":
            return

        # Write the file, alert the user
        createTemplateBusStop(self.log, save_filepath)
        self.log.logGeneral("Successfully created the bus stop template"\
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
                ("Excel 97-2003 Workbook", "*.xls"), 
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
            self.log.logGeneral("Updating Excel 97-2003 .xls document...")
            self.update()
            try:
                convertFormat(self.log, filepath, save_filepath)
            except:
                self.log.logFailure("Workbook file format update failed.")
                print(traceback.format_exc())
                return

        # Update the internal values
        self.log.logGeneral("Updating workbook values to current format.")
        self.update()
        try:
            if extension == ".xls":
                convertValues(self.log, save_filepath)
            else:
                convertValues(self.log, filepath, save_filepath=save_filepath)
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


    def setBusStop(self):
        """
        Asks the user to select a bus stop workbook. Sets the bus stop
        member variable to match. Sets to None if cancelled.
        """
        # Ask the user for the ride checks workbook
        filepath = askopenfilename(
            title="Select Bus Stop Workbook",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )

        # If cancel was selected, set None and return
        if filepath is None or filepath == "":
            self.bus_stop_filepath = None
            return
        
        # Set the member variable to the selected filepath, alert the user
        self.bus_stop_filepath = filepath
        self.log.logGeneral("Selected bus stop file '" + \
            self.bus_stop_filepath + "'")


    def runRouteSummary(self):
        """
        Runs the route summary generator with the selected filepaths. Will
        run setRideCheck and/or setBusStop if the respective filepaths have
        not been set.
        """
        # Get the respective sheet filepaths if they haven't been selected
        if self.ride_checks_filepath is None:
            self.setRideChecks()

            # check a valid filepath was submitted. If not, exit.
            if self.ride_checks_filepath is None:
                self.log.logGeneral("Route summary generation was cancelled by the user.")
                return
        if self.bus_stop_filepath is None:
            self.setBusStop()

            # check a valid filepath was submitted. If not, exit.
            if self.bus_stop_filepath is None:
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
                self.ride_checks_filepath, self.bus_stop_filepath, \
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
        self.bus_stop_filepath = None
        self.ride_checks_filepath = None
        

# Main program start
root = tkinter.Tk()
app = MainWindow(root)
root.wm_title("Route Summaries Utility")
root.geometry(str(DEFAULT_WINDOW_WIDTH) + "x" + str(DEFAULT_WINDOW_HEIGHT))
root.mainloop()