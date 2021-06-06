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
from datetime import date
import webbrowser

from TemplateGeneratorRideChecks import createTemplateRideChecks
from TemplateGeneratorRouteInfo import createTemplateRouteSummary
from RouteSummaryGenerator import generateSummary

# Constants
DEFUALT_WINDOW_SIZE = "500x200"
REPO_URL = "https://github.com/qcjames53/AJM-RouteSummaries"
DEFUALT_RIDECHECKS_PREFIX = "ridechecks"
DEFAULT_ROUTE_INFO_PREFIX = "routeinfo"
DEFAULT_ROUTE_SUMMARY_PREFIX = "summary"

def getDateString():
    """
    Helper function for getting the date as a string.
    
    @returns An ISO 8601 formatted date as a string.
    """
    return str(date.today())


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
        menu.add_cascade(label="Template", menu=templateMenu)

        # Create Run Button
        menu.add_command(label="Run", command=self.runRouteSummary)

        # Create Help Menu
        helpMenu = tkinter.Menu(menu)
        helpMenu.add_command(label="Open Repository", 
            command=self.openRepository)
        menu.add_cascade(label="Help", menu=helpMenu)


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
        messagebox.showinfo("Message", message)

    
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
        createTemplateRideChecks(save_filepath)
        self.applicationMessage("Successfully created the ride check template\
             '" + save_filepath + "'")

    
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
        createTemplateRouteSummary(save_filepath)
        self.applicationMessage("Successfully created the route info template\
             '" + save_filepath + "'")


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
        self.applicationMessage("Selected ride checks file '" + \
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
        self.applicationMessage("Selected route info file '" + \
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
                self.applicationMessage("Route summary generation was cancelled by the user.")
                return
        if self.route_info_filepath is None:
            self.setRouteInfo()

            # check a valid filepath was submitted. If not, exit.
            if self.route_info_filepath is None:
                self.applicationMessage("Route summary generation was cancelled by the user.")
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
            self.applicationMessage("Route summary generation was cancelled by the user.")
            return

        # Run the RouteSummaryGenerator utility and log
        generateSummary(self.ride_checks_filepath, self.route_info_filepath, \
            save_filepath)
        self.applicationMessage("Successfully created the route summary\
             '" + save_filepath + "'")
        



# Main program start
root = tkinter.Tk()
app = MainWindow(root)
root.wm_title("Route Summaries Utility")
root.geometry(DEFUALT_WINDOW_SIZE)
root.mainloop()