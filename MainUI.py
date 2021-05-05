# Date: 2021-05-04
#
# The main UI for the graphical driver application. Interfaces with the
# command-line utility to generate an output spreadsheet.

import tkinter
import webbrowser

# Constants
DefaultWindowSize = "500x200"
RepoUrl = "https://github.com/qcjames53/AJM-RouteSummaries"

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
        selectMenu.add_command(label="Ride Checks")
        selectMenu.add_command(label="Route Information")
        fileMenu.add_cascade(label="Select Input File", menu=selectMenu)
        fileMenu.add_command(label="Exit", command=self.exitProgram)
        menu.add_cascade(label="File", menu=fileMenu)

        # Create Template Menu
        templateMenu = tkinter.Menu(menu)
        generateMenu = tkinter.Menu(templateMenu)
        generateMenu.add_command(label="Ride Checks")
        generateMenu.add_command(label="Route Information")
        templateMenu.add_cascade(label="Generate Template", menu=generateMenu)
        menu.add_cascade(label="Template", menu=templateMenu)

        # Create Run Button
        menu.add_command(label="Run")

        # Create Help Menu
        helpMenu = tkinter.Menu(menu)
        helpMenu.add_command(label="Open Repository", 
            command=self.openRepository)
        menu.add_cascade(label="Help", menu=helpMenu)


    def openRepository(self):
        """
        Opens the repository for the RouteSummaryGenerator project.
        """
        webbrowser.open(RepoUrl, new=1)


    def exitProgram(self):
        """
        Exits the GUI program safely.
        """
        exit()


# Main program start
root = tkinter.Tk()
app = MainWindow(root)
about = AboutWindow(app)
root.wm_title("Route Summaries Utility")
root.geometry(DefaultWindowSize)
root.mainloop()