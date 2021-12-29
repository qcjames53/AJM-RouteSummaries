# Author: Quinn James (qj@quinnjam.es)
# 
# A program for generating a template bus stop info log workbook
#
# More details about this project can be found in the README file or at:
#   https://github.com/qcjames53/AJM-RouteSummaries

from tkinter import messagebox

from Log import Log

def createTemplateBusStop(log:Log, filepath):
    # TODO - Potentially remove this option?
    messagebox.showinfo("Bus Stop Template Generation", 
        "This feature coming soon...")