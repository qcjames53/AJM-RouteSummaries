# Author: Quinn James (qj@quinnjam.es)
#
# A log object for storing and displaying status of the program.
#
# More details about this project can be found in the README file or at:
#   https://github.com/qcjames53/AJM-RouteSummaries

from enum import Enum
from inspect import getframeinfo, stack, Traceback
from pathlib import Path
from datetime import datetime

# Constants
LOG_SHEET_TITLE = "Log"
LOG_PRINT_TIMESTAMP = False
LOG_PRINT_SEVERITY = True
LOG_PRINT_MESSAGE = True
LOG_PRINT_LOCATION = False

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