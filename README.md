# AJM-RouteSummaries

A Python software solution for summarizing ridership statistics from bus routes. The program uses two Microsoft Excel workbooks as input and ouputs five summaries to an Excel workbook, each summary focusing on various statistics.

## Using The Program
This section of documentation is incomplete and will be updated as the program is developed.

Note that this program has the following dependencies:
* [Python 3.9.5](https://www.python.org/downloads/release/python-395/)

---

## Program Inputs
Route Summaries requires two Microsoft Excel workbook files as input to properly produce summaries: A ride checks data file and a route information file. Template files can be created through the program menu (future feature, not yet implemented).

#### Ride Checks Data File
An MS-Excel workbook containing any number of rows with row 0 reserved for fixed headers. Each row contains the following columns of data, representing one ride check data entry:
* SEQUENCE
    * Sequence number, the index of this ride check data entry. Numbers should be in order, with the first row (excel row 2) starting at sequence number 1. This number is for sorting convenience in the inital workbook.
* DATE
    * An MS-Excel formatted date representing the date of data collection.
* ROUTE
    * The route number of this ride check data entry.
* DIRECTION
    * The direction of the route. Valid inputs are **[IB, OB, NB, EB, SB, WB]**, representing the inbound busses, outbound busses, and the four cardial directions.
* RUN
    * The run number for this route, a unique shift identifier for drivers.
* START TIME
    * An MS-Excel formatted time representing the start time of the current route. This time should be consistent for all entries in a  given date, route, direction, run combination.
* ONBOARD
    * The number of passengers carried over from a previous route on a given run if the number of passengers on board is above 0. Used as a human check in the input spreadsheet. 
* STOP NUMBER
    * The bus stop number of this stop. Stops where no passengers depart or board may be ommited.
* ARRIVAL TIME
    * The actual arrival time for this ride check data entry. May be left blank if SCHEDULE TIME is also left blank.
* SCHEDULE TIME
    * The scheduled arrival time for this ride check data entry.
* OFFS
    * The number of passengers departing the bus at this stop.
* ONS
    * The number of passengers boarding the bus at this stop.
* LOADS
    * The number of passengers on the bus after this stop.
* TIME CHECK
    * The difference between the ARRIVAL TIME and the SCHEDULE TIME for this stop. May be left blank if SCHEDULE TIME is also left blank.

#### Route Information File
An MS-Excel workbook containing any number of sheets, each of which represents a ROUTE-DIRECTION combination of the form 'Rte #D' where # is the ROUTE and D is the DIRECTION as defined above in the [Ride Checks Data File section](https://github.com/qcjames53/AJM-RouteSummaries#ride-checks-data-file).

A template file can be created using the program menu. Each sheet may contain an unlimited number of column-aligned tables on arbitrary rows. The following information is required for each table header:

* ROUTE
    * The route number of this table

* DIRECTION
    * The direction of the route. Valid inputs are **[IB, OB, NB, EB, SB, WB]**, representing the inbound busses, outbound busses, and the four cardial directions. 
* START TIME
    * An MS-Excel formatted time representing the start time of the current route.

Additionally, each row of the table represents a stop on this route. Each row must be populated with the following columns:

* STREET
    * The name of the main street of this stop. This name is flexible depending on the naming convention of the bus route.
* CROSS STREET
    * The name of the cross street of this stop. This name is flexible depending on the naming convention of the bus route.
* NO
    * The stop number for this route. Columns must be in numberical order from top to bottom.

---

## Program Output
Route Summaries outputs an MS-Excel workbook containing seven sheets displaying various summaries. Descriptions of the outputs are detialed below:

#### Log
A sheet that details errors and warnings encountered during generation of this output workbook.

#### Rte Totals
A sheet which shows the passenger totals for each route / direction combination. Formatted as rows of totals, each column representing the following:
* Route No.
    * This line's route number.
* Route
    * A descriptor of this route. Example: University OB
* Ons
    * The total number of boarded passengers.
* Offs
    * The total number of departed passengers.
* Total
    * The total of the boarded and departed passenger counts.

#### Max Load
A sheet which shows the maximum number of passengers on a given route over time. Data is formatted as rows, each column representing the following:
* Route No.
    * This line's route number.
* Route
    * A descriptor of this route. Example: University OB
* Start Time
    * An MS-Excel formatted time representing the start time of the route.
* Ons
    * The total number of boarded passengers.
* Offs
    * The total number of departed passengers.
* Max Load
    * The maximum number of passengers on a bus during this route.

#### Detailed Trip Summaries
A sheet which shows boardings, departures, and carryover passengers for each stop for each route over the course of a day.

#### On Time Detail
A sheet which shows bus routes that arrive on-time or behind schedule. Data is diplayed for each timed stop for each route for each scheduled route start time.

#### Ons, Offs, Tot & Ld
A sheet which displays totals for each stop on each route. Data is output in rows, each column representing the following:
* Route No.
    * This line's route number.
* Route
    * A descriptor of this route. Example: University OB
* Stop No.
    * The stop number of this stop
* Street
    * The name of the main street of this stop.
* Cross Street
    * The name of the cross street of this stop.
* Ons
    * The total number of boarding passengers over all trips at this stop.
* Offs
    * The total number of departing passengers over all trips at this stop.
* Total
    * The total number of boarding and departing passengers over all trips at this stop.
* Load
    * The total number of passengers onboard the bus after this stop.

Each route is prefaced by a row titled 'ONBOARD'. This row represents the total passengers being carried over from a previous route. Additionally, each route is appended by a row titled 'TOTAL'. This row represents the total of all four data columns for all stops on this route.

#### Notes
An additional sheet used to detail operational details which may affect the collected data. This data is manually added after workbook generation.