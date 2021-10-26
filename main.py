#
# This is the main program.  From here we will call the
# export calendar data functions.
#

from ExcelCalendarParser import export_calendar_data
from DataStructs import LodgeTown,location_exceptions, degrees



if __name__ == "__main__":
    print("I'm in main.py main")
    full_list = export_calendar_data()
    print("\n=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-\n")
    print(location_exceptions)
    print(degrees)
    print("\n*****************************************************************\n")
    print(full_list)