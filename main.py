#
# This is the main program.  From here we will call the
# export calendar data functions.
#

from ExcelCalendarParser import export_calendar_data
from ExportDoc import create_word_doc
from DataStructs import locations, location_exceptions, degrees, full_list_calendar_entries, dinner



if __name__ == "__main__":
    print("I'm in main.py main")
    export_calendar_data()
#    print("\n=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-\n")
    print("\n*****************************************************************\n")
#    print(full_list_calendar_entries)
    print(locations)

#     print("\n-------\n")
#     print("Degrees: ")
#     print(degrees)
#     print("\n-------\n")
#     print("Dinner: ")
#     print(dinner)
#     print("\n-------\n")
#     print(location_exceptions)
    create_word_doc()