#
# This is the main program.  From here we will call the
# export calendar data functions.
#

from ExcelCalendarParser import export_calendar_data
from ExportDoc import create_word_doc


def main() -> None:
    export_calendar_data()
    #create_word_doc()

if __name__ == "__main__":
    main()
