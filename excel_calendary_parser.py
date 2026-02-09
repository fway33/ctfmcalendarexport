# This python file will parse the calendar dump Excel spreadsheet
# and produce a text output suitable for editing.
import datetime
import re
import numpy

from openpyxl.reader.excel import load_workbook
from openpyxl import workbook
from openpyxl.cell.cell import Cell

from DataStructs import LodgeTown, BadLocations, lodge_locations, location_exceptions, full_list_calendar_entries, has_dinner, \
    degrees





def export_calendar_data() -> None:
    """ This function will open the Excel workbook, choose the 'data' sheet,
    and ignore the 'calc' sheet, then export data and modify it as needed."""
    workbook = load_workbook('CalDump1.xlsx')
    worksheet = workbook['data']

    # Find out how many rows are populated -- this is different than max_rows.
    populated_rows = len([row for row in worksheet if any(cell.value is not None for cell in row)])

    #DBG
    print(f"Populated rows: {populated_rows}")

    # initialize a list of lists, initialized to -1's, to hold our data when transformed
    # also initialize a list to hold locations when there is some doubt
    for i in range(populated_rows - 1):
        full_list_calendar_entries.insert(i, [-1, -1, -1, -1, -1])
        lodge_locations.insert(i, '')

    if not full_list_calendar_entries[populated_rows - 1]:
        full_list_calendar_entries.pop(populated_rows - 1)

    # Set the ending cell identifier...column "E" with the last populated row number
    ending_cell = 'E' + str(populated_rows)
    #DBG
    print(f"Ending cell: {ending_cell}")
    table = numpy.array([[cell.value for cell in col] for col in worksheet['A2':ending_cell]])

    # The first value in the row is the lodge name.  Need to change that from 'Lodge 001 Hiram'
    # to "Hiram Lodge No. 1".  Run through each column (lodge, event title, event description, location, time)
    # and extract the data, and modify it, then save it.
    modify_lodge(table)
    modify_event_title(table)
    modify_event_descr(table)
    modify_event_location(table)
    modify_date(table)

    #return


def modify_lodge(table: numpy.ndarray) -> None:
    """ This function creates a list of the first value (lodge name) in every list (row) in the
    table. The table is a list of lists of the spreadsheet values, and this list is  are
    all lodge names.  It then modifies those lodge names according to a regex"""

    lodge_array = [item[0] for item in table]

    # Since we will use it many times, compile the regex for the lodge name as
    # written in the cell (i.e., "Lodge 001 Hiram") and capture the values.
    rx = re.compile(r'(Lodge) (\d{3}) ([a-zA-Z\-\'\. ]+)')

    # Loop through the list elements and if the regex matches, the modify the string
    # (i.e., from 'Lodge 001 Hiram' to 'Hiram Lodge No. 1')
    array_idx = 0

    for lodge in lodge_array:
        result = rx.match(lodge)
        if result:
            lodge_string = f"{result.group(3)} {result.group(1)} No. {result.group(2).lstrip('0')}"
            full_list_calendar_entries[array_idx][0] = lodge_string
        array_idx += 1


def modify_event_title(table: numpy.ndarray) -> None:
    # Create an array of the second value in every list in table, the list of lists
    # of the spreadsheet values.  These are all event descriptions.
    event_array = [item[1] for item in table]

    rx = re.compile(r'.*[Dd]egree.*')

    # basically we're looking to see if there is a degree mentioned in the title,
    # otherwise we're not going to modify it.  If there's an instance of degree,
    # put the array index into the degree list -- we'll use that later to prepend
    # the line in the doc with "DEGREE" so that it can easily be identified.
    array_idx = 0
    for event in event_array:
        result = rx.match(event)
        if result is not None:
            degrees.append(array_idx)
        full_list_calendar_entries[array_idx][1] = event
        array_idx += 1

def modify_event_descr(table: numpy.ndarray) -> None:
    event_descr_array = [item[2] for item in table]

    rx = re.compile(r'.*[Dd]egree.*')
    ry = re.compile(r'.*[Dd]inner.*')

    # Similar to the event title, we're looking to see if they mention degree in here.
    # Also, this is usually where they mention 'dinner', so we will add the array index
    # to the dinner or degrees list (checking to make sure we don't already have the index
    # in the degree list from modifying the event title).
    array_idx = 0
    for event_descr in event_descr_array:
        if event_descr:
            result = rx.match(event_descr)
            if result:
                if array_idx not in degrees:
                    degrees.append(array_idx)
            result2 = ry.search(event_descr)
            if result2:
                has_dinner.append(array_idx)
        full_list_calendar_entries[array_idx][2] = event_descr
        array_idx += 1


def modify_event_location(table:numpy.ndarray) -> None:
    event_loc_array = [item[3] for item in table]

    rw = re.compile(r'.*[,]* (.*), (CT USA).*')
    rx = re.compile(r'[\w ]*[,]? (.*), (CT \d{5})[,]*.*')
    ry = re.compile(r'(Lodge) (\d{3}) ([a-zA-Z\-\'\. ]+)')
    rz = re.compile(r'.*, (.*), (CT)[,]* .*')
    ra = re.compile(r'.*, (.*)')

    array_idx = 0

    # Look through all the locations and try to extract a town name.  This gets very tricky.  While there are
    # some locations that are easily parsed, some are more difficult.  We'll maintain two
    # lists -- one of locations (what we parse) and one of locations exceptions (an array index of what we
    # could not parse at all.)   In all instances where there is a location field present, we'll put the
    # town if we can find it into the doc, and the location field too -- the decision of which to use can
    # be made when editing the doc.
    for event_loc in event_loc_array:
        if not event_loc:
            # There was no location specified in the spreadsheet.  However, there is always
            # a lodge name.  Use the lodge name to grab the lodge location from the LodgeTown dictionary
            full_list_calendar_entries[array_idx][3] = LodgeTown[table[array_idx][0]]
        else:
            # There is a location in the cell. Start by trying to pull the town out.
            # BUT, this might be a legit address for a non-lodge event (say pizza night at a pizza place).
            # Leave the location cell alone, but put the possible town into a separate location array
            #
            # Start by matching against the CT {zip} regex
            result = rx.match(event_loc)
            if result:
                # Leaving in the deubgging print statements in this function.  Might need them in future months.
                # print("Match rx")
                # print(result.group(1))

                # This match might have more than just the town.  Run it against another
                # regex to try to get just the town.  If it fails, use what we just got.
                result2 = ra.match(result.group(1))
                if result2:
#                    print(result2.group(1))
                    lodge_locations[array_idx] = result2.group(1)
                else:
                    lodge_locations[array_idx] = result.group(1)
                full_list_calendar_entries[array_idx][3] = event_loc
            else:
                # Since CT {zip} didn't work, see if 'CT' is by itself.
                result2 = rz.match(event_loc)
                if result2 is not None:
#                    print("Match rz")
#                    print(result2.group(1))
                    lodge_locations[array_idx] = result2.group(1)
                    full_list_calendar_entries[array_idx][3] = event_loc
                else:
                    # If not CT {zip} or just CT, it might be CT USA, which sometimes shows up
                    result2 = rw.match(event_loc)
                    if result2:
#                        print("Match rw")
#                        print(result2.group(1))
                        lodge_locations[array_idx] = result2.group(1)
                        full_list_calendar_entries[array_idx][3] = event_loc
                    else:
                        # It did not match on any 'CT' so see if the location is maybe a Lodge name
                        # (i.e., Lodge 001 Hiram).   If so, use the LodgeTown dictionary to get the
                        # address.
                        result2 = ry.fullmatch(event_loc)
                        # Try the fullmatch first, that's easiest
                        if result2:
                            # print("match ry -- fullmatch)")
                            # print(result2)
                            # print(LodgeTown[loc])
                            full_list_calendar_entries[array_idx][3] = LodgeTown[event_loc]
                        else:
                            # Otherwise use search
                            result2 = ry.search(event_loc)
                            if result2:
                                # print("Match ry -- search")
                                # print(result2)
                                lodge = result2.group(1) + " " + result2.group(2) + " " + result2.group(3)

                                # Check that what we got is in the dictionary keys.  There have been
                                # a couple of instances where a string like "Lodge 263 Center Street"
                                # matches here and we don't want to blow up
                                if lodge in LodgeTown.keys():
                                    # print(lodge)
                                    # print(LodgeTown[lodge])
                                    # print(array_idx)
                                    # print(full_list_calendar_entries[array_idx])
                                    full_list_calendar_entries[array_idx][3] = LodgeTown[lodge]
                                else:
                                    # The "lodge" that we found in the location cell didn't have a key
                                    # in the dictionary.  There is a small dictionary of known bad
                                    # locations, see if it is there, and if so put in the correct town
                                    # from the dictionary.
                                    if event_loc.strip() in BadLocations:
                                        # print("*!*!*!*!*!*!* BAD LOCATIONS *!*!*!*!*!*!*!*!")
                                        # print(loc)
                                        full_list_calendar_entries[array_idx][3] = BadLocations[event_loc.strip()]
                                    else:
                                        # We've run out of options here, so write the location into the
                                        # entries, and put this loc into locations_exceptions for further review.
                                        location_exceptions.append(event_loc)
                                        full_list_calendar_entries[array_idx][3] = event_loc
                            else:
                                # search did not rturn up anything.  Try the location against
                                # entries in the BadLocations dictionary.  If not, do the same as above
                                # and put the loc inot the entries, and location_exceptions.
                                # print("TRYING BAD LOC")
                                # print(loc)
                                if event_loc.strip() in BadLocations:
                                    # print("*!*!*!*!*!*!* BAD LOCATIONS *!*!*!*!*!*!*!*!")
                                    # print(loc)
                                    full_list_calendar_entries[array_idx][3] = BadLocations[loc.strip()]
                                else:
                                    location_exceptions.append(event_loc)
                                    full_list_calendar_entries[array_idx][3] = event_loc
        array_idx += 1


def modify_date(table: numpy.ndarray) -> None:
    time_array = [item[4] for item in table]

    # Find the current year.  The style guide says we don't use the year
    # if it is in the current year, so grab this to compare later.
    curr_yr = datetime.datetime.today().year

    array_idx = 0
    for time in time_array:
        time_str = time.strftime("%a %b. %-d")

        # If the year is greater than the current year, add it to the time_str so it prints.
        if time.year > curr_yr:
            time_str += ", "
            time_str += str(time.year)
        time_str += ", "

        # Style guide is to not print the :00 if it is 00 so do that.
        if time.minute == 00:
            time_str += time.strftime("%-I ")
        else:
            time_str += time.strftime("%-I:%M ")

        # We use a.m. and p.m. so make te appropriate choice.
        if time.hour > 12:
            time_str += "p.m."
        else:
            time_str += "a.m."

        full_list_calendar_entries[array_idx][4] = time_str
        array_idx += 1

