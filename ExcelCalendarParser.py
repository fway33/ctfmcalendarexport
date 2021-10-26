# This python file will parse the calendar dump excel spreadsheet
# and produce a text output suitable for editing.
import datetime

import openpyxl
import numpy
import re
import regex
from datetime import date
import calendar
from DataStructs import LodgeTown, degrees, location_exceptions



def get_populated_row_count(ws):
    number_of_rows = ws.max_row
    last_row_index_with_data = 0

    while True:
        if ws.cell(number_of_rows, 3).value is not None:
            last_row_index_with_data = number_of_rows
            break
        else:
            number_of_rows -= 1
    return number_of_rows


def export_calendar_data():
    print("Here is where we'd start to parse")

    # Open the workbook.  There are two worksheets, 'data' and 'calc', we want 'data'
    wb = openpyxl.load_workbook('CalDump1.xlsx')
    ws = wb['data']

    # Find out how many rows are populated -- this is different than max rows.
    rows = get_populated_row_count(ws)

    # create a list of lists, initialized to -1's, to hold our data when transformed
    full_list_calender_entries = [[-1, -1, -1, -1, -1] for i in range(rows)]

#    degrees = []
    locations = ['' for i in range(rows)]
    # Set the ending cell identifier...column "E" with the last populated row number
    end = 'E' + str(rows)
    table = numpy.array([[cell.value for cell in col] for col in ws['A2':end]])

    # The first value in the row is the lodge name.  Need to change that from 'Lodge 001 Hiram'
    # to "Hiram Lodge No. 1".
    modify_lodge(table, full_list_calender_entries)
#    modify_event_title(table, full_list_calender_entries, degrees)
    #   modify_event_descr(table, full_list_calendar_entries, degrees)
#    modify_event_location(table, full_list_calender_entries, locations)
    modify_date(table,full_list_calender_entries)

    print("\n=-=-=-=-=-=-\n")
    print(full_list_calender_entries)
    print(degrees)
    return full_list_calender_entries


def modify_lodge(table, full_list_calender_entries):
    # Create an array of the first value in every list in table, the lists of lists
    # of the spreadsheet values.  These are all lodge names.
    lodge_array = [item[0] for item in table]
    #    print(lodge_array)

    #    print(lodge_array[0])

    # Since we will use it many times, compile the regex for the lodge name as
    # written in the cell (i.e., "Lodge 001 Hiram") and capture the values.
    rx = re.compile(r'(Lodge) (\d{3}) ([a-zA-Z\-\'\. ]+)')

    # Loop through the array elements and if the regex matches, the modify the string
    # (i.e., from 'Lodge 001 Hiram' to 'Hiram Lodge No. 1')
    array_idx = 0
    print(full_list_calender_entries[array_idx])
    for lodge in lodge_array:
        result = rx.match(lodge)
        if result:
            #            print(result.group(3) + " " + result.group(1) + " No. " + result.group(2).lstrip("0"))
            lodge_string = result.group(3) + " " + result.group(1) + " No. " + result.group(2).lstrip("0")
            full_list_calender_entries[array_idx][0] = lodge_string
        array_idx += 1


def modify_event_title(table, full_list_calendar_entries, degrees):
    # Create an array of the second value in every list in table, the lists of lists
    # of the spreadsheet values.  These are all event descriptions.
    event_array = [item[1] for item in table]
    print(event_array)

    rx = re.compile(r'.*[Dd]egree.*')

    array_idx = 0
    for event in event_array:
        result = rx.match(event)
        if result is not None:
            #            print(event, end="|")
            degrees.append(array_idx)
        full_list_calendar_entries[array_idx][1] = event
        array_idx += 1
    print()
    print(degrees)


def modify_event_location(table, full_list_calendar_entries, locations):
    loc_array = [item[3] for item in table]
    print(loc_array)
    print("........................")
    array_idx = 0
    rx = re.compile(r'.*, (.*), (CT \d{5}).*')
    ry = re.compile(r'(Lodge) (\d{3}) ([a-zA-Z\-\'\. ]+)')
    rz = re.compile(r'.*, (.*), (CT).*')


    # loop through the locations.  See what they match, if anything
    for loc in loc_array:
        if loc is None:
            # there was no location specified in the spreadsheet, so leave it empty in the result
            full_list_calendar_entries[array_idx][3] = ""
        else:
            print(loc)
            # start by trying to pull the town out.  BUT, this might be a legit address for a
            # non-lodge event (say pizza night at a pizza place).  Leave the location
            # cell alone, but put the possible town into a separate location array
            result = rx.match(loc)
            if result is not None:
                print(result.group(1))
                locations[array_idx] = result.group(2)
                full_list_calendar_entries[array_idx][3] = loc
            else:
                # See if 'CT' is by itself.
                result2 = rz.match(loc)
                if result2 is not None:
                    print(result2.group(1))
                    locations[array_idx] = result2.group(1)
                    full_list_calendar_entries[array_idx][3] = loc
                else:

                    # it didn't have 'CT' in the location, so see if it is maybe a Lodge 001 Hiram
                    # situation.  If so, use the lodge/town dictionary to get the address.
                    # (could probably just use the dictionary when we grab the lodge but this
                    # is a little more flexible)
                    result2 = ry.fullmatch(loc)
                    if result2 is not None:
                        print(result2)
                        print(LodgeTown[loc])
                        locations[array_idx] = LodgeTown[loc]
                        full_list_calendar_entries[array_idx][3] = loc
                    else:
                        result2 = ry.search(loc)
                        if result2:
                            print(result2)
                            lodge = result2.group(1) + " " + result2.group(2) + " " + result2.group(3)
                            if lodge in LodgeTown.keys():
                                print(LodgeTown[lodge])
#                                locations[array_idx] = LodgeTown[lodge]
                                full_list_calendar_entries[array_idx][3] = LodgeTown[lodge]
                            else:
                                print("********FUCKED UP********" + loc)
                                location_exceptions.append(loc)
                               # For now, just jam it in there...maybe we will develop other regexs later
                                full_list_calendar_entries[array_idx][3] = loc
        array_idx += 1
        print("@@@@@@@@@@@@@@@@@@@@")

def modify_date(table,full_list_calendar_entries):
    time_array = [item[4] for item in table]
    print(time_array)
    curr_yr = datetime.datetime.today().year;

    array_idx = 0
    for time in time_array:
        time_str = time.strftime("%a %b. %-d")
        if time.year > curr_yr:
            print(time.year)
            time_str += ", "
            time_str += str(time.year)
        print(time_str)
        full_list_calendar_entries[array_idx][4] = time_str
        array_idx += 1

