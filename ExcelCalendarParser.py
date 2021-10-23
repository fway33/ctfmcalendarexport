# This python file will parse the calendar dump excel spreadsheet
# and produce a text output suitable for editing.
import openpyxl
import numpy

def get_populated_row_count(ws):
    number_of_rows = ws.max_row
    last_row_index_with_data = 0

    while True:
        if ws.cell(number_of_rows, 3).value != None:
            last_row_index_with_data = number_of_rows
            break
        else:
            number_of_rows -= 1
    return number_of_rows

def export_calendar_data() :
    print("Here is where we'd start to parse")

    wb = openpyxl.load_workbook('CalDump1.xlsx')
    print(wb.get_sheet_names())
    ws = wb['data']
    rows = get_populated_row_count(ws)
    #print(get_populated_row_count(ws))
    columns = ws.max_column
    print(ws['A1'].value," A1")
    print(ws['B1'].value," B1")
    print(ws['C1'].value," C1")
    print(ws['D1'].value," D1")
    print(ws['E1'].value," E1")

    print(type(ws['A1':'E5']))
    end = 'E' + str(rows)
    table = numpy.array([[cell.value for cell in col] for col in ws['A2':end]])
    print(table)
    print ("\n----------\n")
    modify_lodge(table)

def modify_lodge(table):
    #print(table[0:-1,1][1])
    print([item[1] for item in table])


