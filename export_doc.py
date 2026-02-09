from docx import Document

from DataStructs import locations, location_exceptions, degrees, full_list_calendar_entries, dinner


def create_word_doc():

    #create a document object
    document = Document()
    document.add_heading("Calendar Events")

    # full_list_calendar_entries holds all the massaged data from the cells.
    # We will loop through this and write out lines in to the word doc.  The idea is
    # that most of them will be very close to the final format required.  It will be
    # easy to tweak locations etc.
    array_idx = 0;
    for item in full_list_calendar_entries:
        if item[0] == -1:
            continue

        # Each paragraph is one event.
        p = document.add_paragraph()

        # For this paragraph (an event) first we want to see if there is a
        # dinner and/or degree.  Check those arrays first to see if the current array_idx matches
        # anything from the dinner or degrees lists.
        if array_idx in degrees:
            p.add_run('*** DEGREE *** ')
        if array_idx in dinner:
            p.add_run(" DINNER ")

        # Now put the time in, in bold:
        p.add_run(item[4]).bold = True
        p.add_run(', ')

        # Need the lodge name
        p.add_run(item[0])
        run = p.add_run()
        run.add_break()

        # Next up is location.  If there is something in the locations list for this array index
        # print that.  Then put in the full location from the full_list_calendar_entries, delimited by '|'
        if locations[array_idx]:
            p.add_run(locations[array_idx])
        p.add_run('| ')
        p.add_run(item[3])
        p.add_run(' |')

        # Next up event title
        p.add_run(item[1])
        run.add_break()
        # And finally the event description.
        p.add_run(item[2])

        array_idx += 1

    # After writing all that, save the document.
    document.save('scratch.docx')
