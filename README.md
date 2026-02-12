CTFM Calendar Export

This is a very specific project for a very specific purpose.

Each month data is exported from a Google spreadsheet from all the lodge calendars in the state.  The export creates an Excel spreadsheet.  This spreadsheet contains data that will 
feed a Word document that lists statewide lodge events.

The data is exported via openpyxl, and only the first five columns of the calendar data are of interest.  Once the entire spreadsheet is extracted in a list of lists, each column is iterated over
and any data massaging is accomplished.   For example, lodge names will show as "Lodge 001 Hiram," when the document needs to show it as "Hiram Lodge No. 1".  Two of the event columns (title, description) are
search for the word "dinner" or "degree" and the index of that row is saved in a dinner list or a degree list. Those lists are used later when creating the Word document to add a dinner or degree marker
for that event.

The toughest column is event location.  A Connecticut town needs to be extracted from address data that may have been entered in a variety of ways.  Regular expressions, a dictionary of lodges/locations, and a known bad data list (the data has been seen over previous months and can be checked) are all used to try to extract a town.  The location data is also preserved in case it is an even that does not take place at a lodge -- a pizza place or bowling alley for example -- and that will be put into the Word document as well.  Because of the outside location possibility, we cannot just grab a location from the lodge town dictionary.

Finally, all of the data structures are exported into a Word document.  This still needs a bit of manual editing before it can go into the monthly newspaper, but this process has cut the work by 90%.

Currently, the exported spreadsheet is put into the project directory.  Future enhancments will allow for specifying an input xlsx file location and an output docx file location.
