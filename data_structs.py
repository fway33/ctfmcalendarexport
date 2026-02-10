# These are the data structures that are needed in this whole utility.

from enum import Enum
class Col(Enum):
    LODGE = 0
    TITLE = 1
    DESCR = 2
    LOCATION = 3
    DATE = 4

# LodgeTown is a dictionary of lodge names (as they come from Google calendar into the spreadsheet
# and their towns.
lodge_town = {"Lodge 001 Hiram": "New Haven",
             "Lodge 002 St. John's": "Middletown",
             "Lodge 003 Fidelity-St. John's": "Fairfield",
             "Lodge 004 Wyllys-St. John's": "West Hartford",
             "Lodge 005 Union": "Stamford",
             "Lodge 006 Old Well-St. John's": "Norwalk",
             "Lodge 007 King Solomon's": "Woodbury",
             "Lodge 008 America-St. John's": "Stratford",
             "Lodge 009 Compass": "Wallingford",
             "Lodge 010 Wooster": "Colchester",
             "Lodge 011 St. Paul's": "Litchfield",
             "Lodge 012 King Hiram": "Shelton",
             "Lodge 013 Montgomery": "Lakeville",
             "Lodge 014 Frederick-Franklin": "Plainville",
             "Lodge 015 Moriah": "Brooklyn",
             "Lodge 016 Temple": "Chesire",
             "Lodge 017 Federal": "Waterbury",
             "Lodge 018 Hiram": "Sandy Hook",
             "Lodge 019 Washington": "Monroe",
             "Lodge 021 St. Peter's": "New Milford",
             "Lodge 022 Trumbull": "New Haven",
             "Lodge 024 Uriel": "Merrow",
             "Lodge 025 Columbia": "South Glastonbury",
             "Lodge 028 Composite": "Suffield",
             "Lodge 029 Village": "Collinsville",
             "Lodge 030 Day Spring": "Hamden",
             "Lodge 031 Union": "Niantic",
             "Lodge 033 Friendship": "Southington",
             "Lodge 034 Somerset-St. James": "Preston",
             "Lodge 036 Valley": "Simsbury",
             "Lodge 038 St. Alban's": "Branford",
             "Lodge 039 Ark": "Danbury",
             "Lodge 040 Union": "Danbury",
             "Lodge 042 Harmony": "Waterbury",
             "Lodge 043 Estuary": "Old Saybrook",
             "Lodge 046 Putnam": "Woodstock",
             "Lodge 047 Morning Star": "Seymour",
             "Lodge 048 St. Luke's": "Kent",
             "Lodge 049 Jerusalem": "Ridgefield",
             "Lodge 051 Warren": "Portland",
             "Lodge 055 Seneca": "Torrington",
             "Lodge 057 Coastal": "Stonington",
             "Lodge 060 Wolcott": "Stafford",
             "Lodge 063 Corinthian": "North Haven",
             "Lodge 064 St. Andrew's": "Winsted",
             "Lodge 065 Temple": "Westport",
             "Lodge 066 Widow's Son": "Branford",
             "Lodge 067 Harmony": "New Canaan",
             "Lodge 069 Fayette": "Ellington",
             "Lodge 070 Washington": "Windsor",
             "Lodge 073 Manchester": "Manchester",
             "Lodge 076 Liberty-Continental": "Waterbury",
             "Lodge 077 Meridian": "Meriden",
             "Lodge 078 Shepherd-Salem": "Naugatuck",
             "Lodge 079 Wooster": "New Haven",
             "Lodge 081 Washington": "Cromwell",
             "Lodge 085 Acacia": "Cos Cob",
             "Lodge 087 Madison": "Madison",
             "Lodge 088 Hartford Evergreen": "South Windsor",
             "Lodge 089 Ansantawae": "Milford",
             "Lodge 095 Jeptha": "Clinton",
             "Lodge 097 Center": "Meriden",
             "Lodge 101 Evening Star": "Unionville",
             "Lodge 102 Brainard": "Niantic",
             "Lodge 104 Corinthian": "Stratford",
             "Lodge 107 Ivanhoe": "Darien",
             "Lodge 110 Ionic": "North Windham",
             "Lodge 112 Anchor": "East Hampton",
             "Lodge 113 Moosup": "Moosup",
             "Lodge 115 Annawon": "West Haven",
             "Lodge 116 Oxoboxo": "Montville",
             "Lodge 120 Bayview": "Niantic",
             "Lodge 122 Corner Stone-Quinebaug": "Thompson",
             "Lodge 119 Granite": "Haddam",
             "Lodge 125 Cosmopolitan": "New Haven",
             "Lodge 128 Hospitality": "Wethersfield",
             "Lodge 131 Solar": "East Hampton",
             "Lodge 140 Sequin-Level": "Newington",
             "Lodge 142 Ashlar-Aspetuck": "Easton",
             "Lodge 144 Daytime": "Stratford",
             "Lodge 145 Friendship-Tuscan": "Manchester",
             "Lodge 146 Wolcott": "Wolcott",
             "Lodge 148 Unity": "New Britain",
             "Lodge 149 Universal Fraternity": "Stratford",
             "Lodge 150 Purity": "Mystic",
             "Lodge 151 Ouroboros": "Wallingford",
             "Lodge 332 Ashlar": "Wallingford",
             "Lodge 500 Quinta Essentia": "New Haven"
             }

# Sometimes the calendar entries are way off but they're consistent for the specific lodge
# so let's roll an exception dictionary
bad_locations = {"King Hiram Lodge": "Shelton",
                "Main Lodge Room": "Hamden", # be careful with this one.....
                "Warren Lodge #51": "Portland",
                "Fellowcraft Club - Lodge of Instruction": "Portland",
                "Brainard Lodge #102": "Niantic",
                "Annawon Lodge 263 Center St. WH": "West Haven",
                }

# These sequences contain the row numbers if the row contains "degree" or "dinner" information
# degrees and dinner are sets -- the word may appear in the event title or in the event description.  The row number
# only needs to be saved once
degrees = set()
has_dinner = set()

lodge_locations = []
location_exceptions = []
full_list_calendar_entries = [[]]

