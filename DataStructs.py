LodgeTown = {"Lodge 001 Hiram": "New Haven",
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
             "Lodge 014 Frederick-Franklin": "Plainville",
             "Lodge 016 Temple": "Chesire",
             "Lodge 017 Federal": "Waterbury",
             "Lodge 018 Hiram": "Sandy Hook",
             "Lodge 019 Washington": "Monroe",
             "Lodge 021 St. Peter's": "New Milford",
             "Lodge 028 Composite": "Suffield",
             "Lodge 031 Union": "Niantic",
             "Lodge 033 Friendship": "Southington",
             "Lodge 038 St. Alban's": "Branford",
             "Lodge 039 Ark": "Danbury",
             "Lodge 040 Union": "Danbury",
             "Lodge 042 Harmony": "Waterbury",
             "Lodge 043 Estuary": "Old Saybrook",
             "Lodge 047 Morning Star": "Seymour",
             "Lodge 049 Jerusalem": "Ridgefield",
             "Lodge 055 Seneca": "Torrington",
             "Lodge 060 Wolcott": "Stafford",
             "Lodge 064 St. Andrew's": "Winsted",
             "Lodge 065 Temple": "Westport",
             "Lodge 066 Widow's Son": "Branford",
             "Lodge 067 Harmony": "New Canaan",
             "Lodge 069 Fayette": "Ellington",
             "Lodge 081 Washington": "Cromwell",
             "Lodge 085 Acacia": "Cos Cob",
             "Lodge 087 Madison": "Madison",
             "Lodge 095 Jeptha": "Clinton",
             "Lodge 104 Corinthian": "Stratford",
             "Lodge 112 Anchor": "East Hampton",
             "Lodge 115 Annawon": "West Haven",
             "Lodge 119 Granite": "Haddam",
             "Lodge 125 Cosmopolitan": "New Haven",
             "Lodge 128 Hospitality": "Wethersfield",
             "Lodge 131 Solar": "East Hampton",
             "Lodge 140 Sequin-Level": "Newington",
             "Lodge 148 Unity": "New Britain",
             "Lodge 149 Universal Fraternity": "Stratford",
             "Lodge 150 Purity": "Mystic",
             "Lodge 332 Ashlar": "Wallingford"
             }

# Sometimes the calendar entries are way off but they're consistent for the specific lodge
# so let's roll an exception dictionary
BadLocations = {"King Hiram Lodge": "Shelton",
                "Main Lodge Room": "Hamden", # be careful with this one.....
                "Warren Lodge #51": "Portland",
                "Fellowcraft Club - Lodge of Instruction": "Portland",
                "Brainard Lodge #102": "Niantic",
                "Annawon Lodge 263 Center St. WH": "West Haven",
                }

degrees = []
dinner = []
locations = []
location_exceptions = []
full_list_calendar_entries = [[]]
