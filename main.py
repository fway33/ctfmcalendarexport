#
# This is the main program.  From here we will call the
# export calendar data functions.
#
import logging

from export_calendar_parser import export_calendar_data
from export_doc import create_word_doc


def main() -> None:
    logger = logging.getLogger(__name__)
    logger.debug("Main function called")
    export_calendar_data()
    create_word_doc()

if __name__ == "__main__":
    logging.basicConfig(filename='calparser.log', encoding='utf-8', level=logging.DEBUG)

    main()
