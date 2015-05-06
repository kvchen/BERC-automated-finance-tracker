# !/usr/bin/env python
# coding: utf-8

"""
BERC Automated Finance Tracker
A script to facilitate the process of data formatting and aggregation for the 
    Berkeley Energy and Resources Collaborative.
"""

__author__ = "Kevin Chen"
__maintainer__ = "Kevin Chen"
__email__ = "kvchen@berkeley.edu"
__repository__ = "https://github.com/kvchen/BERC-automated-finance-tracker"
__status__ = "Development"
__version__ = "0.0.1"



# Default modules
from bercaft.dispatch import Dispatch
from bercaft.gui import DispatchGUI

import argparse
import logging

from logging.handlers import TimedRotatingFileHandler

LOG_FILENAME = 'logs/bercaft.log'
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(module)s - %(message)s'


def main():
    parser = argparse.ArgumentParser(
        description="Start the BERC Automated Finance Tracker")
    parser.add_argument("-b", "--backup", action="store_true", 
        help="creates a backup of the main spreadsheet")
    parser.add_argument("-u", "--update", action="store_true", 
        help="updates the spreadsheet")
    parser.add_argument("-d", "--debug", action="store_true", 
        help="enable debugging output")
    
    args = parser.parse_args()

    # Log to (up to) three places - GUI, logfile, and console output

    # Uncomment this line to log output to terminal
    logging.basicConfig(format=LOG_FORMAT)

    formatter = logging.Formatter(LOG_FORMAT)
    logger = logging.getLogger('root')
    logger.setLevel(logging.INFO)

    handler = TimedRotatingFileHandler(LOG_FILENAME, when='H')
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    if args.debug:
        logger.setLevel(logging.DEBUG)
        logger.debug("Debugging output enabled")

    try:
        dispatch = Dispatch('config.yml')
        
        if args.backup:
            dispatch.backup()
        elif args.update:
            dispatch.update()
        else:
            gui = DispatchGUI(dispatch)
            gui.mainloop()
            logger.removeHandler(gui.console_handler)
    except Exception as e:
        logger.info("Fatal error encountered!")
    finally:
        logger.info("Cleaning up...")


if __name__ == "__main__":
    main()

