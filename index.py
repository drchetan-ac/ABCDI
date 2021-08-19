#from selenium import webdriver
#Written by Dr. Chetan A C
# Copyright with Dr. Chetan A C, all rights reserved.
# Read license before using or modifying the program.

from src.main import main_run
from easygui import boolbox
import logging
from selenium.webdriver.remote.remote_connection import LOGGER
LOGGER.setLevel(logging.WARNING)
logging.getLogger("urllib3").propagate = False

def main():
    main_run()
    if boolbox(
            "Completed with the data available." + '\n' + " Do you want to enter more data?" + '\n' + '\n' + ' Click Yes to start data entry with fresh results ' + '\n',
            title='Completed with the data entry. Click on Yes to enter more data', choices=('[Y]es', '[N]o'),
            image=None, default_choice='No',
            cancel_choice='No'):
        main_run()
    else:
        quit()
        exit()

if __name__ == "__main__":
    logging.basicConfig(filename='app.log', filemode='a', format='%(name)s - %(levelname)s - %(message)s', level='DEBUG')
    main()