from datetime import datetime
from dateutil import parser

import pytz
import logging
import yaml

from .api.paypal import PayPalAPI
from .api.sheets import SheetsAPI

logger = logging.getLogger('root')


class Dispatch(object):
    def __init__(self, config_file):
        self.load_config(config_file)

        logger.info("Loading API interfaces")
        try:
            self.paypal = PayPalAPI(**self.config['paypal'])
            self.sheets = SheetsAPI(**self.config['sheets'])
        except:
            logger.exception("An interface failed to initialize")
            raise

        logger.info("Dispatch object created successfully")

    
    def load_config(self, config_file):
        try:
            with open(config_file, 'r') as infile:
                self.config = yaml.load(infile)
        except:
            logger.exception("Failed to open configuration file")
            raise

        logger.info("Configuration file loaded successfully")


    def update(self, callback=None):
        try:
            logger.info("Starting updates...")

            logger.info("Fetching latest transactions from Paypal...")
            transactions = self.get_transactions()
            logger.info("{} transactions found".format(len(transactions)))

            if callback:
                callback()
        except:
            logger.exception("Failed to update spreadsheet")
            raise


    def backup(self, callback=None):
        try:
            logger.info("Starting backup...")
            
            
            if callback:
                callback()
        except:
            logger.exception("Failed to create backup")
            raise


    def get_transactions(self):
        start_date = self.config['global']['start_date'].replace(
            tzinfo=pytz.utc)
        end_date = datetime.now().replace(tzinfo=pytz.utc)
        return self.paypal.get_transactions(start_date, end_date)


