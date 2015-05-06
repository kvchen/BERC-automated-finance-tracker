import logging
import os
import pytz
import re
import shutil
import yaml

from datetime import datetime

from .api.paypal import PayPalAPI
from .api.sheets import SheetsAPI
from .api.excel import ExcelAPI

logger = logging.getLogger('root')


class Dispatch(object):
    def __init__(self, config_file):
        self.load_config(config_file)

        logger.info("Loading API interfaces")
        try:
            self.paypal = PayPalAPI(**self.config['paypal'])
            self.sheets = SheetsAPI(**self.config['sheets'])
            self.excel = ExcelAPI(**self.config['excel'])
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
            logger.info("Creating backup...")
            self.create_backup()
            
            if callback:
                callback()
        except:
            logger.exception("Failed to create backup")
            raise


    def create_backup(self):
        filename = self.config['excel']['filename']
        src = os.path.join('workbook', filename)

        name, ext = os.path.splitext(filename)
        date_str = datetime.now().strftime("%Y.%m.%d")

        backup_dir = os.path.join('workbook', 'backups')
        backup_filename = '{0}_{1}{2}'.format(name, 
            datetime.now().strftime("%Y.%m.%d"), ext)

        dst = os.path.join(backup_dir, backup_filename)
        shutil.copy2(src, dst)

        logger.info("Created backup at {}".format(dst))


    def get_transactions(self):
        start_date = self.config['global']['start_date'].replace(
            tzinfo=pytz.utc)
        end_date = datetime.now().replace(tzinfo=pytz.utc)
        return self.paypal.get_transactions(start_date, end_date)


    def cleanup(self):
        self.excel.close(save_changes=False)



