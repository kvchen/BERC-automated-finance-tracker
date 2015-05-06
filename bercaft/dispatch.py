import logging
import os
import pythoncom
import pytz
import re
import shutil
import yaml

from datetime import datetime
from threading import Lock

from .api.paypal import PayPalAPI
from .api.sheets import SheetsAPI
from .api.excel import ExcelAPI

logger = logging.getLogger('root')


class Dispatch(object):
    def __init__(self, config_file):
        self.load_config(config_file)
        self.edit_lock = Lock()

        logger.info("Loading API interfaces")
        try:
            self.paypal = PayPalAPI(**self.config['paypal'])
            self.sheets = SheetsAPI(**self.config['sheets'])
        except:
            logger.exception("An interface failed to initialize!")
            raise

        self.start_date = self.config['global']['start_date'].replace(
            tzinfo=pytz.utc)

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

            logger.info("Ensuring that multithreading support is enabled")
            pythoncom.CoInitialize()

            logger.info("Opening Excel workbook for editing")
            self.excel = ExcelAPI(**self.config['excel'])
            

            # logger.info("Fetching latest transactions from Paypal...")
            # transactions = self.get_transactions()
            # logger.info("Found transaction info for:\n* {} payables\n"
            #     "* {} receivables\n* {} transfers\n* {} other".format(
            #         len(transactions['payables']), 
            #         len(transactions['receivables']), 
            #         len(transactions['transfers']), 
            #         len(transactions['other'])))

            logger.info("Fetching reimbursements from Google form...")
            reimbursements = [r for r in self.sheets.all_payables()
                if r.data['timestamp'] > self.start_date]
            logger.info("{0} reimbursements found for {1}".format(
                len(reimbursements), self.start_date.year))
            test = set(reimbursements)

            logger.info("Resolving deltas...")
            transactions = []
            # reimbursements = []

            logger.info("Updating payables spreadsheet")
            self.update_payables(transactions, reimbursements)

            logger.info("Making changes persistent")
            self.excel.close()

            if callback:
                callback()
        except:
            logger.exception("Failed to update spreadsheet")
            self.excel.close(save_changes=False)
            raise
        finally:
            self.edit_lock.release()


    def backup(self, callback=None):
        try:
            logger.info("Creating backup...")
            self.create_backup()
            
            if callback:
                callback()
        except:
            logger.exception("Failed to create backup")
            raise
        finally:
            self.edit_lock.release()


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
        end_date = datetime.now().replace(tzinfo=pytz.utc)
        transactions = self.paypal.get_transactions(self.start_date, end_date)

        processed = {
            'payables': [], 
            'receivables': [], 
            'transfers': [], 
            'other': [], 
        }

        for transaction in transactions:
            if transaction.data['type'] == 'Payment':
                if transaction.data['net_amount'] < 0:
                    transaction_type = 'payables'
                else:
                    transaction_type = 'receivables'
            elif transaction.data['type'] == 'Transfer':
                transaction_type = 'transfers'
            else:
                transaction_type = 'other'

            processed[transaction_type].append(transaction)

        return processed


    def update_payables(self, transactions, reimbursements):
        """Copies new payable entries into the workbook. Synchronizes with
        Google sheets, but preserves existing entries.
        """
        sheet = self.excel.sheets['payables']
        old_payables = set(sheet.read_entries())
        reimbursements = set(reimbursements)

        new_entries = reimbursements.difference(old_payables)

        sheet.add_entries(list(new_entries))
        sheet.sort_by_time()


