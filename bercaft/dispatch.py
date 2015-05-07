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
from .api.constants import *

logger = logging.getLogger('root')


class Dispatch(object):
    """This object takes care of the actual updating for everything in the 
    spreadsheet.
    """
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
        """Loads the configuration file provided in the init."""
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
            
            logger.info("Fetching latest transactions from Paypal...")
            transactions = self.get_transactions()
            logger.info("Found transaction info for:\n* {} payables\n"
                "* {} receivables\n* {} transfers\n* {} other".format(
                    len(transactions['payables']), 
                    len(transactions['receivables']), 
                    len(transactions['transfers']), 
                    len(transactions['other'])))

            logger.info("Fetching reimbursements from Google form...")
            reimbursements = [r for r in self.sheets.all_payables()
                if r.data['timestamp'] > self.start_date]
            logger.info("{0} reimbursements found for {1}".format(
                len(reimbursements), self.start_date.year))
            test = set(reimbursements)

            logger.info("Updating payables spreadsheet")
            self.update_payables(transactions['payables'], reimbursements)

            logger.info("Updating receivables spreadsheet")
            self.update_receivables(transactions['receivables'])

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

            if not os.path.exists(os.path.join('workbook', 'backups')):
                os.makedirs(os.path.join('workbook', 'backups'))

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
        # Add new payables to sheet
        sheet = self.excel.sheets['payables']
        old_entries = set(sheet.read_entries())
        reimbursements = set(reimbursements)

        new_entries = reimbursements.difference(old_entries)
        sheet.add_entries(list(new_entries))
        logger.info("Added {} payables to spreadsheet".format(
            len(new_entries)))

        sheet.sort_by_time()

        # Verify payables against Paypal transactions
        transactions = dict((t.data['id'], t) for t in transactions)  
        
        # Remove all transactions already in spreadsheet
        all_entries = sheet.read_entries()
        id_range = [row[0] for row in sheet.get_range(
            sheet.first_row, PAYABLE_PAYPAL_ID_COL, 
            sheet.last_row, PAYABLE_PAYPAL_ID_COL)]

        for eid in id_range:
            if eid in transactions:
                transactions.pop(eid)

        logger.info("Found {} unmatched transactions".format(
            len(transactions)))

        # Check for matching payables and transactions
        matches = []
        for idx, (entry, eid) in enumerate(zip(all_entries, id_range)):
            for tid, transaction in transactions.items():
                if eid is None:
                    if self.match_transaction_payable(entry, transaction):
                        matches.append((idx, transaction))
                        transactions.pop(tid)
                        break   # Some messy stuff here

        for idx, transaction in matches:
            sheet.validate_payable(idx, transaction)


    def update_receivables(self, transactions):
        sheet = self.excel.sheets['receivables']

        new_receivables = [t.to_receivable() for t in transactions]

        old_receivables = dict((entry.data['transaction_id'], entry) for entry
            in sheet.read_entries())

        unadded = dict((e.data['transaction_id'], e) for e in new_receivables
            if e.data['transaction_id'] not in old_receivables)

        new_entries = sorted(unadded.values(), 
            key=lambda e: e.data['timestamp'])
        
        sheet.add_entries(new_entries)


    def match_transaction_payable(self, payable, transaction):
        payable_email = payable.data['paypal']
        transaction_email = transaction.data['email']

        if payable_email and transaction_email:
            if payable_email.lower() != transaction_email.lower():
                return False

        payable_amount = payable.data['amount']
        transaction_amount = transaction.data['amount']

        if payable_amount and transaction_amount:
            if abs(payable_amount) != abs(transaction_amount):
                return False

        return True


