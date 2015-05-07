from datetime import datetime
import logging
import os

import win32com.client as win32
from win32com.client import constants as cc

from .models import Payable, Receivable
from .constants import *


logger = logging.getLogger('root')


class ExcelAPI(object):
    def __init__(self, filename, worksheets):
        """Creates an ExcelAPI interface. Utilizes the win32com client for 
        communicating with Microsoft Excel. Opens the workbook and all 
        requested worksheets when instantiated.
        """
        self.client = win32.gencache.EnsureDispatch("Excel.Application")
        self.client.Visible = False

        filepath = os.path.join(os.getcwd(), 'workbook', filename)
        
        self.workbook = self.client.Workbooks.Open(filepath)
        self.worksheets = worksheets

        self.sheets = {
            'payables': Payables(self.sheet('payables')), 
            'receivables': Receivables(self.sheet('receivables')), 
        }

    def sheet(self, sheet_type):
        return self.workbook.Sheets(self.worksheets[sheet_type])

    def close(self, save_changes=True):
        self.workbook.Close(SaveChanges=save_changes)
        self.client.Application.Quit()


class ExcelSpreadsheet(object):
    def __init__(self, sheet):
        self.sheet = sheet
        self.first_row = 0

    def get_range(self, row_start, col_start, row_end, col_end):
        target = self.range(row_start, col_start, row_end, col_end)
        return target.Value

    # Internal methods
    @property
    def last_row(self):
        last_row = self.sheet.Rows.Count
        return self.cell(last_row, self.first_col).End(cc.xlUp).Row

    def cell(self, row, col):
        return self.sheet.Cells(row, col)

    def range(self, row_start, col_start, row_end, col_end):
        range_start = self.cell(row_start, col_start)
        range_end = self.cell(row_end, col_end)
        return self.sheet.Range(range_start, range_end)

    def set_cell(self, row, col, value):
        """Writes value to the cell at one-indexed (row, col)."""
        self.sheet.Cells(row, col).Value = value

    def set_range(self, row, col, contents):
        """Writes a block of contents to the spreadsheet."""
        width, height = len(contents[0]), len(contents)
        target = self.range(row, col, row+height-1, col+width-1)
        target.Value = contents


class Payables(ExcelSpreadsheet):
    def __init__(self, sheet):
        ExcelSpreadsheet.__init__(self, sheet)
        self.first_row = PAYABLE_FIRST_ROW
        self.first_col = PAYABLE_FIRST_COL
        self.last_col = PAYABLE_LAST_COL
        self.sort_col = PAYABLE_SORT_BY
        self.fields = PAYABLE_FIELDS


    def read_entries(self):
        entry_range = self.get_range(self.first_row, self.first_col+1, 
            self.last_row, self.last_col)
        
        entries = []
        for row in entry_range:
            entry = dict((field, row[idx]) for idx, field in enumerate(
                self.fields))
            entries.append(Payable(**entry))
        return entries

    def add_entries(self, entries):
        contents = []
        for idx, payable in enumerate(entries):
            billing_year = payable.data['event_date'].year
            # validation = self.entry_validation.format(row + idx)

            entry = [billing_year] + list(payable)
            contents.append(entry)

        if contents:
            self.set_range(self.last_row+1, self.first_col, contents)

    def sort_by_time(self):
        """Sorts the payables sheet by time."""
        target = self.range(self.first_row, self.first_col, self.last_row, 
            self.last_col)
        key = self.range(self.first_row, self.sort_col, self.last_row, 
            self.sort_col)
        target.Sort(
            Key1=key, 
            Order1=cc.xlAscending, 
            Orientation=cc.xlTopToBottom)

    def validate_payable(self, idx, transaction):
        transaction_values = ((
            transaction.data['id'], 
            transaction.data['timestamp'], 
            transaction.data['timestamp']
        ),)

        self.set_range(self.first_row+idx, PAYABLE_PAYPAL_ID_COL, 
            transaction_values)


class Receivables(ExcelSpreadsheet):
    def __init__(self, sheet):
        ExcelSpreadsheet.__init__(self, sheet)
        self.first_row = RECEIVABLE_FIRST_ROW
        self.first_col = RECEIVABLE_FIRST_COL
        self.last_col = RECEIVABLE_LAST_COL
        self.sort_col = RECEIVABLE_SORT_BY
        self.fields = RECEIVABLE_FIELDS

        
    def read_entries(self):
        entry_range = self.get_range(self.first_row, self.first_col, 
            self.last_row, self.last_col)
        
        entries = []
        for row in entry_range:
            entry = dict((field, row[idx]) for idx, field in enumerate(
                self.fields))
            entries.append(Receivable(**entry))
        return entries


    def add_entries(self, entries):
        contents = [list(entry) for entry in entries]

        if contents:
            self.set_range(self.last_row+1, self.first_col, contents)


    def sort_by_time(self):
        """Sorts the receivables sheet by time."""
        target = self.range(self.first_row, self.first_col, self.last_row, 
            self.last_col)
        key = self.range(self.first_row, self.sort_col, self.last_row, 
            self.sort_col)
        target.Sort(
            Key1=key, 
            Order1=cc.xlAscending, 
            Orientation=cc.xlTopToBottom)


