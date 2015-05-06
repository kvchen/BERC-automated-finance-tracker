from datetime import datetime
import logging
import os

import win32com.client as win32
from win32com.client import constants as cc

from .models import Payable, Receivable


logger = logging.getLogger('root')


class ExcelAPI(object):
    def __init__(self, filename, worksheets):
        self.client = win32.gencache.EnsureDispatch("Excel.Application")
        self.client.Visible = False

        filepath = os.path.join(os.getcwd(), 'workbook', filename)
        
        self.workbook = self.client.Workbooks.Open(filepath)
        self.sheets = {
            key: self.workbook.Sheets(name) 
            for key, name in worksheets.items()
        }


    def close(self, save_changes=True):
        self.workbook.Close(SaveChanges=save_changes)
        self.client.Application.Quit()


class ExcelSpreadsheet(object):
    def __init__(self):
        pass


def read_payable(spreadsheet, idx):
    pass


def read_receivable(spreadsheet, idx):
    pass


def write_payable(spreadsheet, payable):
    pass


def write_receivable(spreadsheet, receivable):
    pass


