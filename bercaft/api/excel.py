from datetime import datetime
import logging

import win32com.client as win32
from win32com.client import constants as cc

from .models import Payable, Receivable


logger = logging.getLogger('root')


class ExcelWorkbook(object):
    def __init__(self, filename):
        self.client = win32.gencache.EnsureDispatch("Excel.Application")
        self.client.Visible = False

        self.workbook = self.client.Workbooks.Open(filename)


    def spreadsheet(self, name):
        try:
            return self.workbook.Sheets(name)
        except:
            raise KeyError("Sheet with name '{}' not found".format(name))


    def close(self, save_changes=False):
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


