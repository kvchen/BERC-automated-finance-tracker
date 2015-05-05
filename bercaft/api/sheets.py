import gspread

from .models import Payable


class SheetsAPI(object):
    def __init__(self, username, password, workbook):
        self.client = gspread.login(username, password)
        self.workbook = self.client.open(workbook)
        self.sheet = self.workbook.get_worksheet(0)


    def all_payables(self):
        """Returns a list of all payables extracted from the reimbursement
        form.
        """
        payables = []

        for row in self.sheet.get_all_values()[1:]:
            row = row[:15]
            if len(row) < 15:
                row += [None] * range(len(Payable._fields)-len(row))

            payables.append(Payable(*row))

        return payables


