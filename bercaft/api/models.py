import pytz

from datetime import datetime
from dateutil import parser

from .constants import *


class Entry(object):
    def __init__(self, **kwargs):
        self.data = {}
        self.ignored_fields = ()

        for field in self.fields:
            self.data[field] = self.clean_field(field, kwargs.get(field))


    def clean_field(self, field, value):
        if value == '':
            return None
        return value


    def __iter__(self):
        for field in self.fields:
            yield self.data[field]


    def __eq__(self, other):
        if not isinstance(other, Entry):
            return False

        for field in self.data:
            if field not in self.ignored_fields:
                if self.data[field] != other.data[field]:
                    return False
        return True


    def __hash__(self):
        return hash(self.data['timestamp'])


class Payable(Entry):
    """Creates a new Payable object using an entry that starts at the
    timestamp and ends at driving reimbursement.
    """
    def __init__(self, **kwargs):
        self.fields = PAYABLE_FIELDS
        Entry.__init__(self, **kwargs)

        self.ignored_fields = PAYABLE_IGNORE_FIELDS


    def clean_field(self, field, value):
        if value is None or value == '':
            return None

        if isinstance(value, str):
            value = value.strip()

        if field in ('timestamp', 'event_date'):
            if isinstance(value, str):
                if field == 'timestamp':
                    date_format = "%m/%d/%Y %H:%M:%S"
                elif field == 'event_date':
                    date_format = "%m/%d/%Y"
                value = datetime.strptime(value, date_format)
            else:
                value = datetime(value.year, value.month, value.day, 
                    value.hour, value.minute, value.second)

            value = pytz.utc.localize(value)
        elif field in ('paypal',):
            value = value.lower()
        elif field in ('amount',):
            if isinstance(value, str):
                value = value.replace('$', '').replace(',', '')
            value = float(value)

        return value


class Receivable(Entry):
    """Creates a new Receivable object using an entry that starts at the
    year and ends at the transaction id.
    """
    def __init__(self, **kwargs):
        self.fields = RECEIVABLE_FIELDS
        Entry.__init__(self, **kwargs)


class Transaction():
    """Creates a new Transaction object using an entry that starts at the
    status and ends at currency type.
    """
    def __init__(self, **kwargs):
        self.fields = TRANSACTION_FIELDS
        Entry.__init__(self, **kwargs)


    def clean_field(self, field, value):
        if value is None:
            return None

        if field in ('amount', 'net_amount', 'fee_amount'):
            value = float(value)
        elif field in ('timestamp',):
            value = parser.parse(value)
        elif field in ('email',):
            value = value.lower()
        elif field in ('name',):
            value = value.title()

        return value


