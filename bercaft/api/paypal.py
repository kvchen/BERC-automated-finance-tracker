""" Handles communications with PayPal API. """

from datetime import datetime, timedelta
from dateutil import parser
from urllib.parse import parse_qs

import logging
from pytz import timezone
import requests

from .models import Transaction


logger = logging.getLogger('root')

# Limit logging for requests library
requests_logger = logging.getLogger('requests')
requests_logger.setLevel(logging.WARNING)


class PayPalAPI(object):
    def __init__(self, username, password, signature, endpoint, api_version):
        self.username = username
        self.password = password
        self.signature = signature
        self.endpoint = endpoint
        self.api_version = api_version


    def call_api(self, method, args):
        """Sends a request to the PayPal NVP API using the given methods and
        arguments.
        """
        payload = {
            "USER": self.username, 
            "PWD": self.password, 
            "SIGNATURE": self.signature, 
            "METHOD": method,
            "VERSION": self.api_version,
        }

        for key, value in args.items():
            payload[key.upper()] = value

        response = requests.post(self.endpoint, data=payload)
        return parse_qs(response.text)


    def get_transactions(self, start_date, end_date):
        """Retrieves all PayPal transactions (rate-limited to 100 at a time, 
        with most recent first) and returns them as a chronologically-sorted
        array.
        """
        transaction_responses = {
            'L_STATUS': 'status', 
            'L_TYPE': 'type', 
            'L_TIMEZONE': 'timezone', 
            'L_TIMESTAMP': 'timestamp', 
            'L_TRANSACTIONID': 'id', 
            'L_NAME': 'name', 
            'L_EMAIL': 'email', 
            'L_AMT': 'amount', 
            'L_FEEAMT': 'fee_amount', 
            'L_NETAMT': 'net_amount', 
            'L_CURRENCYCODE': 'currency'
        }

        transactions = []

        while start_date < end_date:
            payload = {
                'STARTDATE': start_date.isoformat(), 
                'ENDDATE': end_date.isoformat()
            }

            response = self.call_api("TransactionSearch", payload)

            # Check if the response has any entries
            if 'L_STATUS0' not in response:
                break

            idx = 0
            while 'L_STATUS{0}'.format(idx) in response:
                fields = {}

                for rf, tf in transaction_responses.items():
                    field_name = '{0}{1}'.format(rf, idx)

                    if field_name in response:
                        fields[tf] = response[field_name][0]
                    else:
                        fields[tf] = None

                # Format the timestamp field as a datetime object
                fields['timestamp'] = parser.parse(fields['timestamp'])

                new_transaction = Transaction(**fields)
                transactions.append(new_transaction)
                idx += 1

            next_end_date = transactions[-1].timestamp
            if next_end_date == end_date:
                break

            logger.info("Retrieved transaction history from {} to {}"
                .format(next_end_date, end_date))
            end_date = next_end_date

        transactions.reverse()
        return transactions


