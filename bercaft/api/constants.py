# Payable-related constants
PAYABLE_FIRST_ROW = 20
PAYABLE_FIRST_COL = 2
PAYABLE_LAST_COL = 25
PAYABLE_SORT_BY = 3
PAYABLE_PAYPAL_ID_COL = 18

PAYABLE_FIELDS = [
    'timestamp', 
    'requester', 
    'department', 
    'item', 
    'detail', 
    'event_date', 
    'payment_type', 
    'use_of_funds', 
    'notes', 
    'type', 
    'name', 
    'paypal', 
    'address', 
    'amount', 
    'driving_reimbursement'
]

PAYABLE_IGNORE_FIELDS = ('detail', 'notes')


# Receivable-related constants
RECEIVABLE_FIRST_ROW = 11
RECEIVABLE_FIRST_COL = 2
RECEIVABLE_LAST_COL = 25
RECEIVABLE_SORT_BY = 4

RECEIVABLE_FIELDS = [
    'year', 
    'committed_date', 
    'timestamp', 
    'support_type', 
    'organization_type', 
    'budget_line_item', 
    'payee_name', 
    'payee_email', 
    'amount_requested', 
    'amount_committed', 
    'amount_gross', 
    'amount_net', 
    'transaction_id'
]


# Transaction-related constants
TRANSACTION_FIELDS = [
    'status', 
    'type', 
    'timezone', 
    'timestamp', 
    'id', 
    'name', 
    'email', 
    'amount', 
    'fee_amount', 
    'net_amount', 
    'currency'
]

TRANSACTION_RESPONSE_KEYS = {
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


PAYABLE_TRANSACTION_MATCHES = (
    ('paypal', 'email'), 
    ('amount', 'amount')
)