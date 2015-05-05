from collections import namedtuple

payable_fields = [
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

receivable_fields = [
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

transaction_fields = [
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


Payable = namedtuple('Payable', payable_fields)
Receivable = namedtuple('Receivable', receivable_fields)
Transaction = namedtuple('Transaction', transaction_fields)


