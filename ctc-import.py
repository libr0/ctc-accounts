#!/usr/bin/python3
#
#   ctc-import.py
#
#   Imports Clubspark and Stripe data for Coburg Tennis Club into the Gnucash accounting file.
#
#   Written by Chris Boek (c) August 2022
#

import csv
import sys
import argparse
import re

from decimal import Decimal

from os.path import abspath
from sys import argv, exit
import datetime
from datetime import timedelta, date, datetime
from gnucash import Session, Account, Transaction, Split, Query, GncNumeric, SessionOpenMode
from gnucash.gnucash_business import Customer, Employee, Vendor, Job, \
    Address, Invoice, Entry, TaxTable, TaxTableEntry, GNC_AMT_TYPE_PERCENT, \
    GNC_DISC_PRETAX
from gnucash.gnucash_core_c import \
    ACCT_TYPE_ASSET, ACCT_TYPE_RECEIVABLE, ACCT_TYPE_INCOME, \
    GNC_OWNER_CUSTOMER, ACCT_TYPE_LIABILITY

TESTING=True

COURT_BOOKING = 'Court booking at Coburg Tennis Club'
MEMBERSHIP_PREFIX = 'Coburg Tennis Club:'
STRIPE_PAYOUT = 'STRIPE PAYOUT'
REFUND_PREFIX = 'REFUND FOR CHARGE'
EVENT_PREFIX = 'Coburg Tennis Club '
NO_NAME_ENTRY = 'NoName'

def GetAllCustomers(book):
    """Returns all customers in book.
    Posts a query to search for all customers.
    arguments:
        book                the gnucash book to work with
    """

    query = Query()
    query.search_for('gncCustomer')
    query.set_book(book)

    customer_list = []

    for result in query.run():
        customer_list.append(Customer(instance=result))

    query.destroy()

    return customer_list

def EnterBacInvoices(cs_csv, book):
    court_hire_acct = GetCourtHireAcct(book)
    light_hire_acct = GetLightHireAcct(book)
    cs_map = {} 
    cs_map[NO_NAME_ENTRY] = []
    for row in cs_csv:
        player_name = row['Player First Name'] + ' ' + row['Player Last Name']
        if player_name:
            if player_name in cs_map:
                player_list = cs_map[player_name]
                player_list.append(row)
            else:
                cs_map[player_name] = [row]
        else:
            print('CS: No name for row: ' + row + '\n')
            cs_map[NO_NAME_ENTRY].append(row)
            

        date_str = row['Booked Date']
        txn_date = date.fromisoformat(date_str)
        description = "Casual Hire"

        court_hire = Decimal(row['Court Fee'])
        light_hire = Decimal(row['Light Fee'])

        if (court_hire == 0) & (light_hire == 0):
            continue

#        print("Court Hire: $%s" % court_hire)
#        print("Light Hire: $%s" % light_hire)

        # create the BAC invoice
        book_a_court = GetBookACourtCustomer(book)
        invoice_id = book.InvoiceNextID(book_a_court)
        AUD = GetAUDCurrency(book)
        invoice = Invoice(book, invoice_id, AUD, book_a_court, txn_date)
        if court_hire > 0:
            ch_invoice_entry = Entry(book, invoice, txn_date)
            ch_invoice_entry.SetDescription("Court Hire")
            ch_invoice_entry.SetQuantity(GncNumeric(1))
            ch_invoice_entry.SetInvPrice(GncNumeric(row['Court Fee']))
            ch_invoice_entry.SetInvAccount(court_hire_acct)

        if light_hire > 0:
            lh_invoice_entry = Entry(book, invoice, txn_date)
            lh_invoice_entry.SetDescription("Light Hire")
            lh_invoice_entry.SetQuantity(GncNumeric(1))
            lh_invoice_entry.SetInvPrice(GncNumeric(row['Light Fee']))
            lh_invoice_entry.SetInvAccount(light_hire_acct)

        receivables = GetReceivablesAcct(book)
        invoice.PostToAccount(receivables, txn_date, txn_date,"", True, False)

        row['Invoice'] = invoice
        row['Allocated'] = False

    return cs_map

def GetCustomer(book, name, email):
    cust_list = GetAllCustomers(book)
    customer = None
    for cust in cust_list:
        if cust.GetName() == name:
            customer = cust
            break
    if not customer:
        # create one
        cust_id = book.CustomerNextID()
        AUD = GetAUDCurrency(book)
        customer = Customer(book, cust_id, AUD, name)
        address = customer.GetAddr()
        address.SetEmail(email)

    return customer

def CreateInvoiceAndPayment(book, customer, acct, amount, fee, net, description, txn_date):
    # Add an invoice and payment for that customer
    AUD = GetAUDCurrency(book)
    cust_invoice_id = book.InvoiceNextID(customer)
    cust_invoice = Invoice(book, cust_invoice_id, AUD, customer)
    invoice_entry = Entry(book, cust_invoice, txn_date)
    invoice_entry.SetDescription(description)
    invoice_entry.SetQuantity(GncNumeric(1))
    invoice_entry.SetInvPrice(GncNumeric(amount))
    invoice_entry.SetInvAccount(acct)

    receivables = GetReceivablesAcct(book)
    cust_invoice.PostToAccount(receivables, txn_date, txn_date,"", True, False)

    # Apply payment
    trans = MakeStripePayment(book, amount, fee, net, description, txn_date)

    # now associate the payment with the correct invoice
    # TODO make it deal properly with multiple bookings by a customer in a month
    stripe_acct = GetStripeAcct(book)
    cust_invoice.ApplyPayment(trans, stripe_acct, GncNumeric(amount), \
                    GncNumeric(1), txn_date, None, None)
    RecordStripeFee(book, trans, net, fee)


def MakeStripePayment(book, amount, fee, net, description, txn_date):
    # process the payment
    stripe_acct = GetStripeAcct(book)
    stripe_fee_acct = GetStripeFeeAcct(book)
    receivables = GetReceivablesAcct(book)

    trans = Transaction(book)
    trans.BeginEdit()
    split1 = Split(book)
    split1.SetAccount(stripe_acct)
    split1.SetValue(GncNumeric(amount))
    split1.SetParent(trans)

    AUD = GetAUDCurrency(book)
    trans.SetCurrency(AUD)
    trans.SetDate(txn_date.day, txn_date.month, txn_date.year)
    trans.SetDescription(description)
    trans.CommitEdit()

    return trans

def DoStripeTransfer(book, amount, txn_date):
    # process a transfer from stripe account to checking account
    stripe_acct = GetStripeAcct(book)
    checking_acct = GetCheckingAcct(book)
    trans = Transaction(book)
    trans.BeginEdit()
    split1 = Split(book)
    split1.SetAccount(stripe_acct)
    split1.SetValue(GncNumeric(amount))
    split1.SetParent(trans)
    split2 = Split(book)
    split2.SetAccount(checking_acct)
    split2.SetValue(GncNumeric(amount).neg())
    split2.SetParent(trans)
    AUD = GetAUDCurrency(book)
    trans.SetCurrency(AUD)
#    trans.SetDate(txn_date)
    trans.SetDate(txn_date.day, txn_date.month, txn_date.year)
    trans.SetDescription("Stripe Payout")
    trans.CommitEdit()

# Edit the transaction to add the stripe fee split
def RecordStripeFee(book, trans, net, fee):
    if not fee:
        return
    stripe_fee_acct = GetStripeFeeAcct(book)
    trans.BeginEdit()
    sl = trans.GetSplitList()
    for split in sl:
        if split.GetAccount().GetName() == 'Stripe Account':
            split.SetValue(GncNumeric(net))
    s2 = Split(book)
    s2.SetAccount(stripe_fee_acct)
    s2.SetValue(GncNumeric(fee))
    s2.SetParent(trans)
    trans.CommitEdit()


def ProcessStripePayments(stripe_csv, book, cs_map):
    stripe_acct = GetStripeAcct(book)
    stripe_fee_acct = GetStripeFeeAcct(book)
    checking_acct = GetCheckingAcct(book)
    AUD = GetAUDCurrency(book)
#    st_map = {}
    for row in stripe_csv:
        email_address = row['Email address (metadata)']
        customer_id = row['Contact ID (metadata)'] # customer UUID
        mobile_number = row['Mobile number (metadata)']
        if not mobile_number:
            mobile_number = row['Phone number (metadata)']

        # lookup by email address
#        st_map[email_address] = row

        first_name = row['First name (metadata)']
        if not first_name:
            first_name = row['First Name (metadata)']
        if not first_name:
            first_name = row['FirstName (metadata)']

        last_name = row['Last name (metadata)']
        if not last_name:
            last_name = row['Last Name (metadata)']
        if not last_name:
            last_name = row['Surname (metadata)']

        # lookup by first+last name
        player_name = first_name + ' ' + last_name
        if not player_name:
            player_name = NO_NAME_ENTRY

        if player_name in cs_map:
            player_list = cs_map[player_name]

#        if player_name:
#            st_map[player_name] = row
#            if player_name in cs_map:
#                player_list = cs_map[player_name]
#        else:
#            player_name = NO_NAME_ENTRY
#            player_list = cs_map[NO_NAME_ENTRY]

        date_str = row['Available On (UTC)'].split(' ')[0]
    #    year, month, day = [int(f) for f in date_str.split('-')]

#        print("Date String %s" % date_str)

        txn_date = date.fromisoformat(date_str)

#        print("Date " + str(txn_date.year) + str(txn_date.month) + str(txn_date.day) )

        amount = row['Amount']
        fee = row['Fee']
        net = row['Net']

#        print("Amount " + amount + " Fee " + fee + " Net " + net)

        # don't need to record zero fee transactions
        if (amount == 0):
            continue

        description = row['Description']
        txn_type = row['Type']

        if description == COURT_BOOKING:
            # process the payment
            trans = MakeStripePayment(book, amount, fee, net, description, txn_date)

            # now associate the payment with the correct invoice
            # TODO make it deal properly with multiple bookings by a customer in a month
            allocated = False
            for player_row in player_list:
                if player_row['Total Fee'] == amount and not player_row['Allocated']:
                    player_row['Invoice'].ApplyPayment(trans, stripe_acct, GncNumeric(amount), \
                            GncNumeric(1), txn_date, None, None)
                    RecordStripeFee(book, trans, net, fee)
                    player_row['Allocated'] = True
                    allocated = True
                    break
            if not allocated:
                print("Unable to allocate payment to invoice :" + player_row['Total Fee'] \
                        + player_name + '\n')
        
        elif txn_type == 'payout':
            # process a transfer from stripe account to checking account
            DoStripeTransfer(book, amount, txn_date)

        elif description.startswith(REFUND_PREFIX):
            # TODO process the refund.
            # Create a credit note (invoice with negative amount?)
            # Apply a payment
            # create the BAC invoice
            # is there a way of knowing whether it was court hire or light hire ?
            refund_item = re.sub(REFUND_PREFIX + ' \((.*)\)', '\\1', description)
            refund_detail = row["Session (metadata)"]
            if refund_item == COURT_BOOKING:
                # Coburg Tennis Club Tuesday, 19 October 2021 5:00 PM
                court_fee = None
                light_fee = None
                booking_datetime = None
                try:
                    booking_datetime = datetime.strptime(refund_detail, \
                            "Coburg Tennis Club %A, %d %B %Y %I:%M %p")
                except ValueError as e:
                    pass

                if booking_datetime:
                    for key in cs_map:
                        l = cs_map[key]
                        for cs_row in l:
                            # 17/10/2021 10:30:00 AM
                            cs_booking_datetime = datetime.strptime(\
                                    cs_row["Booking Date"] + " " + cs_row["Booking Time"], \
                                    "%Y-%m-%d %H:%M:%S")
                            cs_total_fee = cs_row["Total Fee"]
                            if cs_booking_datetime == booking_datetime and cs_total_fee == amount:
                                court_fee = cs_row['Court Fee']
                                light_fee = cs_row['Light Fee']

                if court_fee == None and light_fee == None:
                    court_fee = amount
                # create the credit note, and apply the payment
                book_a_court = GetBookACourtCustomer(book)
                invoice_id = book.InvoiceNextID(book_a_court)
                invoice = Invoice(book, invoice_id, AUD, book_a_court, txn_date)
                invoice.SetIsCreditNote(True)
                court_hire_acct = GetCourtHireAcct(book)
                light_hire_acct = GetLightHireAcct(book)
                receivables = GetReceivablesAcct(book)
                if court_fee:
                    invoice_entry = Entry(book, invoice, txn_date)
                    invoice_entry.SetDescription("Court Hire Refund")
                    invoice_entry.SetQuantity(GncNumeric(1))
                    invoice_entry.SetInvPrice(GncNumeric(court_fee))
                    invoice_entry.SetInvAccount(court_hire_acct)
                if light_fee:
                    # rest is light fee
                    lf_invoice_entry = Entry(book, invoice, txn_date)
                    lf_invoice_entry.SetDescription("Light Hire Refund")
                    lf_invoice_entry.SetQuantity(GncNumeric(1))
                    lf_invoice_entry.SetInvPrice(GncNumeric(light_fee))
                    lf_invoice_entry.SetInvAccount(light_hire_acct)

                invoice.PostToAccount(receivables, txn_date, txn_date,"", True, False)

                # Apply refund
                trans = Transaction(book)
                trans.BeginEdit()
                split1 = Split(book)
                split1.SetAccount(stripe_acct)
                split1.SetValue(GncNumeric(amount))
                split1.SetParent(trans)

                trans.SetCurrency(AUD)
                trans.SetDate(txn_date.day, txn_date.month, txn_date.year)
                trans.SetDescription("Book A Court Refund")
                trans.CommitEdit()

                # now associate the payment with the correct invoice
                # TODO make it deal properly with multiple bookings by a customer in a month
                invoice.ApplyPayment(trans, stripe_acct, GncNumeric(amount), \
#                invoice.ApplyPayment(None, stripe_acct, GncNumeric(amount), \
                                GncNumeric(1), txn_date, None, None)
                RecordStripeFee(book, trans, net, fee)
            else:
                # We don't know really who it's for, so let's just do a refund payment
                if re.search(MEMBERSHIP_PREFIX, refund_item):
                    refund_acct = GetMembershipAcct(book, refund_item)
#                elif re.search(EVENT_PREFIX, refund_item):
                else:
                    refund_acct = GetEventAcct(book, refund_item)
                # Apply refund
                trans = Transaction(book)
                trans.BeginEdit()
                split1 = Split(book)
                split1.SetAccount(stripe_acct)
                split1.SetValue(GncNumeric(amount))
                split1.SetParent(trans)
                split3 = Split(book)
                split3.SetAccount(refund_acct)
                split3.SetValue(GncNumeric(amount).neg())
                split3.SetParent(trans)

                trans.SetCurrency(AUD)
                trans.SetDate(txn_date.day, txn_date.month, txn_date.year)
                trans.SetDescription("Refund")
                trans.CommitEdit()

        elif description.startswith(MEMBERSHIP_PREFIX):
            # process the membership payment for the individual customer
            # find the customer (using the query methods - see qof.py)
            # create an invoice for the customer
            # apply the payment for the invoice
            # 2021 Adult Half Year Membership, ID: f90ae24f-21f5-46ed-85b6-bee2d22e271d, Customer: Greg Armstrong, googarmstrong@gmail.com, Customer ID:f07b02bb-5dae-4a0d-896d-5d5a6bb0cf8c
            membership_item = re.sub(MEMBERSHIP_PREFIX + ':(.*)', '\\1', description)
            membership_detail = row["Membership (metadata)"]
            customer_name = re.sub('.*Customer: ([^,]*), *(.*)', '\\1', membership_detail)
            customer_email = re.sub('.*Customer: ([^,]*), *(.*),.*', '\\2', membership_detail)
            customer_id = re.sub('.*Customer ID:(.*).*', '\\1', membership_detail)
            customer = GetCustomer(book, customer_name, customer_email)

            # Now add an invoice and payment for that customer
            membership_acct = GetMembershipAcct(book, membership_detail)
            CreateInvoiceAndPayment(book, customer, membership_acct, amount, fee, net, membership_item, txn_date)

        elif description.startswith(EVENT_PREFIX):
            # process the event payment for the individual customer
            # find the customer (using the query methods - see qof.py)
            # create an invoice for the customer
            # apply the payment for the invoice
            event_detail = row['Category (metadata)']
            if event_detail == 'Custom':
                event_detail = row['Course Name (metadata)']
            customer_name = player_name
            customer_email = email_address
            customer = GetCustomer(book, customer_name, customer_email)
            event_acct = GetEventAcct(book, event_detail)
            CreateInvoiceAndPayment(book, customer, event_acct, amount, fee, net, event_detail, txn_date)


def GetCheckingAcct(book):
    root = book.get_root_account()
    assets = root.lookup_by_name("Assets")
    current_assets = assets.lookup_by_name("Current Assets")
    checking_acct = current_assets.lookup_by_name("Checking Account")
    return checking_acct

def GetReceivablesAcct(book):
    root = book.get_root_account()
    assets = root.lookup_by_name("Assets")
    receivables = assets.lookup_by_name("Accounts Receivable")
    return receivables

def GetStripeAcct(book):
    root = book.get_root_account()
    assets = root.lookup_by_name("Assets")
    current_assets = assets.lookup_by_name("Current Assets")
    stripe_acct = current_assets.lookup_by_name("Stripe Account")
    return stripe_acct

def GetCourtHireAcct(book):
    root = book.get_root_account()
    income = root.lookup_by_name("Income")
    casual_hire_acct = income.lookup_by_name("Casual Hire")
    court_hire_acct = casual_hire_acct.lookup_by_name("Court Hire")
    return court_hire_acct

def GetLightHireAcct(book):
    root = book.get_root_account()
    income = root.lookup_by_name("Income")
    casual_hire_acct = income.lookup_by_name("Casual Hire")
    light_hire_acct = casual_hire_acct.lookup_by_name("Hire Light Fees")
    return light_hire_acct

def GetMembershipAcct(book, membership_detail):
    root = book.get_root_account()
    income = root.lookup_by_name("Income")
    membership_acct = income.lookup_by_name("Memberships")
    sub_acct = None
    if re.search('Junior', membership_detail):
        sub_acct = membership_acct.lookup_by_name('Junior')
    elif re.search('Social', membership_detail):
        sub_acct = membership_acct.lookup_by_name('Social')
    elif re.search('Student', membership_detail):
        sub_acct = membership_acct.lookup_by_name('Student')
    elif re.search('Senior', membership_detail):
        sub_acct = membership_acct.lookup_by_name('Senior')
    else: # if re.search('Adult', membership_detail):
        sub_acct = membership_acct.lookup_by_name("Individual")

    return sub_acct

def GetEventAcct(book, event_detail):
    root = book.get_root_account()
    income = root.lookup_by_name("Income")
    event_acct = income.lookup_by_name("Events")
    fundraising_acct = income.lookup_by_name("Fundraising")
    if re.search('OpenCourtSessions', event_detail, re.IGNORECASE):
        sub_acct = event_acct.lookup_by_name('Open Court Sessions')
    elif re.search('MensSocial', event_detail, re.IGNORECASE):
        sub_acct = event_acct.lookup_by_name('MensSocial')
    elif re.search('Girl', event_detail, re.IGNORECASE):
        sub_acct = event_acct.lookup_by_name('Girl Lets Play')
    elif re.search('Club Champ', event_detail, re.IGNORECASE):
        sub_acct = event_acct.lookup_by_name('Club Championships')
    elif re.search('Bingo', event_detail, re.IGNORECASE):
        sub_acct = fundraising_acct.lookup_by_name('Bingo')
    else:
        sub_acct = fundraising_acct.lookup_by_name('Club Events')

    return sub_acct

def GetStripeFeeAcct(book):
    root = book.get_root_account()
    expenses = root.lookup_by_name("Expenses")
    stripe_fee_acct = expenses.lookup_by_name("Stripe Fee")
    return stripe_fee_acct

def GetAUDCurrency(book):
    commod_table = book.get_table()
    AUD = commod_table.lookup('CURRENCY', 'AUD')
    return AUD

def GetBookACourtCustomer(book):
    book_a_court = book.CustomerLookupByID("000001") # Book A Court
    return book_a_court

if __name__ == "__main__":
    filename = 'ctctest.gnucash' if TESTING else 'ctcacounts.gnucash'

    arg_parser = argparse.ArgumentParser(description='Load the Financial Data from Stripe and Clubspark')
    arg_parser.add_argument('--stripefile')
    arg_parser.add_argument('--clubsparkfile')
    arg_parser.add_argument('--gnucashfile', default=filename)

    args = arg_parser.parse_args()

    stripefile = args.stripefile
    clubsparkfile = args.clubsparkfile
    gnucashfile = args.gnucashfile

    s = Session(gnucashfile, SessionOpenMode.SESSION_BREAK_LOCK)
    book = s.book

    # Create the invoices from the Book A Court Clubspark file
    # This allows us to split between Court Hire and Light Hire
    cs_csv = csv.DictReader(open(clubsparkfile, 'r'), skipinitialspace=True)
    cs_map = EnterBacInvoices(cs_csv, book)

    # Now go through the stripe payments report.  
    # This contains all payments, and includes Memberships and Events.
    stripe_csv = csv.DictReader(open(stripefile, 'r'), skipinitialspace=True)
    ProcessStripePayments(stripe_csv, book, cs_map)

    s.save()
    s.end()
