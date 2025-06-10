import _thread
import traceback

from xero_python.accounting import AccountingApi, Contact, Contacts, LineItem, Invoice, LineAmountTypes, Phone, Address
from xero_python.api_client import ApiClient
from xero_python.api_client.configuration import Configuration
from xero_python.api_client.oauth2 import OAuth2Token
from xero_python.identity import IdentityApi
from flask import Flask, url_for, redirect
from flask_oauthlib.contrib.client import OAuth, OAuth2Application
from conf.conf_bridge import CONFIGURATION as conf

from pathlib import Path
from datetime import datetime
import webbrowser
from functools import wraps
import os
import base64
import requests


# configure main flask application
flask_app = Flask(__name__)
flask_app.config["SECRET_KEY"] = os.urandom(16)
flask_app.config["SESSION_TYPE"] = "filesystem"
flask_app.config["ENV"] = "development"
flask_app.config["CLIENT_ID"] = conf["xero_client_id"]
flask_app.config["CLIENT_SECRET"] = conf["xero_client_secret"]
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"


token_list = {}
scope = [
    "offline_access",
    "openid",
    "profile",
    "email",
    "accounting.transactions",
    "accounting.transactions.read",
    "accounting.reports.read",
    "accounting.journals.read",
    "accounting.settings",
    "accounting.settings.read",
    "accounting.contacts",
    "accounting.contacts.read",
    "accounting.attachments",
    "accounting.attachments.read",
    "assets",
    "projects",
    "files",
    "payroll.employees",
    "payroll.payruns",
    "payroll.payslip",
    "payroll.timesheets",
    "payroll.settings"
]
# configure flask-oauthlib application
oauth = OAuth(flask_app)
xero = oauth.remote_app(
    name="xero",
    version="2",
    client_id=flask_app.config["CLIENT_ID"],
    client_secret=flask_app.config["CLIENT_SECRET"],
    endpoint_url="https://api.xero.com/",
    authorization_url="https://login.xero.com/identity/connect/authorize",
    access_token_url="https://identity.xero.com/connect/token",
    refresh_token_url="https://identity.xero.com/conneh_ct/token",
    scope=" ".join(scope)

)  # type: OAuth2Application


# configure xero-python sdk client
api_client = ApiClient(
    Configuration(
        debug=flask_app.config["DEBUG"],
        oauth2_token=OAuth2Token(
            client_id=flask_app.config["CLIENT_ID"], client_secret=flask_app.config["CLIENT_SECRET"]
        ),
    ),
    pool_threads=1,
)
accounting_api = AccountingApi(api_client)

project_type_account_code_map={
    "Restaurant": "41000",
    "Office": "42000",
    "Commercial": "43000",
    "Group House": "44000",
    "Apartment": "45000",
    "Mixed-use Complex": "46000",
    "School": "47000",
    "Others": "48000"
}

bill_type_account_code_map={
    "Mechanical": "51210",
    "Electrical": "51220",
    "Hydraulic": "51230",
    "Fire": "51240",
    "Drafting": "51250",
    "CFD": "51260",
    "Installation": "51270",
    "Others": "51280"
}

xero_invoice_status_map = {
    "DRAFT": "Sent",
    "SUBMITTED": "Sent",
    "AUTHORISED": "Sent",
    "PAID": "Paid",
    "VOIDED": "Voided"
}

xero_bill_status_map = {
    "DRAFT": "Draft",
    "SUBMITTED": "Awaiting Approval",
    "AUTHORISED": "Awaiting Payment",
    "PAID": "Paid",
    "VOIDED": "Voided"
}



def login_xero():
    start_flask()
    webbrowser.open_new_tab("http://localhost:1234/login")



def login_xero2():
    def thread_task():
        start_flask()
    _thread.start_new_thread(thread_task, ())
    webbrowser.open_new_tab("http://localhost:1234/login")


def start_flask():
    flask_app.run(host='localhost', port=1234)
@flask_app.route("/login")
def login():
    redirect_url = url_for("oauth_callback", _external=True)
    response = xero.authorize(callback_uri=redirect_url)
    return response
@flask_app.route("/callback")
def oauth_callback():
    response = xero.authorized_response()
    if response is None or response.get("access_token") is None:
        return f"Access denied: response={str(response)}"
    open(conf["xero_access_token_dir"], 'w').write(response["access_token"])
    open(conf["xero_refresh_token_dir"], 'w').write(response["refresh_token"])
    store_xero_oauth2_token(response)
    return "You are Successfully login, you can go back to the app right now"
@xero.tokengetter
@api_client.oauth2_token_getter
def obtain_xero_oauth2_token():
    return token_list.get("token")

@xero.tokensaver
@api_client.oauth2_token_saver
def store_xero_oauth2_token(token):
    token_list["token"] = token

def xero_token_required(function):
    @wraps(function)
    def decorator(*args, **kwargs):
        xero_token = obtain_xero_oauth2_token()
        if not xero_token:
            return redirect(url_for("login", _external=True))
        return function(*args, **kwargs)
    return decorator
def refresh_token():
    refresh_url = "https://identity.xero.com/connect/token"
    old_refresh_token = open(conf["xero_refresh_token_dir"], 'r').read()
    tokenb4 = f"{conf['xero_client_id']}:{conf['xero_client_secret']}"
    basic_token = base64.urlsafe_b64encode(tokenb4.encode()).decode()
    headers = {
      'Authorization': f"Basic {basic_token}",
      'Content-Type': 'application/x-www-form-urlencoded',
    }
    data = {
        'grant_type': 'refresh_token',
        'refresh_token': old_refresh_token
    }
    response = requests.post(refresh_url, headers=headers, data=data)
    results = response.json()
    open(conf["xero_access_token_dir"], 'w').write(results["access_token"])
    open(conf["xero_refresh_token_dir"], 'w').write(results["refresh_token"])
    store_xero_oauth2_token(
        {
            "access_token": results["access_token"],
            "refresh_token": results["refresh_token"],
            "expires_in": 1800,
            "token_type": "Bearer",
            "scope": scope
        }
    )
@xero_token_required
def get_xero_tenant_id():
    identity_api = IdentityApi(api_client)
    for connection in identity_api.get_connections():
        if connection.tenant_name == conf["xero_tenant_name"]:
            return connection.tenant_id
    raise ValueError("No tenant found for {}".format(conf["xero_tenant_name"]))

@xero_token_required
def get_xero_contacts():
    xero_tenant_id = get_xero_tenant_id()
    where = 'ContactStatus=="ACTIVE"'
    include_archived = 'False'
    contacts = accounting_api.get_contacts(xero_tenant_id, where=where, include_archived=include_archived).to_dict()["contacts"]
    # where = 'ContactStatus=="ARCHIVED"'
    # contacts = accounting_api.get_contacts(xero_tenant_id, include_archived="True", where=where).to_dict()["contacts"]
    result = {}
    for contact in contacts:
        if not contact["account_number"] is None:
            result[contact["account_number"]] = contact
    return contacts

@xero_token_required
def get_xero_contact(contact_id):
    xero_tenant_id = get_xero_tenant_id()
    contact = accounting_api.get_contact(xero_tenant_id, contact_id).to_dict()["contacts"][0]
    return contact

def format_xero_contact_name(full_name, company_name):
    name = []
    if not company_name is None:
        name.append(company_name)
    if not full_name is None:
        name.append(full_name)
    return "-".join(name)
@xero_token_required
def create_xero_contact(contact_name, account_number, full_name, email, phone_number, abn, address):
    xero_tenant_id = get_xero_tenant_id()
    try:
        if not phone_number is None:
            phone_number = [
                Phone(
                    phone_number=phone_number,
                    phone_type="MOBILE"
                )
            ]
        if not address is None:
            address = [
                Address(
                    address_type="POBOX",
                    address_line1=address
                )
            ]
        new_contact = Contact(
            name=contact_name,
            account_number=account_number,
            first_name=full_name,
            email_address=email,
            phones=phone_number,
            tax_number=abn,
            addresses=address
        )
        contacts = Contacts(contacts=[new_contact])
        contact_response = accounting_api.create_contacts(xero_tenant_id, contacts)
        return contact_response.to_dict()["contacts"][0]["contact_id"],''
    except Exception as e:
        return None,str(e)
# def create_xero_contact(contact_name, account_number, full_name, email, phone_number, abn, address):
#     if not phone_number is None:
#         phone_number = [
#             Phone(
#                 phone_number=phone_number,
#                 phone_type="MOBILE"
#             )
#         ]
#     if not address is None:
#         address = [
#             Address(
#                 address_type="POBOX",
#                 address_line1=address
#             )
#         ]
#     new_contact = Contact(
#         name=contact_name,
#         account_number=account_number,
#         first_name=full_name,
#         email_address=email,
#         phones=phone_number,
#         tax_number=abn,
#         addresses=address
#     )
#     contacts = Contacts(contacts=[new_contact])
#     contact_response = accounting_api.create_contacts(xero_tenant_id, contacts)
#     return contact_response.to_dict()["contacts"][0]["contact_id"]

@xero_token_required
def create_or_update_xero_contact(contact_name, account_number, full_name, email, phone_number, abn, address):
    xero_tenant_id = get_xero_tenant_id()
    if not phone_number is None:
        phone_number = [
            Phone(
                phone_number=phone_number,
                phone_type="MOBILE"
            )
        ]
    if not address is None:
        address = [
            Address(
                address_type="POBOX",
                address_line1=address
            )
        ]
    new_contact = Contact(
        name=contact_name,
        account_number=account_number,
        first_name=full_name,
        email_address=email,
        phones=phone_number,
        tax_number=abn,
        addresses=address
    )
    contacts = Contacts(contacts=[new_contact])
    contact_response = accounting_api.update_or_create_contacts(xero_tenant_id, contacts)
    return contact_response.to_dict()["contacts"][0]["contact_id"]


@xero_token_required
def update_xero_contact(contact_id, contact_name, account_number, first_name, last_name, email, phone_number, abn, address):
    xero_tenant_id = get_xero_tenant_id()
    if not phone_number is None:
        phone_number = [
            Phone(
                phone_number=phone_number,
                phone_type="MOBILE"
            )
        ]
    if not address is None:
        address = [
            Address(
                address_type="POBOX",
                address_line1=address
            )
        ]
    contact = Contact(
        name=contact_name,
        account_number=account_number,
        first_name=first_name,
        last_name=last_name,
        email_address=email,
        phones=phone_number,
        tax_number=abn,
        addresses=address
    )
    accounting_api.update_contact(xero_tenant_id, contact_id, contact)

def delete_xero_contact(contact_id):
    xero_tenant_id = get_xero_tenant_id()
    contact = Contact(
        account_number="",
        contact_status="ARCHIVED"
    )
    accounting_api.update_contact(xero_tenant_id, contact_id, contact)

@xero_token_required
def get_xero_invoices():
    xero_tenant_id = get_xero_tenant_id()
    where = 'Type=="ACCREC" and Status!="DELETED"'
    invoices = accounting_api.get_invoices(xero_tenant_id, where=where).to_dict()["invoices"]
    result = {}
    for invoice in invoices:
        result[invoice["invoice_id"]] = invoice
    return result

@xero_token_required
def get_xero_invoice(invoice_id):
    xero_tenant_id = get_xero_tenant_id()
    invoice = accounting_api.get_invoice(xero_tenant_id, invoice_id).to_dict()["invoices"][0]
    return invoice

@xero_token_required
def get_xero_invoice_status(invoice_id):
    invoice = get_xero_invoice(invoice_id)
    return {
        "status": xero_invoice_status_map[invoice["status"]],
        # "fee_amount": invoice["sub_total"],
        "contact": invoice["contact"]["contact_id"],
        "last_payment_date": None if invoice["payments"] is None else invoice["payments"][0]["date"],
        "payment_amount": invoice["amount_paid"]
    }

@xero_token_required
def get_xero_invoice_status_by_invoice_number(invoice_number):
    invoice = get_xero_invoice(invoice_number)
    return {
        "invoice_id": invoice["invoice_id"],
        "status": invoice["status"],
        "contact": invoice["contact"]["account_number"],
        "fee_amount": invoice["sub_total"],
        "last_payment_date": None if invoice["payments"] is None else invoice["payments"][0]["date"],
        "payment_amount": invoice["amount_paid"]
    }

@xero_token_required
def create_xero_invoice(invoice_number, items, project_type, project_name, contact_id):
    xero_tenant_id = get_xero_tenant_id()
    line_item_list = []
    for item in items:
        line_item_list.append(
            LineItem(
                description=item["Item"],
                quantity=1,
                unit_amount=item["Fee"],
                tax_type="OUTPUT",
                account_code=project_type_account_code_map[project_type]
            )
        )

    invoice = Invoice(
            type="ACCREC",
            contact=Contact(contact_id),
            date=datetime.today(),
            due_date=datetime.today(),
            line_items=line_item_list,
            invoice_number=invoice_number,
            reference=project_name,
            status="DRAFT"
    )
    api_response = accounting_api.create_invoices(xero_tenant_id, invoice)
    return api_response.to_dict()["invoices"][0]["invoice_id"]

@xero_token_required
def update_xero_invoice(invoice_id, invoice_number, items, project_type, project_name, contact_id):
    xero_tenant_id = get_xero_tenant_id()
    line_item_list = []
    for item in items:
        line_item_list.append(
            LineItem(
                description=item["Item"],
                quantity=1,
                unit_amount=float(item["Fee"]),
                tax_type="OUTPUT",
                account_code=project_type_account_code_map[project_type]
            )
        )

    invoice = Invoice(
            type="ACCREC",
            contact=Contact(contact_id=contact_id),
            date=datetime.today(),
            due_date=datetime.today(),
            line_items=line_item_list,
            invoice_number=invoice_number,
            reference=project_name
    )
    accounting_api.update_invoice(xero_tenant_id, invoice_id, invoice)


@xero_token_required
def update_xero_invoice_contact(invoice_id, contact_id):
    xero_tenant_id = get_xero_tenant_id()
    invoice = Invoice(
            contact=Contact(contact_id=contact_id)
    )
    accounting_api.update_invoice(xero_tenant_id, invoice_id, invoice)

@xero_token_required
def update_xero_bill_contact(invoice_id, contact_id):
    xero_tenant_id = get_xero_tenant_id()
    invoice = Invoice(
            contact=Contact(contact_id=contact_id)
    )
    accounting_api.update_invoice(xero_tenant_id, invoice_id, invoice)

def update_xero_invoice_status(invoice_id, status):
    xero_tenant_id = get_xero_tenant_id()
    invoice = Invoice(
        status=status
    )
    accounting_api.update_invoice(xero_tenant_id, invoice_id, invoice)

@xero_token_required
def get_xero_bills():
    xero_tenant_id = get_xero_tenant_id()
    where = 'Type=="ACCPAY" and Status!="DELETED"'
    invoices = accounting_api.get_invoices(xero_tenant_id, where=where).to_dict()["invoices"]
    result = {}
    for invoice in invoices:
        result[invoice["invoice_id"]] = invoice
    return result

@xero_token_required
def get_xero_bill(bill_id):
    xero_tenant_id = get_xero_tenant_id()
    bill = accounting_api.get_invoice(xero_tenant_id, bill_id).to_dict()["invoices"][0]
    return bill
@xero_token_required
def get_xero_bill_status(bill_id):
    bill = get_xero_bill(bill_id)
    return {
        "contact": bill["contact"]["contact_id"],
        "status": xero_bill_status_map[bill["status"]],
        "paid_data": bill["fully_paid_on_date"]
    }

@xero_token_required
def create_xero_bill(service, amount, type, no_gst, contact_id, file_dir):
    xero_tenant_id = get_xero_tenant_id()
    path = Path(file_dir)
    reference = path.stem
    file_name = path.name
    line_item_list = [
        LineItem(
            description=service,
            quantity=1,
            unit_amount=amount,
            account_code=bill_type_account_code_map[type],
            tax_type="OUTPUT" if no_gst else "INPUT"
        )
    ]
    bill = Invoice(
            type="ACCPAY",
            date=datetime.today(),
            due_date=datetime.today(),
            contact=Contact(contact_id=contact_id),
            line_items=line_item_list,
            line_amount_types=LineAmountTypes.NOTAX if no_gst else LineAmountTypes.EXCLUSIVE,
            invoice_number=reference,
            reference=reference,
            status="DRAFT"
        )
    api_response = accounting_api.create_invoices(xero_tenant_id, bill)
    bill_id = api_response.to_dict()["invoices"][0]["invoice_id"]
    open_file = open(file_dir, 'rb')
    body = open_file.read()
    accounting_api.create_invoice_attachment_by_file_name(xero_tenant_id, bill_id, file_name, body)
    return bill_id

@xero_token_required
def update_xero_bill(bill_id, service, amount, type, no_gst, contact_id):
    xero_tenant_id = get_xero_tenant_id()
    line_item_list = [
        LineItem(
            description=service,
            quantity=1,
            unit_amount=amount,
            account_code=bill_type_account_code_map[type],
            tax_type="OUTPUT" if no_gst else "INPUT"
        )
    ]
    bill = Invoice(
            type="ACCPAY",
            contact=Contact(contact_id),
            line_items=line_item_list,
            line_amount_types=LineAmountTypes.NOTAX if no_gst else LineAmountTypes.EXCLUSIVE,
    )
    accounting_api.update_invoice(xero_tenant_id, bill_id, bill)
@xero_token_required
def delete_xero_bill(bill_id):
    xero_tenant_id = get_xero_tenant_id()
    bill = Invoice(
        type="ACCPAY",
        invoice_id=bill_id,
        status="DELETED"
    )
    accounting_api.update_invoice(xero_tenant_id, bill)

# if __name__ == '__main__':
#     login_xero()

    # refresh_token()
#     update_xero_invoice_contact("b0803945-e89e-4a66-b749-4743368633f1", "a0c2c18a-b560-483d-979a-958accd52d39")