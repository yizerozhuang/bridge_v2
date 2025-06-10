import traceback
import psutil
from win32com import client as win32client
import pythoncom
import os
from datetime import date, datetime


def convert_time_format(date_str):
    try:
        date_obj = datetime.strptime(date_str, "%d-%b-%Y")
    except:
        date_obj = datetime.strptime(date_str, "%d-%m-%Y")
    day = date_obj.day
    month = date_obj.strftime("%B")
    year = date_obj.year
    if day == "01":
        day_suffix = "st"
    elif day == "02":
        day_suffix = "nd"
    else:
        day_suffix = "th"
    formatted_date = f"{day}<sup>{day_suffix}</sup> {month} {year}"
    return formatted_date


def f_email_fee(subject,client_email,main_contact_email,email2cc,message2email,pdf_dir):
    try:
        if not "OUTLOOK.EXE" in (p.name() for p in psutil.process_iter()):
            os.startfile("outlook")
        ol = win32client.Dispatch("Outlook.Application", pythoncom.CoInitialize())
        olmailitem = 0x0
        newmail = ol.CreateItem(olmailitem)
        newmail.Subject =subject
        newmail.To = f'{client_email}; {main_contact_email}'
        if email2cc!='':
            newmail.CC = email2cc
        newmail.BCC = "bridge@pcen.com.au"
        newmail.GetInspector()
        index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))
        newmail.HTMLbody = newmail.HTMLbody[:index + 1] + message2email + newmail.HTMLbody[index + 1:]
        newmail.Attachments.Add(pdf_dir)
        newmail.Display()
    except:
        traceback.print_exc()



def f_email_invoice(subject, client_email, main_contact_email, message2email, pdf_dir):
    try:
        if not "OUTLOOK.EXE" in (p.name() for p in psutil.process_iter()):
            os.startfile("outlook")
        ol = win32client.Dispatch("Outlook.Application", pythoncom.CoInitialize())
        olmailitem = 0x0
        newmail = ol.CreateItem(olmailitem)
        newmail.Subject =subject
        newmail.To = f'{client_email}; {main_contact_email}'
        newmail.BCC ="bridge@pcen.com.au"
        newmail.GetInspector()
        index = newmail.HTMLbody.find(">", newmail.HTMLbody.find("<body"))
        newmail.HTMLbody = newmail.HTMLbody[:index + 1] + message2email + newmail.HTMLbody[index + 1:]
        newmail.Display()
        newmail.Attachments.Add(pdf_dir)
    except:
        traceback.print_exc()


