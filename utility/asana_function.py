import asana
from conf.conf_bridge import CONFIGURATION as conf
from utility.sql_function import get_value_from_table, get_value_from_table_with_filter, format_output
from datetime import datetime
import pytz

asana_configuration = asana.Configuration()
asana_configuration.access_token = '2/1203283895754383/1206354773081941:c116d68430be7b2832bf5d7ea2a0a415'
asana_api_client = asana.ApiClient(asana_configuration)

project_api_instance = asana.ProjectsApi(asana_api_client)
task_api_instance = asana.TasksApi(asana_api_client)
custom_fields_api_instance = asana.CustomFieldsApi(asana_api_client)
custom_fields_setting_api_instance = asana.CustomFieldSettingsApi(asana_api_client)
template_api_instance = asana.TaskTemplatesApi(asana_api_client)
time_tracking_entries_api_instance = asana.TimeTrackingEntriesApi(asana_api_client)
stories_api_instance = asana.StoriesApi(asana_api_client)
workspace_gid = '1198726743417674'
# test workplace_id
# workspace_gid = "1208546788985072"

mp_project_name = "MP"
invoice_project_name = "Invoice status"
bill_project_name = "Bill status"

user_json = format_output(get_value_from_table("users"))

def clean_response(response):
    response = response.to_dict()["data"]
    return response

def name_id_map(api_list):
    res = dict()
    for item in api_list:
        res[item["name"]] = item["gid"]
    return res

def flatter_custom_fields(asana_task):
    if not "custom_fields" in asana_task.keys():
        return asana_task
    for field in asana_task["custom_fields"]:
        assert not field["name"] in asana_task.keys(), f'{field["name"]} already exists'
        if f"{field['resource_subtype']}_value" in field.keys():
            asana_task[field["name"]] = field[f"{field['resource_subtype']}_value"]
        elif f"{field['resource_subtype']}_values" in field.keys():
            asana_task[field["name"]] = field[f"{field['resource_subtype']}_values"]
        else:
            raise ValueError(f"Unexpect field {field['name']} in {asana_task}")
    return asana_task

def format_sub_tasks(all_subtask):
    for subtask in all_subtask:
        subtask["projects_name"] = []
        for subtask_project in subtask["projects"]:
            subtask["projects_name"].append(subtask_project["name"])
    return all_subtask

def format_datetime(datetime_obj):
    return None if datetime_obj is None else {"date": str(datetime_obj)[0:10]}

asana_projects = clean_response(project_api_instance.get_projects_for_workspace(workspace_gid))
asana_projects_id_map = name_id_map(asana_projects)

mp_custom_fields_setting = clean_response(
    custom_fields_setting_api_instance.get_custom_field_settings_for_project(asana_projects_id_map[mp_project_name]))
mp_custom_fields = [custom_field["custom_field"] for custom_field in mp_custom_fields_setting]
mp_custom_field_id_map = name_id_map(mp_custom_fields)

mp_status_field = clean_response(custom_fields_api_instance.get_custom_field(mp_custom_field_id_map["Status"]))
mp_status_id_map = name_id_map(mp_status_field["enum_options"])

mp_service_filed = clean_response(custom_fields_api_instance.get_custom_field(mp_custom_field_id_map["Services"]))
mp_service_id_map = name_id_map(mp_service_filed["enum_options"])

mp_contact_field = clean_response(custom_fields_api_instance.get_custom_field(mp_custom_field_id_map["Contact Type"]))
mp_contact_id_map = name_id_map(mp_contact_field["enum_options"])

invoice_custom_fields_setting = clean_response(
    custom_fields_setting_api_instance.get_custom_field_settings_for_project(asana_projects_id_map[invoice_project_name]))
invoice_custom_fields = [custom_field["custom_field"] for custom_field in invoice_custom_fields_setting]
invoice_custom_field_id_map = name_id_map(invoice_custom_fields)

invoice_status_custom_field = clean_response(
    custom_fields_api_instance.get_custom_field(invoice_custom_field_id_map["Invoice status"]))
invoice_status_id_map = name_id_map(invoice_status_custom_field["enum_options"])

bill_custom_fields_setting = clean_response(
    custom_fields_setting_api_instance.get_custom_field_settings_for_project(asana_projects_id_map[bill_project_name]))
bill_custom_fields = [custom_field["custom_field"] for custom_field in bill_custom_fields_setting]
bill_custom_field_id_map = name_id_map(bill_custom_fields)

bill_status_field = clean_response(custom_fields_api_instance.get_custom_field(bill_custom_field_id_map["Bill status"]))
bill_status_id_map = name_id_map(bill_status_field["enum_options"])

bill_types_field = clean_response(custom_fields_api_instance.get_custom_field(bill_custom_field_id_map["Type"]))
bill_type_id_map = name_id_map(bill_types_field["enum_options"])

def get_asana_task(asana_id):
    return flatter_custom_fields(clean_response(task_api_instance.get_task(asana_id)))

def delete_asana_task(asana_id):
    task_api_instance.delete_task(asana_id)

def get_asana_email(asana_id):
    """Get an asana task's email content

    :param str asana_id: the asana id of the project (required)
    :return: str: the asana task's email content in html format
    """
    return get_asana_task(asana_id)["notes"]

def get_asana_estimated_time(asana_id):
    """Get an asana task's estimated time in minutes

    :param str asana_id: the asana id of the task (required)
    :return: float: the estimated time of the task in minutes
    """
    return get_asana_task(asana_id)["Estimated time"]

def get_asana_actual_time(asana_id):
    """Get an asana task's actual time in minutes

    :param str asana_id: the asana id of the task (required)
    :return: float: the actual time of the task in minutes
    """
    return get_asana_task(asana_id)["actual_time_minutes"]

def update_asana_estimated_time(asana_id, estimated_time):
    """Update an asana task's estimated time

    :param str asana_id: the asana id of the task (required)
    :param float estimated_time: the estimated time of the task in minutes, can be none(required)
    """
    body = {
        "custom_fields": {
            mp_custom_field_id_map["Estimated time"]: estimated_time
        }
    }
    body = asana.TasksTaskGidBody(body)
    task_api_instance.update_task(task_gid=asana_id, body=body)

def update_asana_actual_time(asana_id, actual_time):
    """Update an asana task's actual time

    :param str asana_id: the asana id of the task (required)
    :param float actual_time: the actual time of the task in minutes, can be none(required)
    """
    body = {
        "data": {
            "duration_minutes": actual_time
        }
    }
    asana_response = clean_response(time_tracking_entries_api_instance.get_time_tracking_entries_for_task(asana_id))
    if actual_time == 0:
        for i in range(len(asana_response)):
            time_tracking_id = asana_response[i]["gid"]
            time_tracking_entries_api_instance.delete_time_tracking_entry(time_tracking_id)
    else:
        if len(asana_response) == 0:
            time_tracking_entries_api_instance.create_time_tracking_entry(body, asana_id)
        else:
            actual_time_id = clean_response(time_tracking_entries_api_instance.get_time_tracking_entries_for_task(asana_id))[0]["gid"]
            time_tracking_entries_api_instance.update_time_tracking_entry(body, actual_time_id)
            for i in range(1, len(asana_response)):
                time_tracking_id = asana_response[i]["gid"]
                time_tracking_entries_api_instance.delete_time_tracking_entry(time_tracking_id)

def get_asana_sub_tasks(asana_id):
    """Get all the subtasks of an asana project

    :param str asana_id: the asana id of the task (required)
    :return dict: all the asana subtasks that categorized into invoices bills and subtasks
                  filtering out tasks that starts with --- and separator
    """
    off_set = None
    opt_fields = ["name", "assignee", "due_on", "is_rendered_as_separator", "completed", "projects.name", "custom_fields", "actual_time_minutes"]
    all_subtasks = {"Subtasks":{}, "Invoices":{}, "Bills":{}}

    while True:
        if off_set is None:
            asana_response = task_api_instance.get_subtasks_for_task(asana_id, opt_fields=opt_fields, limit=100)
        else:
            asana_response = task_api_instance.get_subtasks_for_task(asana_id, opt_fields=opt_fields, limit=100,
                                                                     offset=off_set)
        all_asana_subtasks = format_sub_tasks(clean_response(asana_response))
        for sub_task in all_asana_subtasks:
            if not sub_task["is_rendered_as_separator"] and not sub_task["name"].startswith("---"):
                sub_task = flatter_custom_fields(sub_task)
                if invoice_project_name in sub_task["projects_name"]:
                    all_subtasks["Invoices"][sub_task["gid"]] = {
                        "name": sub_task["name"]
                    }
                elif bill_project_name in sub_task["projects_name"]:
                    all_subtasks["Bills"][sub_task["gid"]] = {
                        "name": sub_task["name"]
                    }
                else:
                    if not sub_task["assignee"] is None:
                        if len(get_value_from_table_with_filter("users", "asana_id", sub_task["assignee"]["gid"]))!=0:
                            assignee = get_value_from_table_with_filter("users", "asana_id", sub_task["assignee"]["gid"])[0][0]
                        else:
                            assignee = None
                    else:
                        assignee = None
                    all_subtasks["Subtasks"][sub_task["gid"]] = {
                        "name": sub_task["name"],
                        "due_on": sub_task["due_on"],
                        "completed":sub_task["completed"],
                        "assignee": assignee,
                        "actual_time": sub_task["actual_time_minutes"],
                        "estimated_time": sub_task["Estimated time"]
                    }
        if asana_response.to_dict()["next_page"] is None:
            return all_subtasks
        off_set = asana_response.to_dict()["next_page"]["offset"]

def update_asana_sub_task(asana_id, assignee_email=None, due_on=None, completed=None):
    """Update the asana task status, will not overwrite the existing value

    :param str asana_id: the asana id of the task (required)
    :param str assignee_email: the email address of assignee (optional)
    :param datetime due_at: the due date of the task will only consider the date but time(optional)
    :param bool completed: the asana id of the task (optional)
    """
    sub_task = get_asana_task(asana_id)
    body = {}
    if sub_task["assignee"] is None:
        if not assignee_email is None:
            body["assignee"] = user_json[assignee_email]["supervisor_asana_id"]
    if sub_task["due_on"] is None:
        if not due_on is None:
            body["due_on"] = due_on.strftime("%Y-%m-%d")
    if not completed is None:
        if sub_task["completed"] != completed:
            body["completed"] = completed

    if len(body) !=0:
        body = asana.TasksTaskGidBody(body)
        task_api_instance.update_task(task_gid=asana_id, body=body)

def add_project_to_asana_task(asana_id, project_type):
    """Add the project to an asana task

    :param str asana_id: the asana id of the task (required)
    :param str project_type: the project that need to be inserted (required)
    """

    body = asana.TaskGidAddProjectBody({"project": asana_projects_id_map[project_type]})
    task_api_instance.add_project_for_task(task_gid=asana_id, body=body)

def delete_project_to_asana_task(asana_id, project_type):
    """Delete the project to an asana task

    :param str asana_id: the asana id of the task (required)
    :param str project_type: the project that need to be deleted (required)
    """

    body = asana.TaskGidAddProjectBody({"project": asana_projects_id_map[project_type]})
    task_api_instance.remove_project_for_task(task_gid=asana_id, body=body)

def update_asana_project_tags(asana_id, project_type):
    """Update the asana project tags

    :param str asana_id: the asana id of the task (required)
    :param str project_type: the corresponding project type (required)
    """

    asana_task = get_asana_task(asana_id)
    current_projects_list = [project["name"] for project in asana_task["projects"]]
    project_type_list = conf["project_types"]
    for p_type in project_type_list:
        if p_type == project_type:
            continue
        if p_type in current_projects_list:
            delete_project_to_asana_task(asana_id, p_type)
    if not project_type in current_projects_list:
        add_project_to_asana_task(asana_id, project_type)
def format_asana_contact(name, company):
    """Format the asana contact information

    :param str name: name of the contact
    :param str company: name of the company
    :return str: the formatted contact string
    """

    result = []
    if not company is None:
        result.append(company)
    if not name is None:
        result.append(name)
    return "-".join(result)


def _get_asana_project_service(project):
    asana_service = [service["name"] for service in project["Services"]]
    result_service = []
    for service in conf["services_list"]:
        if service in asana_service:
            result_service.append(service)
    return result_service

def get_asana_project(asana_id):
    """Update the asana project tags

    :param str asana_id: the asana id of the task (required)
    :param str project_type: the corresponding project type (required)
    """

    project = get_asana_task(asana_id)
    return {
        "name": project["name"],
        "status": project["Status"]["name"] if not project["Status"] is None else None,
        "services": _get_asana_project_service(project) if not project["Services"] is None else [],
        "shop_name": project["Shop name"] if not project["Shop name"] is None else "",
        "apt": project["Apt/Room/Area"] if not project["Apt/Room/Area"] is None else "",
        "basement": project["Basement/Car Spots"] if not project["Basement/Car Spots"] is None else "",
        "notes": project["Feature/Notes"] if not project["Feature/Notes"] is None else "",
        "client_name": project["Client"],
        "main_contact_name": project["Main Contact"],
        "main_contact_type": project["Contact Type"]["name"] if not project["Contact Type"] is None else None,
        "total_fee": project["Fee ExGST"],
        "total_paid_fee": project["Total Paid InGST"],
        "overdue_fee": project["Overdue InGST"]
    }

def create_asana_project(project_type):
    current_project_template = clean_response(template_api_instance.get_task_templates(project=asana_projects_id_map[project_type]))

    template_id_map = name_id_map(current_project_template)

    api_respond = clean_response(template_api_instance.instantiate_task(template_id_map["P:\\300000-XXXX"]))
    asana_id = api_respond["new_task"]["gid"]

    add_project_to_asana_task(asana_id, mp_project_name)

    asana_url = get_asana_task(asana_id)['permalink_url']

    return asana_id, asana_url
def update_asana_project(asana_id, name, status, services, shop_name, apt, basement, notes, client_name,
                         main_contact_name, main_contact_type, total_fee, total_paid_fee, overdue_fee):
    if not services is None:
        services = [mp_service_id_map[service] for service in services]
    if not main_contact_type is None:
        main_contact_type = mp_contact_id_map[main_contact_type]
    asana_update_body = asana.TasksTaskGidBody(
        {
            "name": name,
            "custom_fields": {
                mp_custom_field_id_map["Status"]: mp_status_id_map[status],
                mp_custom_field_id_map["Services"]: services,
                mp_custom_field_id_map["Shop name"]: shop_name,
                mp_custom_field_id_map["Apt/Room/Area"]: apt,
                mp_custom_field_id_map["Basement/Car Spots"]: basement,
                mp_custom_field_id_map["Feature/Notes"]: notes,
                mp_custom_field_id_map["Client"]: client_name,
                mp_custom_field_id_map["Main Contact"]: main_contact_name,
                mp_custom_field_id_map["Contact Type"]: main_contact_type,
                mp_custom_field_id_map["Fee ExGST"]: total_fee,
                mp_custom_field_id_map["Total Paid InGST"]: total_paid_fee,
                mp_custom_field_id_map["Overdue InGST"]: overdue_fee,
            }
        }
    )
    task_api_instance.update_task(task_gid=asana_id, body=asana_update_body)
def update_asana_email(asana_id, email):#add email content to asana
    body = asana.TasksTaskGidBody({"notes":email})
    task_api_instance.update_task(task_gid=asana_id, body=body)
def get_asana_invoice(asana_id):
    invoice = get_asana_task(asana_id)
    return {
        "name": invoice["name"],
        "status": invoice["Invoice status"]["name"] if not invoice["Invoice status"] is None else None,
        "payment_date": datetime.strptime(invoice["Payment Date"]["_date"], "%Y-%m-%d").date() if not invoice["Payment Date"] is None else None,
        "payment_ingst": invoice["Payment InGST"],
        "net": invoice["Net"]
    }

def create_asana_invoice(asana_id):
    invoice_status_templates = clean_response(template_api_instance.get_task_templates(project=asana_projects_id_map[invoice_project_name]))
    template_id_map = name_id_map(invoice_status_templates)
    invoice_template_id = template_id_map["INV Template"]
    body = asana.TasksTaskGidBody({"name": f"INV xxxxxx"})
    response = clean_response(template_api_instance.instantiate_task(task_template_gid=invoice_template_id, body=body))
    new_inv_task_gid = response['new_task']["gid"]
    body = asana.TasksTaskGidBody({"parent": asana_id, "insert_before": None})
    task_api_instance.set_parent_for_task(body=body, task_gid=new_inv_task_gid)
    return new_inv_task_gid

def update_asana_invoice(asana_invoice_id, name, status, payment_date, payment_ingst, net):
    if not payment_date is None:
        payment_date = format_datetime(payment_date)
    body = {
        "name": name,
        "custom_fields":{
            invoice_custom_field_id_map["Invoice status"]: invoice_status_id_map[status],
            invoice_custom_field_id_map["Net"]: net,
            invoice_custom_field_id_map["Payment InGST"]: payment_ingst,
            invoice_custom_field_id_map["Payment Date"]: payment_date
        }
    }
    body = asana.TasksTaskGidBody(body)
    task_api_instance.update_task(task_gid=asana_invoice_id, body=body)

def get_asana_bill(asana_id):
    bill = get_asana_task(asana_id)
    return {
        "name": bill["name"],
        "status": bill["Bill status"],
        "contact": bill["From"],
        "bill_in_date": bill["Bill In Date"],
        "paid_date": bill["Paid Date"],
        "type": bill["Type"],
        "amount_exgst": bill["Amount Excl GST"],
        "amount_ingst": bill["Amount Incl GST"],
        "headup": bill["HeadsUp"]
    }

def create_asana_bill(asana_id):
    bill_status_templates = clean_response(template_api_instance.get_task_templates(project=asana_projects_id_map[bill_project_name]))
    template_id_map = name_id_map(bill_status_templates)
    bill_template_id = template_id_map["BIL 4"]
    body = asana.TasksTaskGidBody({"name": f"BIL xxxxxx"})
    response = template_api_instance.instantiate_task(body=body, task_template_gid=bill_template_id).to_dict()
    new_bill_task_gid = response["data"]['new_task']["gid"]
    body = asana.TasksTaskGidBody({"parent": asana_id, "insert_before": None})
    task_api_instance.set_parent_for_task(body=body, task_gid=new_bill_task_gid)
    return new_bill_task_gid


def update_asana_bill(asana_bill_id, name, status, contact, bill_in_date, paid_date,  type, amount_exgst, amount_ingst, headup):
    body = {
        "name": name,
        "custom_fields": {
            bill_custom_field_id_map["Bill status"]: bill_status_id_map[status],
            bill_custom_field_id_map["From"]: contact,
            bill_custom_field_id_map["Bill In Date"]: format_datetime(bill_in_date),
            bill_custom_field_id_map["Paid Date"]: format_datetime(paid_date),
            bill_custom_field_id_map["Type"]: bill_type_id_map[type],
            bill_custom_field_id_map["Amount Excl GST"]: amount_exgst,
            bill_custom_field_id_map["Amount Incl GST"]: amount_ingst,
            bill_custom_field_id_map["HeadsUp"]: headup
        }
    }
    body = asana.TasksTaskGidBody(body)
    task_api_instance.update_task(task_gid=asana_bill_id, body=body)

def get_asana_tasks_by_user_by_date(user_email, due_at):
    user_id = user_json[user_email]["asana_id"]
    due_at = due_at.strftime("%Y-%m-%d")
    opt_fields = ["name", "parent.name", "actual_time_minutes", "custom_fields"]
    result = {}
    # off_set = None
    # while True:
    #     if off_set is None:
    asana_response = task_api_instance.search_tasks_for_workspace(workspace_gid=workspace_gid, assignee_any=user_id,
                                                                    due_on=due_at, opt_fields=opt_fields)
        # else:
        #     asana_response = task_api_instance.search_tasks_for_workspace(workspace_gid=workspace_gid,
        #                                                                   assignee_any=user_id,
        #                                                                   due_on=due_at, opt_fields=opt_fields)
    tasks = clean_response(asana_response)

    for task in tasks:
        try:
            task = flatter_custom_fields(task)
            # print(task)
            result[task["gid"]] = {
                "name": task["name"],
                "parent": task["parent"]["name"],
                "actual_time": task["actual_time_minutes"],
                "estimated_time": task["Estimated time"]
            }
        except:
            pass
    return result


def get_asana_tasks_from_project(project_id):
    return clean_response(task_api_instance.get_tasks_for_project(project_id))

def update_asana_1on1_subtask(project_id, txt):
    tasks = get_asana_tasks_from_project(project_id)
    tasks_id = tasks[-1]["gid"]
    email_content = get_asana_email(tasks_id)
    update_asana_email(tasks_id, email_content+txt)


def convert_sydney_time_to_utc(sydney_time):
    sydney_tz = pytz.timezone('Australia/Sydney')
    sydney_time = sydney_tz.localize(sydney_time)
    utc_time = sydney_time.astimezone(pytz.utc)
    return utc_time

def create_asana_meeting_subtask(task_id, date_time, name, notes, assignee):
    if not assignee is None:
        assignee = user_json[assignee]["asana_id"]
    if not date_time is None:
        date_time = convert_sydney_time_to_utc(date_time)

    body = {
        "data":{
            "name": name,
            "due_at": date_time,
            "notes": notes,
            "assignee":assignee,
            "projects": asana_projects_id_map["Meeting Schedule - All Staff"]
        }
    }
    subtask_id = clean_response(task_api_instance.create_subtask_for_task(body, task_id))["gid"]
    all_subtasks = get_asana_sub_tasks(task_id)["Subtasks"]
    last_subtask = list(all_subtasks.keys())[-1]
    task_api_instance.set_parent_for_task({"data":{"parent":task_id, "insert_after":last_subtask}}, subtask_id)

def create_asana_task_comment(task_id, comment):
    body = {
        "data":{
            "text":comment
        }
    }
    stories_api_instance.create_story_for_task(body, task_id)

# if __name__ == '__main__':
#     create_asana_task_comment("1209267961076047", "thx123")
# date_str = "16 December 2024 3:00 PM"
# date_obj = datetime.strptime(date_str, "%d %B %Y %I:%M %p")
# create_asana_meeting_subtask("1200002694172636", date_obj,"test meeting","test notes", "yitong@forgebc.com.au")