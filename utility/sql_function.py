import mysql.connector

bridge_db = mysql.connector.connect(
    host="192.168.1.199",
    user="bridge",
    password="PcE$yD2023",
    database="bridge"
)

cursor = bridge_db.cursor()
cursor.execute("USE bridge")

def get_all_the_table_names():
    cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_SCHEMA='bridge'")
    return cursor.fetchall()

def get_cursor_description():
    #should use this after get values
    return [description[0] for description in cursor.description]

def get_number_of_columns_of_table(table_name):
    cursor.execute(f"SELECT count(*) FROM information_schema.columns WHERE table_name='{table_name}'")
    return cursor.fetchone()[0]
def get_value_from_table(table_name):
    command = f"SELECT * FROM {table_name}"
    cursor.execute(command)
    return cursor.fetchall()
def get_value_from_table_with_filter(table_name, filter, value):
    command = f'SELECT * FROM {table_name} WHERE {filter}="{value}"'
    cursor.execute(command)
    return cursor.fetchall()

def get_value_from_table_with_filters(table_name, filter1, value1, filter2, value2):
    command = f'SELECT * FROM {table_name} WHERE ({filter1}="{value1}" AND {filter2}="{value2}")'
    cursor.execute(command)
    return cursor.fetchall()

def insert_value_into_table(table_name, value):
    number_of_columns = get_number_of_columns_of_table(table_name)
    command = f"INSERT INTO {table_name} VALUES ({', '.join(['%s' for _ in range(number_of_columns)])})"
    cursor.execute(command, value)
    bridge_db.commit()
def replace_value_in_table(table_name, value):
    number_of_columns = get_number_of_columns_of_table(table_name)
    command = f"REPLACE INTO {table_name} VALUES ({', '.join(['%s' for _ in range(number_of_columns)])})"
    cursor.execute(command, value)
    bridge_db.commit()
def update_value_in_table(table_name, dic, filter, value):
    update = ", ".join([f"{key}=Null" if value is None else f"{key}='{value}'" for key, value in dic.items()])
    command = f"UPDATE {table_name} SET {update} WHERE {filter}='{value}'"
    cursor.execute(command)
    bridge_db.commit()

def delete_value_in_table(table_name, filter, value):
    command = f"DELETE FROM {table_name} WHERE {filter} = '{value}'"
    cursor.execute(command)
    bridge_db.commit()

def check_if_column_in_tabel(table, column):
    command = f"SHOW COLUMNS FROM {table} LIKE '{column}'"
    cursor.execute(command)
    return not cursor.fetchone() is None

def delete_project(quotation_number):
    cursor.execute("SET FOREIGN_KEY_CHECKS=0")
    tables = get_all_the_table_names()
    for table in tables:
        table = table[0]
        if check_if_column_in_tabel(table, "quotation_number"):
            command = f"DELETE FROM {table} WHERE quotation_number='{quotation_number}'"
            cursor.execute(command)
    cursor.execute("SET FOREIGN_KEY_CHECKS=0")
    bridge_db.commit()


def format_output(data, keys=None):
    description = cursor.description
    column_names = [col[0] for col in description]
    if keys is None:
        return {row[0]:dict(zip(column_names, row)) for row in data}
    if type(keys) == list:
        return {tuple(row[i] for i in keys):dict(zip(column_names, row)) for row in data}

def order_dict(dict, list):
    return {k:dict[k] for k in list}

def create_backup():
    import os
    backup_dir = 'B:/01.Bridge/app/backup'
    backup_file = 'mydatabase_backup_test.bak'
    backup_command = 'BACKUP DATABASE bridge TO DISK=\'' + os.path.join(backup_dir, backup_file) + '\''
    cursor.execute(backup_command)

# if __name__ == '__main__':
#     check_if_column_in_tabel("clients", "quotation_number")
#     delete_project("506000BK")