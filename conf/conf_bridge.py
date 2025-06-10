import os

CONFIGURATION = {}
CONFIGURATION["len_per_line"]=100
CONFIGURATION["project_types"]=["Restaurant", "Office", "Commercial", "Group House", "Apartment", "Mixed-use Complex", "School", "Others"]
CONFIGURATION["services_list"]=["Mechanical Service", "Mechanical Review","Kitchen Ventilation" ,"CFD Service", "Electrical Service", "Hydraulic Service", "Fire Service", "Miscellaneous", "Installation"]
CONFIGURATION["adobe_address"]=r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
CONFIGURATION["backup_folder"]="B:\\"
CONFIGURATION["resource_dir"] = "T:\\00-Template-Do Not Modify\\00-Bridge template"
CONFIGURATION["template_dir"]=r'B:\02.Copilot\Database\template_dir'
CONFIGURATION["recycle_bin_dir"] = "B:\\01.Bridge\\app\\recycle_bin"
CONFIGURATION["database_dir"] = "B:\\01.Bridge\\app\\database"
CONFIGURATION["database_dir_copilot_bridge"] = "B:\\01.Bridge\\app\\database"
CONFIGURATION["email_password"]="PcE$yD2023"
CONFIGURATION["imap_server"]="sg2plzcpnl505807.prod.sin2.secureserver.net"
CONFIGURATION["smap_server"]="sg2plzcpnl505807.prod.sin2.secureserver.net"
CONFIGURATION["smap_port"]=587
CONFIGURATION["calculation_sheet"]="Preliminary Calculation v2.6.xlsx"

CONFIGURATION["Gray button style"]="QPushButton {background-color: rgb(200, 200, 200);color: rgb(0, 0, 0);border-style: outset;border-width: 2px;border-radius: 10px;border-color: rgb(80, 80, 80);font: bold 12px;padding: 6px;}QPushButton:pressed {background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);border-style: inset;}"
CONFIGURATION["Red button style"]="QPushButton {background-color: rgb(165, 42, 42);color: rgb(255, 255, 255);border-style: outset;border-width: 2px;border-radius: 10px;border-color: rgb(80, 80, 80);font: bold 10px;padding: 6px;}QPushButton:pressed {background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);border-style: inset;}"
CONFIGURATION['Text gray style']="font: 12pt 'Calibri'; border-style: outset; color: rgb(0, 0, 0); border-color: rgb(0, 0, 0);background-color: rgb(210, 210, 210);"
CONFIGURATION['Text red style']="font: 12pt 'Calibri'; border-style: outset; color: rgb(0, 0, 0); border-color: rgb(0, 0, 0);background-color: rgb(255, 0, 0);"
CONFIGURATION['Text green style']="font: 12pt 'Calibri'; border-style: outset; color: rgb(0, 0, 0); border-color: rgb(0, 0, 0);background-color: rgb(0, 128, 0);"
CONFIGURATION['Text purple style']="font: 12pt 'Calibri'; border-style: outset; color: rgb(255, 255, 255); border-color: rgb(0, 0, 0);background-color: rgb(128, 0, 128);"
CONFIGURATION['Text yellow style']="font: 12pt 'Calibri'; border-style: outset; color: rgb(0, 0, 0); border-color: rgb(0, 0, 0);background-color: rgb(255, 165, 0);"
CONFIGURATION['Text original style']="font: 12pt 'Calibri'; border-style: outset; color: rgb(0, 0, 0); border-color: rgb(0, 0, 0);background-color: rgb(170, 170, 170);"
CONFIGURATION['Text gray_white style']="font: 12pt 'Calibri'; color: rgb(0, 0, 0); border-color: rgb(255, 255, 255);background-color: rgb(210, 210, 210);"
CONFIGURATION['Text red_white style']="font: 12pt 'Calibri'; color: rgb(0, 0, 0); border-color: rgb(255, 255, 255);background-color: rgb(255, 0, 0);"
CONFIGURATION['Text green_white style']="font: 12pt 'Calibri'; color: rgb(0, 0, 0); border-color: rgb(255, 255, 255);background-color: rgb(0, 128, 0);"
CONFIGURATION['Text purple_white style']="font: 12pt 'Calibri'; color: rgb(255, 255, 255); border-color: rgb(255, 255, 255);background-color: rgb(128, 0, 128);"
CONFIGURATION['Text yellow_white style']="font: 12pt 'Calibri'; color: rgb(0, 0, 0); border-color: rgb(255, 255, 255);background-color: rgb(255, 165, 0);"
CONFIGURATION['Text original_white style']="font: 12pt 'Calibri'; color: rgb(0, 0, 0); border-color: rgb(255, 255, 255);background-color: rgb(170, 170, 170);"
CONFIGURATION['Text not change style']="font: 12pt 'Calibri'; color: rgb(0, 0, 0); border-color: rgb(255, 255, 255);background-color: rgb(170, 170, 170);"
CONFIGURATION['Table font']=11
CONFIGURATION['Service short dict'] = {"Mechanical Service": 'mech', "CFD Service": 'cfd',"Electrical Service": 'ele',"Hydraulic Service": 'hyd', "Fire Service": 'fire',"Mechanical Review": 'mechrev', "Miscellaneous": 'mis',"Installation": 'install', "Variation": 'var'}
CONFIGURATION['Service full dict'] = {"mech": 'Mechanical Service', "cfd": 'CFD Service', "ele": 'Electrical Service',"hyd": 'Hydraulic Service', "fire": 'Fire Service', "mechrev": 'Mechanical Review',"mis": 'Miscellaneous', "install": 'Installation', "var": 'Variation'}
CONFIGURATION["xero_client_id"] = "92582E6BA77A41F0B5076D3E5B442A24"
CONFIGURATION["xero_client_secret"] = "YmhTPLEHqGhjYFOK0uPowcpVsgdLJ2ZKYD_PKq-rjGJVQIml"




CONFIGURATION["email_username"]="bridge@pcen.com.au"
CONFIGURATION["from_addr"]="bridge@pcen.com.au"
CONFIGURATION["admin_addr"]="admin@pcen.com.au"
CONFIGURATION["working_dir"]="P:\\"
CONFIGURATION["accounting_dir"] = "A:\\00-Bridge Database"
CONFIGURATION["remittances_dir"] = "S:\\01.Expense invoice"
CONFIGURATION["bills_dir"] = "S:\\01.Expense invoice"
CONFIGURATION["xero_access_token_dir"] = os.path.join(CONFIGURATION["resource_dir"], "txt", "xero_access_token.txt")
CONFIGURATION["xero_refresh_token_dir"] = os.path.join(CONFIGURATION["resource_dir"], "txt", "xero_refresh_token.txt")
CONFIGURATION["Asana workplace"]='1198726743417674'
CONFIGURATION["xero_tenant_name"]="PREMIUM CONSULTING ENGINEERS PTY LTD"


# CONFIGURATION["email_username"]="yitong@pcen.com.au"
# CONFIGURATION["from_addr"]="yitong@pcen.com.au"
# CONFIGURATION["admin_addr"]="yitong@pcen.com.au"
# CONFIGURATION["working_dir"]="B:\\03. Bridge Testing Environment"
# CONFIGURATION["accounting_dir"] = "B:\\03. Bridge Testing Environment"
# CONFIGURATION["remittances_dir"] = "B:\\03. Bridge Testing Environment"
# CONFIGURATION["bills_dir"] = "B:\\03. Bridge Testing Environment"
# CONFIGURATION["xero_access_token_dir"] = r"B:\03. Bridge Testing Environment\xero_access_token.txt"
# CONFIGURATION["xero_refresh_token_dir"] = r"B:\03. Bridge Testing Environment\xero_refresh_token.txt"
# CONFIGURATION["Asana workplace"]="1208546788985072"
# CONFIGURATION["xero_tenant_name"] = "Demo Company (AU)"






