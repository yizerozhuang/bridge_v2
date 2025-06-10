import ast
import os

CONFIGURATION = {}
CONFIGURATION["temp-a3-restaurant"] = "T:\\00-Template-Do Not Modify\\09-AutoCAD and Bluebeam\\01-Mech template\\A3-Restaurant"
CONFIGURATION["temp-a1-restaurant"] = "T:\\00-Template-Do Not Modify\\09-AutoCAD and Bluebeam\\01-Mech template\\A1-Restaurant"
CONFIGURATION["temp-a0-restaurant"] = "T:\\00-Template-Do Not Modify\\09-AutoCAD and Bluebeam\\01-Mech template\\A0-Restaurant"
CONFIGURATION["temp-a3-others"] = "T:\\00-Template-Do Not Modify\\09-AutoCAD and Bluebeam\\01-Mech template\\A3-Others"
CONFIGURATION["temp-a1-others"] = "T:\\00-Template-Do Not Modify\\09-AutoCAD and Bluebeam\\01-Mech template\\A1-Others"
CONFIGURATION["temp-a0-others"] = "T:\\00-Template-Do Not Modify\\09-AutoCAD and Bluebeam\\01-Mech template\\A0-Others"

CONFIGURATION["mm_to_pixel"]=2.83465
CONFIGURATION["c_temp_dir"] = "C:\\Copilot_template"
CONFIGURATION["bluebeam_dir_file"] = "C:\\Copilot_database\\Bluebeam_dir.txt"
CONFIGURATION["password_file"] = "C:\\Copilot_database\\password.txt"
CONFIGURATION["bluebeam_engine_dir"]=r'C:\Program Files\Bluebeam Software\Bluebeam Revu\2017\Script\ScriptEngine.exe'
CONFIGURATION["bluebeam_dir"]=r"C:\Program Files\Bluebeam Software\Bluebeam Revu\2017\Revu\Revu.exe"
if os.path.exists(CONFIGURATION["bluebeam_dir_file"]):
    with open(CONFIGURATION["bluebeam_dir_file"], 'r', encoding='utf-8') as file:
        file_content = file.read()
        content = ast.literal_eval(file_content)
        CONFIGURATION["bluebeam_engine_dir"] = content['bluebeam_engine_dir']
        CONFIGURATION["bluebeam_dir"] = content['bluebeam_dir']
CONFIGURATION["trans_timewait_dir"] = 'B:\\02.Copilot\\Database\\trans_wait_time.json'
CONFIGURATION["trans_dir"] = "B:\\02.Copilot\\Copilot_file_trans"
CONFIGURATION["backup_folder"]="B:\\"
CONFIGURATION["database_dir"] = "B:\\02.Copilot\\Database"
CONFIGURATION["project_folder"]="P:\\"
CONFIGURATION["database_dir_ori"] = "B:\\01.Bridge\\app\\database"
CONFIGURATION['Table font']=10
CONFIGURATION["recycle_bin_dir"] = "B:\\02.Copilot\\recycle_plot"
CONFIGURATION['cancel_task_txt']='B:\\02.Copilot\\Copilot_file_trans\\cancelled task.txt'


CONFIGURATION["Gray button style"]="QPushButton {background-color: rgb(200, 200, 200);color: rgb(0, 0, 0);border-style: outset;border-width: 2px;border-radius: 10px;border-color: rgb(80, 80, 80);font: bold 12px;min-width: 5.5em;padding: 6px;}QPushButton:pressed {background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);border-style: inset;}"
CONFIGURATION["Color button style 1"]="QPushButton {background-color: rgb(165, 42, 42);color: rgb(255, 255, 255);border-style: outset;border-width: 2px;border-radius: 10px;border-color: rgb(80, 80, 80);font: bold 12px;min-width: 5.5em;padding: 6px;}QPushButton:pressed {background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);border-style: inset;}"
CONFIGURATION["Color button style 2"]="QPushButton {background-color: rgb(51, 153, 204);color: rgb(255, 255, 255);border-style: outset;border-width: 2px;border-radius: 10px;border-color: rgb(80, 80, 80);font: bold 12px;min-width: 5.5em;padding: 6px;}QPushButton:pressed {background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);border-style: inset;}"
CONFIGURATION["Color button style 3"]="QPushButton {background-color: rgb(0, 51, 102);color: rgb(255, 255, 255);border-style: outset;border-width: 2px;border-radius: 10px;border-color: rgb(80, 80, 80);font: bold 12px;min-width: 5.5em;padding: 6px;}QPushButton:pressed {background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);border-style: inset;}"
CONFIGURATION["Color button style 4"]="QPushButton {background-color: rgb(51, 102, 102);color: rgb(255, 255, 255);border-style: outset;border-width: 2px;border-radius: 10px;border-color: rgb(80, 80, 80);font: bold 12px;min-width: 5.5em;padding: 6px;}QPushButton:pressed {background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);border-style: inset;}"
CONFIGURATION["Color button style 5"]="QPushButton {background-color: rgb(204, 153, 51);color: rgb(255, 255, 255);border-style: outset;border-width: 2px;border-radius: 10px;border-color: rgb(80, 80, 80);font: bold 12px;min-width: 5.5em;padding: 6px;}QPushButton:pressed {background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);border-style: inset;}"
CONFIGURATION["Color button style 6"]="QPushButton {background-color: rgb(153, 102, 153);color: rgb(255, 255, 255);border-style: outset;border-width: 2px;border-radius: 10px;border-color: rgb(80, 80, 80);font: bold 12px;min-width: 5.5em;padding: 6px;}QPushButton:pressed {background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);border-style: inset;}"



CONFIGURATION['Text red style']="font: 12pt 'Calibri'; border-style: outset; color: rgb(0, 0, 0); border-color: rgb(0, 0, 0);background-color: rgb(255, 0, 0);"
CONFIGURATION['Text normal style']="font: 12pt 'Calibri'; border-style: outset; color: rgb(0, 0, 0); border-color: rgb(0, 0, 0);background-color: rgb(170, 170, 170);"