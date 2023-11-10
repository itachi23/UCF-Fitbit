import os
from openpyxl import workbook
import utils
from services.static.paths import ROOT_DIR

def create_sub_directories(root_path,folders):
    for dir in folders:
       child =  os.path.join(root_path,dir[0])
       os.makedirs(child)
       if(len(dir) > 1):
           for sub_dir in dir[1:]:
              inner_child = os.path.join(child,sub_dir)
              os.makedirs(inner_child)

# root_directory = r"D:\fitbit_user_data"
folders_2 = [["Intraday"]]
folders = [["Heart_Rate"],["SPO2"],["Breathing_Rate"],["Heart Rate Variability"]]
token_workbook = utils.get_workbook("tokens.xlsx",[])
# user_path_workbook = utils.get_workbook("users_path.xlsx",["user_id","mail","path"])
sheet = token_workbook.active
# path_sheet = user_path_workbook.active
for row in sheet.iter_rows(min_row = 2,values_only = True):
    user_mail = row[1]
    user_id = row[4]
    id = ""
    for i in range(len(user_mail)):
        value = user_mail[i]
        if value.isdigit():
            id = id + value
    folder_path = os.path.join(ROOT_DIR,id)
    folder_path = os.path.join(folder_path, "Intraday")
    # os.makedirs(folder_path)
    print(folder_path)
    create_sub_directories(folder_path,folders)
    data = [user_id, user_mail, folder_path]
    # path_sheet.append(data)
    # user_path_workbook.save("users_path.xlsx")

