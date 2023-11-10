import os
import sys
current = os.path.dirname(os.path.realpath(__file__))
parent = os.path.dirname(current)
sys.path.append(parent)
import utils
from static.paths import ROOT_DIR

token_workbook = utils.get_workbook("tokens.xlsx",[])
sheet = token_workbook.active
for row in sheet.iter_rows(min_row = 2,values_only = True):
    id = utils.get_user_id(row[1])
    user_path = os.path.join(ROOT_DIR,id)
    folder = os.path.join(user_path,"Intraday")
    folder = os.path.join(folder,"Intraday")
    os.rmdir(folder)
    # file_names = os.listdir(folder)
    # for file in file_names:
    #     file_to_delete = os.path.join(folder,file)
    #     os.remove(file_to_delete)