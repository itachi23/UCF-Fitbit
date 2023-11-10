import requests
import openpyxl
import utils
import column_names
import creds
from datetime import datetime
#58507fa233ce5e45b46bb57b7bccddb29fcac864
#df40ecf557122dc6e87c43a032737ade08f13384
def get_tokens(authorization_code,code_verifier):
    data = f"grant_type=authorization_code&code={authorization_code}&redirect_uri={creds.REDIRECT_URL}&code_verifier={code_verifier}"
    headers = {"Authorization": f"Basic {creds.BASIC_TOKEN}", "Content-Type": "application/x-www-form-urlencoded"}
    TOKEN_URL = 'https://api.fitbit.com/oauth2/token'
    response = requests.post(TOKEN_URL,data = data,headers = headers)
    print(response.text)
    if(response.status_code != 200):
        return None
    data = response.json()
    refresh_token = data['refresh_token']
    access_token = data['access_token']
    user_id = data['user_id']
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return [refresh_token,access_token,user_id,created_at]

user_workbook = utils.get_workbook("users_with_google_accounts.xlsx",[])
token_workbook = utils.get_workbook("tokens.xlsx",column_names.token_workbook_column_names)
token_worksheet = token_workbook.active
user_sheet = user_workbook.active
for row in user_sheet.iter_rows(min_row=13,max_row=13,values_only=True):
    row_data = [row[0],row[1],"","","",""]
    authorization_code = "24be215e08a3a13dcddc06e393a53896905af11c"
    code_verifier = "4q1h0q0j382v3d1c5n455x6l1o675n1a2d0t203g2r6e2f6i3t110o6h3q6x1u2d723e4d2r6z013b2o6w5o1q393l302x1o3o3v704k5b5s5l2x3f445s3x4k1j5q71"
    data = get_tokens(authorization_code,code_verifier)
    if(data is None):
        print("could not get tokens")
        break
    row_data[2] = data[0]
    row_data[3] = data[1]
    row_data[4] = data[2]
    row_data[5] = data[3]
    token_worksheet.append(row_data)
    token_workbook.save("tokens.xlsx")