from openpyxl import Workbook
from openpyxl import load_workbook 
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import requests
import creds
import codes
import os
import time
import utils
from datetime import datetime
import proxy_test

def get_authorization_code(url):
    return url[29:69]

def fitbit_login(driver,URL,code_verifier,row,workbook):
    print(row)
    sheet = workbook.active
    mail_id, pwd, user_name = row
    row_data = []
    refresh_token = ""
    access_token = ""
    user_id = ""
    curr_url = ""
    created_at = ""
    try:
        driver.get(URL)
        time.sleep(6)
        username = driver.find_element(By.ID, 'ember591')
        username.clear()
        username.send_keys(mail_id)
        password = driver.find_element(By.ID, 'ember592')
        password.clear()
        password.send_keys(pwd)
        sign_in_button = driver.find_element(By.ID, 'ember632')
        sign_in_button.click()
        time.sleep(7)
        if("localhost" in driver.current_url):
            curr_url = driver.current_url               
        else:
            scope = driver.find_element(By.ID,'selectAllScope')
            scope.click()
            allow_button = driver.find_element(By.ID, 'allow-button')
            allow_button.click()
            time.sleep(7)
            curr_url = driver.current_url
        refresh_token, access_token,user_id = get_tokens(curr_url,TOKEN_URL,headers,code_verifier)
        print(f"user_name {user_name} mail-id {mail_id} refresh_token {refresh_token}")
        created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        row_data = [user_name,mail_id,refresh_token,access_token,user_id,created_at]
        sheet.append(row_data)
    except Exception as e:
        print(f"An error occurred while processing user {mail_id}", e)


def get_tokens(curr_url,TOKEN_URL,headers,code_verifier):
    authorization_code = get_authorization_code(curr_url)
    return get_access_token(TOKEN_URL,headers,authorization_code,code_verifier)

def get_access_token(TOKEN_URL,headers,authorization_code,code_verifier):
    data = f"grant_type=authorization_code&code={authorization_code}&redirect_uri={creds.REDIRECT_URL}&code_verifier={code_verifier}"
    response = requests.post(TOKEN_URL,data = data,headers = headers)
    data = response.json()
    refresh_token = data['refresh_token']
    access_token = data['access_token']
    user_id = data['user_id']
    return [refresh_token,access_token,user_id]

def perform(driver,new_workbook, user_workbook):
    user_sheet = user_workbook.active
    new_workbook_sheet = new_workbook.active
    # row_num = 6    
    for row in user_sheet.iter_rows(min_row=6,max_row=10,values_only=True):  
        # PROXY = PROXIES[row_num%30]
        CODE_VERIFIER, CODE_CHALLENGE = codes.generate_codes()
        OAUTH_URL = f'https://www.fitbit.com/oauth2/authorize?client_id={creds.CLIENT_ID}&response_type=code&code_challenge={CODE_CHALLENGE}&code_challenge_method=S256&scope=activity%20heartrate%20location%20nutrition%20oxygen_saturation%20profile%20respiratory_rate%20settings%20sleep%20social%20temperature%20weight%20electrocardiogram%20cardio_fitness%20social'
        try:
            fitbit_login(driver,OAUTH_URL,CODE_VERIFIER,row, new_workbook)
            driver.quit()
            driver = None
            driver = webdriver.Chrome()
        except Exception as e:
            print("An error occurred", e)
        new_workbook.save('tokens.xlsx')

TOKEN_URL = 'https://api.fitbit.com/oauth2/token'
headers = {"Authorization": f"Basic {creds.BASIC_TOKEN}", "Content-Type": "application/x-www-form-urlencoded"}
driver = webdriver.Chrome()
token_workbook = utils.get_workbook("tokens.xlsx",["Name","Mail","Refresh Token", "Access Token", "User id","Created at"])
user_workbook = utils.get_workbook("Creds.xlsx",[])
perform(driver,token_workbook, user_workbook)



        


