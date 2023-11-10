from requests.exceptions import RequestException, Timeout, HTTPError, ConnectionError
from services.static.paths import SYNC_STATUS_PATH
from services.static.urls import TOKEN_URL,DEVICES_URL
from loguru import logger
import logconfig
from openpyxl import Workbook,load_workbook
import requests
import creds
from datetime import datetime
import utils
import time
import os
import token_refresher
import column_names
# import services.service_utils as su


def is_response_valid(response):
    if not isinstance(response, requests.Response):
        logger.error(response)
        return False
   
    return True

def perform(url,access_token):
    last_sync_date = get_last_sync_date(url,access_token)
    if(last_sync_date is None):
        return None
    time_difference = get_time_difference(last_sync_date)
    return [last_sync_date,time_difference]

def get_url(user_id):
    return DEVICES_URL.replace("{}",user_id)

def get_last_sync_date(url,access_token):
    response = make_http_request(url,access_token)  
    
    if not is_response_valid(response):
        return None
    
    data = response.json()

    if(len(data) == 0):
        logger.error("Response is empty")
        return None

    last_sync = None
    for devices in data:
        if (devices.get("deviceVersion") == "Charge 5"):
            last_sync = devices.get("lastSyncTime")
            break
        
    return last_sync
    

def make_http_request(url,access_token):
    headers = {"authorization":f"Bearer {access_token}","accept-language":"en_US"}

    try:
        response = requests.get(url,headers=headers)
        response.raise_for_status()
        return response
    
    except ConnectionError as conn_err:
        return conn_err

    except Timeout as timeout_err:
        return timeout_err
    
    except HTTPError as http_err:
        return http_err
    
    except RequestException as req_err:
        return req_err

def get_time_difference(last_sync_date):
    sync_date = datetime.strptime(last_sync_date,"%Y-%m-%dT%H:%M:%S.%f")
    time_difference = (datetime.now() - sync_date).total_seconds()/(3600 * 24)
    return time_difference

logger.configure(**logconfig.configure_logs(SYNC_STATUS_PATH))

current_date = datetime.now().strftime("%m-%d-%Y")
file_name = f"sync_status_{current_date}.xlsx"
sync_workbook = utils.get_workbook(file_name, column_names.sync_status_column_names)
token_workbook = utils.get_workbook("tokens.xlsx",column_names.token_workbook_column_names)
token_sheet = token_workbook.active
sync_sheet = sync_workbook.active
headers = {"Authorization": f"Basic {creds.BASIC_TOKEN}", "Content-Type": "application/x-www-form-urlencoded"}
row_number = 2
for row in token_sheet.iter_rows(min_row=row_number,max_row = 2,values_only=True): 
    logger.info(f"processing user {row[1]}")
    access_token = row[3]
    row_data = [row[0],row[1],"","","",""]
    if(token_refresher.is_access_token_expired(row[5])):
        logger.info(f"token expired for user {row[4]}")
        access_token = token_refresher.generate_new_access_token(TOKEN_URL,row[2],row[4],headers,token_workbook,row_number)
        row_number+=1

        if(access_token is None):
            status = f"Error while generating new access token"
            logger.error(status)
            row_data[5] = status
            sync_sheet.append(row_data)
            continue
        else:
            time.sleep(5)

    url = get_url(row[4])      
    result = perform(url,access_token)
    curr_date = datetime.now()
    row_data[3] = datetime.strftime(curr_date, "%Y-%m-%dT%H:%M:%S")
    if(result is None):
        status =  "Error while retrieving last sync date"
        logger.error(status)
        row_data[5] = status
    else:
        row_data[2],row_data[4] = result[0],result[1]
        row_data[5] = "successful"
    sync_sheet.append(row_data)
    sync_workbook.save(file_name)
    logger.info(f"processed user {row[1]}")