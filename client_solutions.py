import requests
from bs4 import BeautifulSoup as bs
import urllib.parse
import time
from datetime import datetime, timedelta
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from base64 import b64decode
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import re

# SEP : https://apidocs.symantec.com/home/saep
scope = ['https://www.googleapis.com/auth/drive']
gcredentials = ServiceAccountCredentials.from_json_keyfile_name('-.json', scope)
gc = gspread.authorize(gcredentials)
# sh = gc.create('USER List')
sh = gc.open('USER List')
# print(sh.id)

d = datetime.utcnow() - timedelta(hours=-9)
d = d.strftime('%Y%m%d')

# sh.share('yeonghoon2@gmail.com', perm_type='user', role='writer')
# worksheet = sh.worksheet(str(d)+'-solutions')
worksheet = sh.duplicate_sheet(source_sheet_id=0, insert_sheet_index=0, new_sheet_name=str(d)+'-Solutions')
worksheet.clear()
solutions_cell_list = worksheet.range('A1:K1')


solutions_cell_list[0].value = 'Name'
solutions_cell_list[1].value = 'Host'
solutions_cell_list[2].value = 'IP Address'
solutions_cell_list[3].value = 'MAC Address'
solutions_cell_list[4].value = 'SEP'
solutions_cell_list[5].value = 'Last Scaned Date'
solutions_cell_list[6].value = 'Updated Date'
solutions_cell_list[7].value = 'DLP'
solutions_cell_list[8].value = 'Updated Date'
solutions_cell_list[9].value = 'EDR'
solutions_cell_list[10].value = 'Updated Date'

worksheet.update_cells(solutions_cell_list)


def get_SEP_token():
    url = 'https://-:8446/sepm/api/v1/identity/authenticate'
    
    payload = {
        "username" : "-",
        "password" : "-",
        "domain" : ""
    }
    
    header = {
        'Content-Type' : 'application/json'
    }

    r = requests.post(url, headers = header, json = payload, verify=False)
    result = r.json()
    return result['token']

def get_SEP_totalPages():
    token = get_SEP_token()
    url = 'https://-:8446/sepm/api/v1/computers'
    header = {
        'Content-Type' : 'application/json',
        'Authorization': 'Bearer ' + str(token)
    }
    r = requests.get(url, headers = header, verify=False)
    result = r.json()
    return result['totalPages']

def get_SEP():  
    token = get_SEP_token()
    header = {
        'Content-Type' : 'application/json',
        'Authorization': 'Bearer ' + str(token)
    }
    totalPages = get_SEP_totalPages()
    for i in range(1, totalPages+1):
        url = 'https://-:8446/sepm/api/v1/computers?pageIndex='        
        url = url + str(i)
        r = requests.get(url, headers = header, verify=False)
        result = r.json()
        for i in result['content']:
            # print(i)
            values_list = worksheet.col_values(1) # Name column
            cell_list_index = 'A'+str(len(values_list)+1)+':K'+str(len(values_list)+1)
            cell_list = worksheet.range(cell_list_index)
            cell_list[0].value = i['logonUserName']
            computername = i['computerName']
            computername = computername.replace(' ','_')
            computername = computername.replace('의','ui')
            computername = computername.replace('’','')
            computername = computername.replace("'","")
            computername = computername.replace('-','_')
            computername = computername.replace('(','')
            computername = computername.replace(')','')
            computername = computername.lower()
            cell_list[1].value = computername
            cell_list[2].value = i['ipAddresses'][0]
            cell_list[3].value = i['macAddresses'][0]
            cell_list[4].value = i['agentVersion']
            if str(i['lastScanTime']) == '0':
                cell_list[5].value = ''
            else:
                cell_list[5].value = str(datetime.fromtimestamp(i['lastScanTime'] / 1e3))
            cell_list[6].value = str(datetime.fromtimestamp(i['lastUpdateTime'] / 1e3))
            worksheet.update_cells(cell_list)
            time.sleep(1)
get_SEP()


def get_DLP():
    options = Options()
    options.add_argument('--headless')
    driver = webdriver.Chrome(executable_path='./chromedriver', options=options)
    driver.implicitly_wait(3)
    actions = webdriver.ActionChains(driver)
    driver.get('https://-/ProtectManager/Logon')
    
    driver.find_element_by_name('username').send_keys('-')
    driver.find_element_by_name('j_password').send_keys('-')
    driver.find_element_by_xpath('//*[@id="login-link"]').click()
    driver.get('https://-/ProtectManager/enforce/admin/endpoint/health/troubleshoot#')
    driver.find_element_by_xpath('//*[@id="agentListTable_length"]/label/select/option[5]').click()
    html = driver.find_element_by_css_selector('#agentListTable > tbody').get_attribute('innerHTML')
    driver.close()
    
    sep_ip_list = worksheet.col_values(3)
    sep_host_list = worksheet.col_values(2)

    soup = bs(html, 'html')
    # print(soup)
    tmp_odd = soup.find_all('tr', class_='odd')
    tmp_even = soup.find_all('tr', class_='even')
    tmp = tmp_odd + tmp_even

    for j in range(1,len(sep_ip_list)):
        # print(sep_ip_list[j])
        # print(sep_host_list[j])
        for i in tmp:
            result = re.search('value\(operand_1\)=(.*)&amp',str(i))
            host = result.group(1)
            host = host.replace('-','_')
            host = host.lower()
            # print(host)
            # result = re.search('id="logged-in-user\d{1,5}">(.*)?</span></td><td><span id="agent-group', str(i))
            # name = result.group(1)
            result = re.search('id="agent-ip\d{1,5}">(.*)?</span></td><td><span id="agent-version', str(i))
            ip = result.group(1)
            # print(ip)
            result = re.search('id="agent-version\d{1,5}">(.*)?</span></td></tr>', str(i))
            version = result.group(1)
            result = re.search('id="agent-last-connection\d{1,5}">(.*)?</span></td><td><span id="agent-os', str(i))
            updated = result.group(1)
            if sep_host_list[j] == host or sep_ip_list[j] == ip :
                cell_list_index = 'H'+str(j+1)+':I'+str(j+1)
                cell_list = worksheet.range(cell_list_index)
                cell_list[0].value = version
                cell_list[1].value = updated
                worksheet.update_cells(cell_list)
                time.sleep(2)
                
get_DLP()