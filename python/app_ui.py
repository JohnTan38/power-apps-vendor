from selenium import webdriver # 1 login 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time

def login_esker():
    driver = webdriver.Chrome()
    driver.get("https://az3.ondemand.esker.com/ondemand/webaccess/asf/home.aspx")
    driver.maximize_window()
    time.sleep(1)

    driver.find_element(By.XPATH, '//*[@id="ctl03_tbUser"]').send_keys("john.tan@sh-cogent.com.sg")
    driver.find_element(By.XPATH, '//*[@id="ctl03_tbPassword"]').send_keys("Esker3838")
    driver.find_element(By.XPATH, '//*[@id="ctl03_btnSubmitLogin"]').click()
    time.sleep(3)
    return driver
def hover(driver, x_path):
    elem_to_hover = driver.find_element(By.XPATH, x_path)
    hover = ActionChains(driver).move_to_element(elem_to_hover)
    hover.perform()
#time.sleep(2)
driver = login_esker()
def hover_arrow(driver, x_path):
    elem_to_hover = driver.find_element(By.XPATH, x_path)
    hover = ActionChains(driver).move_to_element(elem_to_hover)
    hover.perform()

"""
x_path_hover = '//*[@id="mainMenuBar"]/td/table/tbody/tr/td[36]/a/div' #arrow
hover_arrow(driver, x_path_hover)

try:
    #drop_down=driver.find_element(By.XPATH, '//*[@id="DOCUMENT_TAB_100872215"]/a/div[2]').click()
    tables=driver.find_element(By.XPATH, '//*[@id="CUSTOMTABLE_TAB_100872176"]')
    tables.click()
    time.sleep(1)
except Exception as e:
    print(e) #VENDOR INVOICES (SUMMARY) #TABLES
time.sleep(1)
"""

import pyautogui, os #20251002 great! #2
from pathlib import Path
import win32com.client  #esker vendor update Great 20241129!
import time #create inbox subfolder, download attachments, move email to subfolder.
import os, re
import datetime as dt
import pandas as pd
import numpy as np
import openpyxl
from datetime import datetime
"""
email = 'john.tan@sh-cogent.com.sg'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.Folders(email).Folders("Inbox")

date_time = dt.datetime.now()
lastTwoDaysDateTime = dt.datetime.now() - dt.timedelta(days = 2)
newDay = dt.datetime.now().strftime('%Y%m%d')
path_vendor_update = r"C:/Users/john.tan/Documents/power_apps_esker_vendor/esker_vendor_update/"

sub_folder1 = inbox.Folders['esker_vendor']
try:
    myfolder = sub_folder1.Folders[newDay] #check if fldr exists, else create
    #print('folder exists')
except:
    sub_folder1.Folders.Add(newDay)
    #print('subfolder created')
dest = sub_folder1.Folders[newDay]

i=0
#messages = inbox.Items
messages = sub_folder1.Items
lastTwoDaysMessages = messages.Restrict("[ReceivedTime] >= '" +lastTwoDaysDateTime.strftime('%d/%m/%Y %H:%M %p')+"'") #AND "urn:schemas:httpmail:subject"Demurrage" & "'Bill of Lading'")
for message in lastTwoDaysMessages:
        if (("ESKER VENDOR EMAIL") or ("esker vendor email")) in message.Subject:
            
            for attachment in message.Attachments:
                      #print(attachment.FileName)
                      try:
                            attachment.SaveAsFile(os.path.join(path_vendor_update, str(attachment).split('.xlsx')[0]+'.xlsx'))#'_'+str(i)+
                            i +=1
                      except Exception as e: 
                            print(e)

path_vendor_update = r"C:/Users/john.tan/Documents/power_apps_esker_vendor/esker_vendor_update/"
paths = [(p.stat().st_mtime, p) for p in Path(path_vendor_update).iterdir() if p.suffix == ".xlsx"] #Save .xlsx files paths and modification time into paths
paths = sorted(paths, key=lambda x: x[0], reverse=True) # Sort by modification time
last_path = paths[0][1] ## Get the last modified file
#excel_vendor_update = 'vendor_update.xlsx'
try:
    vendor_update = pd.read_excel(last_path, sheet_name='vendors', engine='openpyxl')
except FileNotFoundError:
    print(f"Error: File '{last_path}' not found in path '{path_vendor_update}'.")
    time.sleep(10)
    exit()
"""
def format_vendor_data(df_vendor_update):
    """
    Formats the vendor data in the DataFrame, ensuring:
    1. 'vendor_number' and 'postal_code' are numeric
    2. 'postal_code' has 6 digits
    3. Empty or missing values in 'street', 'city', 'postal_code', and 'country_code' are replaced with empty strings
    Args:
        df_vendor_update (pd.DataFrame): The DataFrame containing vendor data.
    Returns:
        pd.DataFrame: The formatted DataFrame.
    """
    # Convert 'vendor_number' and 'postal_code' to numeric
    df_vendor_update['vendor_number'] = pd.to_numeric(df_vendor_update['vendor_number'], errors='coerce')
    #df_vendor_update['postal_code'] = pd.to_numeric(df_vendor_update['postal_code'], errors='coerce')
    
    # Ensure 'postal_code' has 6 digits
    #df_vendor_update['postal_code'] = df_vendor_update['postal_code'].astype(str).str.zfill(6)
    df_vendor_update.fillna('', inplace=True) # Replace empty or missing values with empty strings
    return df_vendor_update

vendor_update = pd.DataFrame({
    'company_code': ['SG80'],
    'vendor_number': ['1000338436'],
    'vendor_name': ['SPEEDYLINK LOGISTICS SDN BHD']
    
})

df_vendor_update = format_vendor_data(vendor_update)

def create_log_file(path):
    """
    Checks if a log file exists at the specified path.
    If not, creates a new one with the current date and time.
    """
    filename = f"log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
    full_path = os.path.join(path, filename)

    if not os.path.exists(full_path):
        with open(full_path, 'w') as f:
            f.write("")  # Create an empty file

    return full_path
"""
path_vendor_update = r"C:/Users/john.tan/Documents/power_apps_esker_vendor/esker_vendor_update/"
log_file = create_log_file(path_vendor_update+'Log/') # Create the log file if it doesn't exist

list_company_code =[]
list_vendor_number =[]
"""
def start_time():
    start_time=datetime.now()
    return start_time

def vendor_update_process(df_vendor_update):
    """
    Processes vendor updates by iterating through the provided DataFrame.
    For each vendor, it performs a series of actions to update vendor information.
    Args:
        df_vendor_update (pd.DataFrame): The DataFrame containing vendor data.
    """
    for i in range(len(df_vendor_update)):
        print(f"company_code {df_vendor_update.loc[i, 'company_code']}")

        pyautogui.moveTo(100,330, duration=2)
        pyautogui.click()
        pyautogui.typewrite('S2P - Vendors')
        pyautogui.press('ENTER')
        try:
            s2p_vendors=driver.find_element(By.XPATH, '//*[@id="ViewSelector"]/div/div/div/div[1]/div[1]/span')
            #print('updating vendors')
        except Exception as e:
            print(e)

        try:
            btn_new=driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_CommonActionList"]/tbody/tr/td[1]/a').click()
        except Exception as e:
            print(e)
        time.sleep(1)
        try:
            btn_continue=driver.find_element(By.XPATH, '//*[@id="form-container"]/div[5]/div[3]/div[2]/div[3]/a[1]').click()
        except Exception as e:
            print(e)
        time.sleep(1)

        try:
            input_company_code=driver.find_element(By.XPATH, '//*[@id="DataPanel_eskCtrlBorder_content"]/div/div/table/tbody/tr[1]/td[2]/div/input')
            input_company_code.send_keys(df_vendor_update.loc[i, 'company_code'])
            time.sleep(0.5)
        except Exception as e:
            print(e)           
        actions = ActionChains(driver)
        actions.send_keys(Keys.TAB).perform()
        try:
            input_vendor_number=driver.find_element(By.XPATH, '//*[@id="DataPanel_eskCtrlBorder_content"]/div/div/table/tbody/tr[2]/td[2]/div/input')
            input_vendor_number.send_keys(str(df_vendor_update.loc[i, 'vendor_number']))
            vendor_number = df_vendor_update.loc[i, 'vendor_number']
            time.sleep(0.5)
        except Exception as e:
            print(e)
        actions.send_keys(Keys.TAB).perform()
        try:
            input_vendor_name=driver.find_element(By.XPATH, '//*[@id="DataPanel_eskCtrlBorder_content"]/div/div/table/tbody/tr[3]/td[2]/div/input')
            input_vendor_name.send_keys(df_vendor_update.loc[i, 'vendor_name'])
            time.sleep(0.5)
        except Exception as e:
            print(e)
        """
        actions.send_keys(Keys.TAB*15).perform()
        try:
            input_street=driver.find_element(By.XPATH, '//*[@id="VendorAddress_eskCtrlBorder_content"]/div/div/table/tbody/tr[2]/td[2]/div/input')
            input_street.send_keys(df_vendor_update.loc[i, 'street'])
            time.sleep(0.5)
        except Exception as e:
            print(e)
        actions.send_keys(Keys.TAB*2).perform()
        try:
            input_city=driver.find_element(By.XPATH, '//*[@id="VendorAddress_eskCtrlBorder_content"]/div/div/table/tbody/tr[4]/td[2]/div/input')
            input_city.send_keys(df_vendor_update.loc[i, 'city'])
            time.sleep(0.5)
        except Exception as e:
            print(e)
        actions.send_keys(Keys.TAB).perform()
        try:
            input_postal_code=driver.find_element(By.XPATH, '//*[@id="VendorAddress_eskCtrlBorder_content"]/div/div/table/tbody/tr[5]/td[2]/div/input')
            input_postal_code.send_keys(df_vendor_update.loc[i, 'postal_code'])
            time.sleep(0.5)
        except Exception as e:
            print(e)
        actions.send_keys(Keys.TAB*2).perform()
        try:
            input_country_code=driver.find_element(By.XPATH, '//*[@id="VendorAddress_eskCtrlBorder_content"]/div/div/table/tbody/tr[7]/td[2]/div/input')
            input_country_code.send_keys(df_vendor_update.loc[i, 'country_code'])
            time.sleep(0.5)
        except Exception as e:
            print(e)
        actions.send_keys(Keys.TAB).perform()
        """

        try:
            btn_save=driver.find_element(By.XPATH, '//*[@id="form-footer"]/div[1]/a[1]/span')
            btn_save.click()
            time.sleep(1)
        except Exception as e:
            print(f"failed to save: {e}")

        list_company_code.append(df_vendor_update.loc[i, 'company_code'])
        list_vendor_number.append(df_vendor_update.loc[i, 'vendor_number'])
        return list_vendor_number

    with open(log_file, 'a') as f:  # Open in append mode
        f.write(f"Process completed: {datetime.now()}\n")
        f.write(f"Updated entities: {', '.join(list_company_code)}\n")
        f.write(f"Updated vendor: {vendor_number}\n")

def log_entry(start_time):
    with open(log_file, 'a') as f:  # Open in append mode
        f.write(f"Process start: {start_time}")
        f.write(f"Process completed: {datetime.now()}\n")
        f.write(f"Updated entities: {', '.join(list_company_code)}\n")
        f.write(f"Updated vendor: {list_vendor_number}\n")


if __name__ == "__main__":
    login_esker()
    driver=login_esker()
    x_path_hover = '//*[@id="mainMenuBar"]/td/table/tbody/tr/td[36]/a/div' #arrow
    hover_arrow(driver, x_path_hover)

    try:
        #drop_down=driver.find_element(By.XPATH, '//*[@id="DOCUMENT_TAB_100872215"]/a/div[2]').click()
        tables=driver.find_element(By.XPATH, '//*[@id="CUSTOMTABLE_TAB_100872176"]')
        tables.click()
        time.sleep(1)
    except Exception as e:
        print(e) #VENDOR INVOICES (SUMMARY) #TABLES
    time.sleep(1)

    df_vendor_update = format_vendor_data(vendor_update)
    path_vendor_update = r"C:/Users/john.tan/Documents/power_apps_esker_vendor/esker_vendor_update/"
    log_file = create_log_file(path_vendor_update+'Log/') # Create the log file if it doesn't exist

    list_company_code = []
    list_vendor_number = []
    vendor_update_process(df_vendor_update)
    time.sleep(3)
    driver.quit()
    log_entry()
    print("Process completed.")
    time.sleep(2)
