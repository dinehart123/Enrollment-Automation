from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select


import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import time
import os
import os.path

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email_validator import validate_email, EmailNotValidError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

SHEET_ID = "***********************************"
SHEET_NAME  = "********** Automation Sheet"
RANGE_NAME = "Sheet1"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROW_SAVE_PATH = os.path.join(BASE_DIR, "row_save.txt")

SMTP_SERVER = '*************'  
SMTP_PORT = ************  
USERNAME = '***********' 
PASSWORD = '*****************'  

def get_google_credentials():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception as e:
                    print(f"Error refreshing credentials: {e}")
                    creds = None
            else:
                creds = None

    if not creds:
        flow = InstalledAppFlow.from_client_secrets_file("*********************.json", SCOPES)
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds

def scrap_sheets():
    creds = get_google_credentials()

    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SHEET_ID, range=RANGE_NAME).execute()
        return result.get("values", [])
    except HttpError as error:
        print(f"HTTP error occurred: {error}")
        return []


def write_to_google_sheet(sheet_name, row, column, value):
    try:
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        client = gspread.authorize(creds)
        spreadsheet = client.open(sheet_name)
        sheet = spreadsheet.sheet1
        sheet.update_cell(row, column, value)
        print(f"Successfully wrote value to {sheet_name} at row {row}, column {column}")
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"Spreadsheet {sheet_name} not found.")
    except gspread.exceptions.APIError as api_error:
        print(f"API error occurred: {api_error}")
    except Exception as e:
        print(f"An error occurred: {e}")


def get_first_empty_row_in_column(list):
    for i in range (len(list)):
        if(list[i][0] == ""):
            return i
    return len(list)


def send_email(sender, recipient, subject, body, attachment_path=None):
    try:
        validate_email(recipient)
        
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = recipient
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'plain'))
        
        if attachment_path:
            with open(attachment_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={attachment_path}')
                msg.attach(part)
        
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as server:
            server.starttls()
            server.login(USERNAME, PASSWORD)
            server.send_message(msg)
        print(f"Email sent to {recipient}")
    except EmailNotValidError as e:
        print(f"Invalid email address: {str(e)}")
    except smtplib.SMTPConnectError:
        print(f"Failed to connect to the server at {SMTP_SERVER}:{SMTP_PORT}")
    except smtplib.SMTPAuthenticationError:
        print("SMTP Authentication error. Check your username and password.")
    except smtplib.SMTPException as e:
        print(f"SMTP error occurred: {str(e)}")
    except Exception as e:
        print(f"Failed to send email to {recipient}: {str(e)}")
    

def type_in_id(id, keys):
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, id)))
    # input_element.clear()  # Clear the field before sending keys
    input_element.send_keys(keys)

def search_taxid(taxid):
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="filter2"]/span/input')))
    driver.execute_script("arguments[0].scrollIntoView();", input_element)
    input_element.send_keys(remove_hyphens(taxid))

def search_npi(npi):
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="filter3"]/span/input')))
    driver.execute_script("arguments[0].scrollIntoView();", input_element)
    input_element.send_keys(npi)

def search_name(name):
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="filter1"]/span/input')))
    driver.execute_script("arguments[0].scrollIntoView();", input_element)
    input_element.send_keys(name)

def check_checkbox(id):
    input_element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, id)))
    if input_element.is_selected():
        return 1
    else:
        return 0

def click_checkbox(id, clist):
    input_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, id)))
    driver.execute_script("arguments[0].scrollIntoView();", input_element)
    if clist == 1 and not input_element.is_selected():
        input_element.click()

def get_values(id):
    input_element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, id)))
    return input_element.get_attribute("value")

def click_values(id, vlist):
    input_element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, id)))
    input_element.send_keys(vlist)

def click_claim_type(response):
    if response == "HCFA":
        return "ClaimFormType_0"
    elif response == "UB":
        return "ClaimFormType_1"
    elif response == "Both":
        return "ClaimFormType_2"

def provider_existence(provider_list):
    try:
        search_taxid(provider_list[9])
        search_npi(provider_list[11])
        input_element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pars"]/tbody[2]/tr[1]/td[2]/a')))
    except:
        try:
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "clearfilters"))).click()
            search_taxid(provider_list[9])
            search_npi(provider_list[10])
            input_element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pars"]/tbody[2]/tr[1]/td[2]/a')))
        except:
            try:
                WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "clearfilters"))).click()
                search_taxid(provider_list[9])
                search_npi(provider_list[11])
                search_name(provider_list[4])
                input_element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pars"]/tbody[2]/tr[1]/td[2]/a')))
            except:
                WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "clearfilters"))).click()
                search_taxid(provider_list[9])
                search_npi(provider_list[10])
                search_name(provider_list[4])
                input_element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pars"]/tbody[2]/tr[1]/td[2]/a')))
    try:
        input_element.click()
    except:
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "clearfilters"))).click()
        pass


def save_provider_and_refresh():
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnSave"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mainform"]/div[3]/div[1]/div/ul/li[3]/a'))).click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "clearfilters"))).click()

def remove_hyphens(input_str):
    return input_str.replace('-', '')

def remove_underscores (input_str):
    return input_str.replace('_', '')

def close_contact():
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="editcontactmodal"]/div[3]/button[2]'))).click()

def get_products_and_percentages(provider_list):
    pplist = []
    search_taxid(remove_hyphens(provider_list[9]))

    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pars"]/tbody[2]/tr[1]/td[2]/a'))).click()

    pplist.append(check_checkbox("GHCoverage"))
    pplist.append(check_checkbox("WCCoverage"))
    pplist.append(check_checkbox("AutoCoverage"))
    pplist.append(check_checkbox("MACoverage"))

    pplist.append(get_values("DiscountPctOff"))
    pplist.append(get_values("DiscountFeeSchedule"))
    pplist.append(get_values("MAPct"))
    pplist.append(get_values("RBRVSPct"))

    # print("ContractID is: " + get_values("ContractID"))
    pplist.append(get_values("ContractID"))

    if(check_checkbox("ClaimFormType_2")):
        pplist.append("Both")
    elif(check_checkbox("ClaimFormType_1")):
        pplist.append("UB")
    elif(check_checkbox("ClaimFormType_0")):
        pplist.append("HCFA")
        


    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mainform"]/div[3]/div[1]/div/ul/li[3]/a'))).click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "clearfilters"))).click()

    return pplist


def check_provider_existence(taxid,pnpi,gnpi):
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "clearfilters"))).click()
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="filter2"]/span/input')))
    driver.execute_script("arguments[0].scrollIntoView();", input_element)
    input_element.send_keys(remove_hyphens(taxid)) #provider_list[9]
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="filter3"]/span/input')))
    input_element.send_keys(pnpi)
    input_element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pars"]/tbody[2]/tr[1]/td[2]/a')))
    if input_element.click() == "NoSuchElementException":

        input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="filter3"]/span/input')))
        input_element.send_keys(gnpi)




def load_browser():
    driver.get("https://*******.***********.com/*****.****")
    type_in_id("UserName", "*************")
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,"Password")))
    input_element.send_keys("**************" + Keys.ENTER)

def new_location(provider_list):
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "newlocation"))).click()
    type_in_id("L_Address", provider_list[14])
    type_in_id("L_Suite",provider_list[15])
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,"L_Zipcode")))
    input_element.send_keys(Keys.CONTROL, 'a')
    input_element.send_keys(provider_list[16])
    type_in_id("L_City", provider_list[17])
    type_in_id("L_State", provider_list[18])
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,"L_Phone")))
    input_element.send_keys(Keys.CONTROL, 'a')
    input_element.send_keys(provider_list[19])
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,"L_Fax")))
    input_element.send_keys(Keys.CONTROL, 'a')
    input_element.send_keys(provider_list[20])
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnSaveLocation"))).click()

def more_location(provider_list, start):
    # try:
        print(start)
        for i in range(start,len(provider_list),7):
            if(provider_list[i] != ""):
                print("adding new location")
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "newlocation"))).click()
                type_in_id("L_Address", provider_list[i])
                type_in_id("L_Suite",provider_list[i+1])
                input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,"L_Zipcode")))
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(provider_list[i+2])
                type_in_id("L_City", provider_list[i+3])
                type_in_id("L_State", provider_list[i+4])
                input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,"L_Phone")))
                input_element.send_keys(Keys.CONTROL, 'a')
                try:
                    if(provider_list[i+5] != ""):
                        input_element.send_keys(provider_list[i+5])
                    else:
                        input_element.send_keys(provider_list[19])
                except:
                    input_element.send_keys(provider_list[19])
                print("typed phone number")
                try:
                    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,"L_Fax")))
                    input_element.send_keys(Keys.CONTROL, 'a')
                    input_element.send_keys(provider_list[i+6])
                except:
                    pass
                print("about to save")
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnSaveLocation"))).click()
    # except:
    #     pass

def new_billing(provider_list):
    type_in_id("BillingAddress", provider_list[21])
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,"BillingZipcode")))
    input_element.send_keys(Keys.CONTROL, 'a')
    input_element.send_keys(provider_list[22])    
    type_in_id("BillingCity", provider_list[23])
    type_in_id("BillingState", provider_list[24])
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,"BillingPhone")))
    input_element.send_keys(Keys.CONTROL, 'a')
    input_element.send_keys(provider_list[25])
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,"BillingFax")))
    input_element.send_keys(Keys.CONTROL, 'a')
    input_element.send_keys(provider_list[26])


def add_par(provider_list, pplist):
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "newbutton"))).click()
    type_in_id("LastName", provider_list[3])
    type_in_id("FirstName", provider_list[4])
    type_in_id("MiddleInit", provider_list[5])
    type_in_id("Degree", provider_list[6])
    type_in_id("EffectiveDate", provider_list[7])
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,"TaxID")))
    input_element.send_keys(Keys.CONTROL, 'a')
    input_element.send_keys(provider_list[9])
    type_in_id("BillingNPINumber", provider_list[11])
    type_in_id("ProviderTypeID", provider_list[12])
    type_in_id("FacilityGroupName", provider_list[30])
    Select(driver.find_element("id", "ContractID")).select_by_value(pplist[8])


    click_checkbox("GHCoverage", pplist[0])
    click_checkbox("WCCoverage", pplist[1])
    click_checkbox("AutoCoverage", pplist[2])
    click_checkbox("MACoverage", pplist[3])
    click_values("DiscountPctOff", pplist[4])
    click_values("DiscountFeeSchedule", pplist[5])
    click_values("MAPct", pplist[6])
    click_values("RBRVSPct", pplist[7])
    if(provider_list[13] != "HCFA" or provider_list[13] != "UB" or provider_list[13] != "Both"):
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,click_claim_type(pplist[9])))).click()
    else:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,click_claim_type(provider_list[13])))).click()
    new_billing(provider_list)
    new_location(provider_list)
    try:
        if provider_list[32] != "":
            more_location(provider_list,32)
    except:
        pass
    type_in_id("PrimarySpecialtyID", provider_list[27])
    try:
        type_in_id("OtherSpecialtyID", provider_list[28])
    except:
        pass
    try:
        type_in_id("OtherSpecialty2ID", provider_list[29])
    except:
        pass



if __name__ == "__main__":
    body = ""
    enrollment_list = scrap_sheets()

    service = Service(executable_path="chromedriver.exe")
    options = webdriver.ChromeOptions()
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--allow-insecure-localhost")
    options.add_argument("--allow-running-insecure-content")
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()

    load_browser()

    curr_row = get_first_empty_row_in_column(enrollment_list)
    first_row = curr_row
    print("starting row is:")
    print(curr_row)
    
    pplist = get_products_and_percentages(enrollment_list[curr_row])
    for i in range(curr_row, len(enrollment_list)):
        # print("First column is: " + enrollment_list[i][0])
        if curr_row >= first_row and enrollment_list[i][9] != enrollment_list[i-1][9] and enrollment_list[i][0] == "":
            pplist = get_products_and_percentages(enrollment_list[i]) 
        add_par(enrollment_list[i], pplist)

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnSave"))).click()
        print("added")
        par_id = get_values("ParId")
        write_to_google_sheet(SHEET_NAME,i+1,1,par_id)
        save_provider_and_refresh()
        curr_row = i+1
        # body = "Hello, the provider " + str(enrollment_list[i][3]) + " " + str(enrollment_list[i][4]) + " has been successfully enrolled. Effective date: " + enrollment_list[i][7] +"; ID Number: " + str(par_id)
        # send_email('*******@*******.com', str(enrollment_list[i][2]), str(enrollment_list[i][3]) + " " + str(enrollment_list[i][4]) + " Enrollment Complete", body, None)


    print("exited")



    time.sleep(20000)


    driver.quit()



