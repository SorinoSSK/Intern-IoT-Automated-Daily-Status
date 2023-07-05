import json
import time
import requests
import datetime
import os, shutil
import pandas as pd
from os.path import exists

from docx import Document
from docx.oxml.shared import qn
from docx.oxml.xmlchemy import OxmlElement

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import ssl



### Global Variables
MADS_FILE_NAME              = "mads_config.xlsx"
VFT_FILE_NAME               = "vft_config.xlsx"
VFT_FILE_NAME_PATH          = "Continuous Execution/vft_config.xlsx"
UNITS_SHEET                 = "units_sheet"
DATA_SHEET                  = "data_sheet"
ACCOUNT_SHEET               = "account_sheet"
EMAIL_SHEET                 = "email_sheet"
OUTPUT_FILE                 = "unit_status.docx"
# Account details and fields to be initialized
DATAPLICITY_LOGIN = {}
ACCOUNT_LOGIN               = "to intialize"
LOGS_DISPLAY_PAGE_SIZE      = "to intialize"
LOGS_DISPLAY_PAGE_NUMBER    = "to intialize"
WITHIN_HOURS                = "to intialize"
WITHIN_DAYS                 = "to intialize"
# Email details and authentication
SMTP_SERVER                 = None
SENDER_EMAIL                = None
RECIPIENTS_DAILY            = None
RECIPIENTS_HOURLY           = None
EMAIL_API_KEY               = None
PORT                        = None

# Error message
FAILED_RETRIEVAL = "Likely a server issue. Refresh the unit's logs data page on platform."

### User-input data
def configure_mads(
    file = None
):
    if file == None:
        configure_account_fields(MADS_FILE_NAME)
        configure_data_fields(MADS_FILE_NAME)
        configure_email_fields(MADS_FILE_NAME)
    if os.path.exists(file):
        configure_account_fields(file)
        configure_data_fields(file)
        configure_email_fields(file)


    df = pd.read_excel(io=file, sheet_name=UNITS_SHEET)
    df = df.fillna("")
    df.drop(df.columns[4], axis=1, inplace=True)
    
    units = {}
    for _, row in df.iterrows():
        units[row[0]] = (row[1], row[2], row[3])
    return units

def configure_vft(
    file = None                                              
    ):
    if(file == None):
        configure_account_fields(VFT_FILE_NAME)
        configure_data_fields(VFT_FILE_NAME)
        configure_email_fields(VFT_FILE_NAME)
    if(os.path.exists(file)):
        configure_account_fields(file)
        configure_data_fields(file)
        configure_email_fields(file)


    df = pd.read_excel(io=file, sheet_name=UNITS_SHEET)
    df = df.fillna("")
    df.drop(df.columns[4], axis=1, inplace=True)
    
    units = {}
    for _, row in df.iterrows():
        units[row[0]] = (row[1], row[2], row[3])
    return units

def configure_account_fields(PLATFORM_FILE_NAME):
    global ACCOUNT_LOGIN
    ACCOUNT_LOGIN = {} # dict where keys are email and password
    
    df = pd.read_excel(io=PLATFORM_FILE_NAME, sheet_name=ACCOUNT_SHEET)
    ACCOUNT_LOGIN["email"] = df.iloc[0][1]
    ACCOUNT_LOGIN["password"] = df.iloc[1][1]
    DATAPLICITY_LOGIN["email"] = df.iloc[2][1]
    DATAPLICITY_LOGIN["password"] = df.iloc[3][1]

def configure_data_fields(PLATFORM_FILE_NAME):
    global LOGS_DISPLAY_PAGE_SIZE, LOGS_DISPLAY_PAGE_NUMBER, WITHIN_HOURS, WITHIN_DAYS
    
    df = pd.read_excel(io=PLATFORM_FILE_NAME, sheet_name=DATA_SHEET)
    WITHIN_HOURS = int(df.iloc[0][1])
    WITHIN_DAYS = int(df.iloc[1][1])
    LOGS_DISPLAY_PAGE_SIZE = int(df.iloc[2][1])
    LOGS_DISPLAY_PAGE_NUMBER = int(df.iloc[3][1])
    
def configure_email_fields(PLATFORM_FILE_NAME):
    def parse_recipients_field(recipients):
        lst = recipients.split(',')
        for i in range(len(lst)):
            lst[i] = lst[i].strip()
        return lst
        
    global SMTP_SERVER, SENDER_EMAIL, RECIPIENTS_DAILY, RECIPIENTS_HOURLY, EMAIL_API_KEY
    df = pd.read_excel(io=PLATFORM_FILE_NAME, sheet_name=EMAIL_SHEET)

    SMTP_SERVER = df.iloc[0][1]
    PORT = df.iloc[1][1]
    SENDER_EMAIL = df.iloc[2][1]
    EMAIL_API_KEY = df.iloc[3][1]
    RECIPIENTS_DAILY = parse_recipients_field(df.iloc[5][1])
    RECIPIENTS_HOURLY = parse_recipients_field(df.iloc[4][1])
    

### Logic
# ONLINE    -- Dataplicity online + platform showing logs in the last WITHIN_HOURS
# PARTIAL   -- Dataplicity online + platform showing logs in the last WITHIN_DAYS
# OFFLINE   -- Dataplicity offline or not showing logs from last WITHIN_DAYS
def run_dataplicity_status():
    # get auth token
    url = "https://apps.dataplicity.com/auth/"
    
    rq = requests.post(url, data=DATAPLICITY_LOGIN)
    token = rq.json()["token"]

    endpoint = "https://apps.dataplicity.com/devices/"
    headers = {"Authorization": f"Token {token}"}

    response = requests.get(endpoint, headers=headers)
    json_dump = response.json()
    statuses = {}
    for unit in json_dump:
        # offline by default
        statuses[unit["name"]] = "offline"
        if unit["online"]:
            statuses[unit["name"]] = "online"

    return statuses


def run_mad_status(dataplicity_status, units):
    os.makedirs("data_dump", exist_ok=True)
    
    # key: name, value: (online/offline, loc, remarks)
    status = {}
    count = 1

    # get auth token
    url = "https://datakrewtech.com/api/sign-in"
    rq = requests.post(url, data=ACCOUNT_LOGIN)
    token = rq.json()["access_token"]

    curr_time = get_current_time()
    start_time = get_partial_from(curr_time, WITHIN_DAYS) # consider logs from WITHIN_DAYS ago

    for key, values in units.items():
        # get details
        unit_name = values[0]
        loc = values[1]
        remarks = values[2] # to display if offline or partial

        # Likely a mismatch in unit's name since all units on platform must be linked to dataplicity
        if unit_name not in dataplicity_status:
            print("No matching name on dataplicity for " + unit_name)
            continue
        
        # Unit is offline on dataplicity
        if dataplicity_status[unit_name] == "offline":
            print("Dataplicity indicates " + unit_name + " is offline")
            status[unit_name] = ("offline", loc, remarks)
            continue
        
        endpoint = (
            "https://datakrewtech.com/api/iot_mgmt/orgs/3/projects/70/gateways/"
            + str(key)
            + "/data_dump_index"
        )
        headers = {"Authorization": f"Bearer {token}"}
        params = {
            "page_size": LOGS_DISPLAY_PAGE_SIZE,
            "page_number": LOGS_DISPLAY_PAGE_NUMBER,
            "to_date": curr_time,
            "from_date": start_time,
        }    

        response = requests.get(endpoint, headers=headers, params=params)

        if response.status_code != 200:
            print("Error in fetching " + unit_name + " data for MADs, HTTP status code: ", response.status_code)
            status[unit_name] = ("error", loc, FAILED_RETRIEVAL)
            if response.status_code == 500:
                response = retry_ping(response, unit_name, endpoint, headers, params, 0)
            continue # do not process further

        try:
            json_dump = response.json()
            json_string = json.dumps(json_dump, indent=4)
            filename = "data_dump/data_dump_" + str(count) + ".json"
            with open(filename, "w") as outfile:
                outfile.write(json_string)
                
        except json.JSONDecodeError:
            print(
                "JSONDecodeError for "
                + unit_name
                + ", check if unit_id is entered correctly in config.json"
            )

        count += 1


        start_track = get_online_from(curr_time, WITHIN_HOURS) # must show data WITHIN_HOURS to be considered 'online'
        data_logs = json_dump["data_dumps"]
        
        
        if len(data_logs) > 0:
            timestamp_epoch = data_logs[0]["data"]["timestamp"]

            if timestamp_epoch * 1000 >= start_track:
                status[unit_name] = ("online", loc, "")
                print(unit_name + ": " + "online")
            else:
                status[unit_name] = ("partial", loc, remarks)
                print(unit_name + ": " + "partial data in the last " + str(WITHIN_DAYS) + " days")

            num_entries = json_dump["total_entries"]
            
        else: # no logs found
            remarks = "No logs data in the last " + str(WITHIN_DAYS) + " days" # overwrite
            status[unit_name] = ("offline", loc, remarks)
            print(unit_name + " found but no logs data in the last " + str(WITHIN_DAYS) + " days")

    return status

def retry_ping(response, unit_name, endpoint, headers, params, count):
    if count < 5:
        print("Error probably due to delay in data loading, sleep for 5 seconds.")
        time.sleep(5)
        response = requests.get(endpoint, headers=headers, params=params)
        if response.status_code == 500:
            retry_ping(response, unit_name, endpoint, headers, params, count+1)
        else:
            if response.status_code != 200:
                print("Error in fetching " + unit_name + " data, HTTP status code: ", response.status_code)
            return response
    else:
        print("Error in fetching " + unit_name + " data, HTTP status code: ", response.status_code)
        return response

def run_vft_status(dataplicity_status, units):
    os.makedirs("data_dump", exist_ok=True)

    # key: name, value: online/offline
    status = {}
    count = 1

    # get auth token
    login_url = "https://backend.vflowtechiot.com/api/sign-in"
    login_rq = requests.post(login_url, data=ACCOUNT_LOGIN)
    login_token = login_rq.json()["access_token"]

    url = "https://backend.vflowtechiot.com/api/orgs/3/sign-in"
    org_headers = {"Auth-Token": f"{login_token}"}
    rq = requests.post(url, headers=org_headers)
    token = rq.json()["access_token"]
    

    curr_time = get_current_time()
    start_time = get_partial_from(curr_time, WITHIN_DAYS) # consider logs from WITHIN_DAYS days ago

    for key, values in units.items():
        # get details
        unit_name = values[0]
        loc = values[1]
        remarks = values[2] # to display if offline or partial

        # Likely a mismatch in unit's name since all units on platform must be linked to dataplicity
        if unit_name not in dataplicity_status:
            print("No matching name on dataplicity for " + unit_name)
            continue
        
        # Unit is offline on dataplicity
        if dataplicity_status[unit_name] == "offline":
            print("Dataplicity indicates " + unit_name + " is offline")
            status[unit_name] = ("offline", loc, remarks)
            continue
        
        endpoint = (
            "https://backend.vflowtechiot.com/api/iot_mgmt/orgs/3/projects/70/gateways/"
            + str(key)
            + "/data_dump_index"
        )
        headers = {"Authorization": f"Bearer {token}"}
        params = {
            "page_size": LOGS_DISPLAY_PAGE_SIZE,
            "page_number": LOGS_DISPLAY_PAGE_NUMBER,
            "to_date": curr_time,
            "from_date": start_time,
        }

        response = requests.get(endpoint, headers=headers, params=params)

        if response.status_code != 200:
            print("Error in fetching " + unit_name + " data for VFT, HTTP status code: ", response.status_code)
            status[unit_name] = ("error", loc, FAILED_RETRIEVAL)
            continue # do not process further

        try:
            json_dump = response.json()
            json_string = json.dumps(json_dump, indent=4)
            filename = "data_dump/data_dump_" + str(count) + ".json"
            with open(filename, "w") as outfile:
                outfile.write(json_string)
                
        except json.JSONDecodeError:
            print(
                "JSONDecodeError for "
                + unit_name
                + ", check if unit_id is entered correctly in config.json"
            )

        count += 1
       
        start_track = get_online_from(curr_time, WITHIN_HOURS) # must show data WITHIN_HOURS to be considered 'online'
        data_logs = json_dump["data_dumps"]
    
        if len(data_logs) > 0:
            timestamp_epoch = data_logs[0]["data"]["timestamp"]

            if timestamp_epoch * 1000 >= start_track:
                status[unit_name] = ("online", loc, "")
                print(unit_name + ": " + "online")
            else:
                status[unit_name] = ("partial", loc, remarks)
                print(unit_name + ": " + "partial data in the last " + str(WITHIN_DAYS) + " days")

            num_entries = json_dump["total_entries"]
            
        else: # no logs found
            remarks = "No logs data in the last " + str(WITHIN_DAYS) + " days" # overwrite
            status[unit_name] = ("offline", loc, remarks)
            print(unit_name + " found but no logs data in the last " + str(WITHIN_DAYS) + " days")
    return status
        

### Utils
def _set_cell_background(cell, fill, color = None, val = None):
    """
    @fill: Specifies the color to be used for the background
    @color: Specifies the color to be used for any foreground
    pattern specified with the val attribute
    @val: Specifies the pattern to be used to lay the pattern
    color over the background color.
    """
    cell_properties = cell._element.tcPr
    if cell_properties.xpath("w:shd"): # exists existing shading
        cell_shading = cell_properties.xpath("w:shd")[0]
    else: # add new w:shd element
        cell_shading = OxmlElement("w:shd")

    if fill:
        cell_shading.set(qn("w:fill"), fill)
    if color:
        pass #TODO
    if val:
        pass #TODO

    #extend cell props with shading
    cell_properties.append(cell_shading)


def get_current_time():
    return int(time.time() * 1000)

def get_partial_from(curr_time, WITHIN_DAYS):
    return curr_time - 60 * 60 * 24 * WITHIN_DAYS * 1000

def get_online_from(curr_time, WITHIN_HOURS):
    return curr_time - 60 * 60 * WITHIN_HOURS * 1000

def remove_data_dump():
    folder = 'data_dump/'
    for filename in os.listdir(folder):
        if filename != "status.json":
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print('Failed to delete %s. Reason: %s' % (file_path, e))

def get_greetings(time):
    hour = int(time.split(":")[0])
    if hour >= 4 and hour < 12:
        return "Good Morning,"
    elif hour >= 12 and hour < 18:
        return "Good Afternoon,"
    elif hour >= 18 and hour <= 23:
        return "Good Evening,"
    else: 
        return "Hi,"
### Main Code
def generate_report(mads = False):
    statusDict = {}
    rtn = []
    document = Document()
    document.add_heading("Unit Status", 0)
    table = document.add_table(rows=1, cols=5)
    heading_cells = table.rows[0].cells
    heading_cells[0].text = "Unit Name"
    if mads:
        heading_cells[1].text = "On MADs"
    else:
        heading_cells[1].text = "On VFT"
    print("Generating report for", heading_cells[1].text)
    rtn.append(heading_cells[1].text)
    heading_cells[2].text = "On Dataplicity"
    heading_cells[3].text = "Location"                     
    heading_cells[4].text = "Remarks"

    # because there's a dependency between dataplicity status and platform status
    # if dataplicity offline --> unit is offline regardless of logs shown
    dataplicity_status = None # handled below
    
    if mads:
        tracked_units = configure_mads()
        dataplicity_status = run_dataplicity_status()
        platform_status = run_mad_status(dataplicity_status, tracked_units)
    else:
        tracked_units = configure_vft()
        dataplicity_status = run_dataplicity_status()
        platform_status = run_vft_status(dataplicity_status, tracked_units)            

    for unit, values in platform_status.items():
        unit_status_platform = values[0]
        unit_status_dataplicity = dataplicity_status[unit]
        location = values[1]
        remark = values[2]
        cells = table.add_row().cells
        if unit_status_platform == "online":
            _set_cell_background(cells[1], "CEEDD0")
        elif unit_status_platform == "offline":
            _set_cell_background(cells[1], fill="F6CACF")
        elif unit_status_platform == "partial":
            _set_cell_background(cells[1], fill="FBEBA6")
        elif unit_status_platform == "error":
            _set_cell_background(cells[1], fill="FF0000")

        if unit_status_dataplicity == "online":
            _set_cell_background(cells[2], "CEEDD0")
        elif unit_status_dataplicity == "offline":
            _set_cell_background(cells[2], fill="F6CACF")
        elif unit_status_platform == "partial":
            _set_cell_background(cells[2], fill="FBEBA6")
        cells[0].text = unit
        cells[1].text = unit_status_platform
        cells[2].text = unit_status_dataplicity
        cells[3].text = location
        cells[4].text = remark
        # print(unit, values)
        statusDict[unit] = values[0]
    
    rtn.append(statusDict)
    table.style = "Table Grid"

    document.save(OUTPUT_FILE)
    remove_data_dump()
    return rtn

##############################################
########## Duplicate file for report generator

def generate_report(
        mads        = False,                                        # By default, MADs is not triggered
        vft         = True,                                         # Instead, VFT is triggered 
        path        = None                                          # Define the file to be parsed in (path declared)                                      
    ):
##### SECTION 1 
##### CHECK FILE AVAILABILITY
    if(path == None):                                               # If no file was parsed in as params
        print("Error: No excel file was read in the execution")     # Print error message and
        print("Configuration use default file")
        path = VFT_FILE_NAME                                        # Default path
    
    if(not os.path.exists(path)):                                   # If the directory does not exists
        print("Error: The directory/file requested does not exist") # Print error message and
        #return 1                                                    # Return ID Error of 1
    
    if(not mads and not vft):                                       # If neither of the state is triggered
        print("Error: MADs mode or VFT mode are neither used")      # Print error message
        #return 2                                                    # Return ID Error of 2

##### SECTION 2
##### GENERATE FILE
    statusDict = {}                                                 # Dictionary of status
    rtn = []                                                        
    document = Document()                                           # Create a new document (docx)
    document.add_heading("Unit Status", 0)                          # Add heading for the docx
    table = document.add_table(rows=1, cols=5)                      # Add table for the docx 

    heading_cells = table.rows[0].cells                             # Populate Title for each column
    heading_cells[0].text = "Unit Name"                             # Unit name
    if mads:                                                        # On which platform
        heading_cells[1].text = "On MADs"
    if vft:
        heading_cells[1].text = "On VFT"
    print("Generating report for", heading_cells[1].text)
    rtn.append(heading_cells[1].text)
    heading_cells[2].text = "On Dataplicity"                        # On Dataplicity
    heading_cells[3].text = "Location"                              # Location of Unit
    heading_cells[4].text = "Remarks"                               # Notes about the unit

    # because there's a dependency between dataplicity status and platform status
    # if dataplicity offline --> unit is offline regardless of logs shown
    dataplicity_status = None # handled below
    
    if mads:
        tracked_units      = configure_mads(path)
        dataplicity_status = run_dataplicity_status()
        platform_status    = run_mad_status(dataplicity_status, tracked_units)

    if vft:
        tracked_units      = configure_vft(path)
        dataplicity_status = run_dataplicity_status()
        platform_status    = run_vft_status(dataplicity_status, tracked_units)            

    for unit, values in platform_status.items():
        unit_status_platform = values[0]
        unit_status_dataplicity = dataplicity_status[unit]
        location = values[1]
        remark = values[2]
        cells = table.add_row().cells
        if unit_status_platform == "online":
            _set_cell_background(cells[1], "CEEDD0")
        elif unit_status_platform == "offline":
            _set_cell_background(cells[1], fill="F6CACF")
        elif unit_status_platform == "partial":
            _set_cell_background(cells[1], fill="FBEBA6")
        elif unit_status_platform == "error":
            _set_cell_background(cells[1], fill="FF0000")

        if unit_status_dataplicity == "online":
            _set_cell_background(cells[2], "CEEDD0")
        elif unit_status_dataplicity == "offline":
            _set_cell_background(cells[2], fill="F6CACF")
        elif unit_status_platform == "partial":
            _set_cell_background(cells[2], fill="FBEBA6")
        cells[0].text = unit
        cells[1].text = unit_status_platform
        cells[2].text = unit_status_dataplicity
        cells[3].text = location
        cells[4].text = remark
        # print(unit, values)
        statusDict[unit] = values[0]
    
    rtn.append(statusDict)
    table.style = "Table Grid"

    document.save(OUTPUT_FILE)
    remove_data_dump()
    return rtn






def checkUnitStatus(system, state):
    jsonFile = open('data_dump/status.json')
    unitPrevStatus = json.load(jsonFile)
    offlineUnits = []
    for i in unitPrevStatus:
        try:
            if unitPrevStatus[i] != state[i] and state[i] == "offline":
                offlineUnits.append(i)
        except:
            print("Error Occured")
    if offlineUnits != []:
        sendHourlyEmail(system, offlineUnits)

def StoreStatus(statusDict):
    json_string = json.dumps(statusDict, indent=4)
    filename = "data_dump/status.json"
    with open(filename, "w") as outfile:
        outfile.write(json_string)

### Craft and send email
def sendDailyEmail(System):
    port = PORT # For SSL

    formatted_time = datetime.datetime.now().strftime("%H:%M") # get current time
    
    message = MIMEMultipart()
    message['Subject'] = 'Daily Status Report'
    message['From'] = SENDER_EMAIL
    message['To'] = ", ".join(RECIPIENTS_DAILY)

    # add message body
    body = get_greetings(formatted_time) + "\n\n"
    body += "This Is An Automated Status Report for \"" + formatted_time + "\"\n"
    body += "The Status Is Collected " + System + "\n\n"
    body += "Regards," + "\n"
    body += "OfficeUnit-CT3" + "\n"
    message.attach(MIMEText(body, 'plain'))

    with open(OUTPUT_FILE, 'rb') as attachment:
        attachment = MIMEApplication(attachment.read(), _subtype='docx')
        attachment.add_header('content-disposition', 'attachment', filename=OUTPUT_FILE)
        message.attach(attachment)
    
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_SERVER, port, context=context) as server:
        server.login(SENDER_EMAIL, EMAIL_API_KEY)
        server.sendmail(SENDER_EMAIL, RECIPIENTS_DAILY, message.as_string())

def sendHourlyEmail(System, AlertsDevice=[]):
    port = PORT # For SSL

    formatted_time = datetime.datetime.now().strftime("%H:%M") # get current time
    
    message = MIMEMultipart()
    message['Subject'] = '(Alert) Hourly Status Report'
    message['From'] = SENDER_EMAIL
    message['To'] = ", ".join(RECIPIENTS_HOURLY)

    # add message body
    body = get_greetings(formatted_time) + "\n\n"
    body += "This is an alert report! Current time: " + formatted_time + "\n"
    body += "The Following Device(s) have their status changed to offline.\n"
    body += "---------------------------------\n"
    for i in AlertsDevice:
        body += i + "\n"
    body += "---------------------------------\n"
    body += "\n"
    body += "Regards," + "\n"
    body += "OfficeUnit-CT3" + "\n"
    message.attach(MIMEText(body, 'plain'))

    with open(OUTPUT_FILE, 'rb') as attachment:
        attachment = MIMEApplication(attachment.read(), _subtype='docx')
        attachment.add_header('content-disposition', 'attachment', filename=OUTPUT_FILE)
        message.attach(attachment)
    
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_SERVER, port, context=context) as server:
        server.login(SENDER_EMAIL, EMAIL_API_KEY)
        server.sendmail(SENDER_EMAIL, RECIPIENTS_HOURLY, message.as_string())

if __name__ == "__main__":
    print("Starting Script...")
    isMADs = False
    isVFT  = True
    rtn = generate_report(isMADs, isVFT, path = VFT_FILE_NAME_PATH)
    StoreStatus(rtn[1])
    validSend = True
    while(1):
        currTime = datetime.datetime.now().strftime("%H:%M")
        mins = currTime.split(":")[1]
        if (currTime == "14:00" or currTime == "14:01") and validSend:
            print("==============================================")
            print("Starting 1400 Status Check.")
            print("==============================================")
            rtn = generate_report(isMADs, isVFT, path = VFT_FILE_NAME_PATH)               #Generate report
            sendDailyEmail(rtn[0])                      #Send Daily Email
            checkUnitStatus(rtn[0], rtn[1]);            #Check Status
            StoreStatus(rtn[1])                         #Overwrite Status
            validSend = False
            print("==============================================")
        elif (mins == "30" or mins == "31") and not validSend:
            print("==============================================")
            print("Resetting Checks.")
            validSend = True
            print("==============================================")
        elif (mins == "00" or min =="01") and validSend:
            print("==============================================")
            print("Starting Hourly Check for", currTime, ".")
            print("==============================================")
            rtn = generate_report(isMADs, isVFT, path = VFT_FILE_NAME_PATH)               #Generate report
            checkUnitStatus(rtn[0], rtn[1]);            #Check Status
            StoreStatus(rtn[1])                         #Overwrite Status
            validSend = False
            print("==============================================")
