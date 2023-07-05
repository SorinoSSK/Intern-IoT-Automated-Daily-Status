import requests
import time
import json
from os.path import exists
import yagmail
import argparse

def send_email_gmail(to, doc):
    yag = yagmail.SMTP('vft.test123@gmail.com', 'wvwdrfcwsaepgjhe')
    subject = "Status Report"
    body = "<h1>Unit Status Report</h1>"
    attachments = doc

    yag.send(to=to, subject=subject, contents=body,
             attachments=attachments)

def get_mads():
    if not exists("tracking.json"):
        data = {}
        data['units'] = {'unit_id': 'unit_name'}
        data['recipients'] = ['email']

        json_string = json.dumps(data, indent=4)

        with open('tracking.json', 'w') as outfile:
            outfile.write(json_string)
        exit("enter unit_id, unit_name and recipients in tracking.json")
    else:
        with open('tracking.json') as json_file:
            file = json.load(json_file)
            units = file['units']
            to = file['recipients']

    status = {}

    # get auth token
    url = "https://datakrewtech.com/api/sign-in"
    myobj = {'email': 'stevenedbert47@gmail.com', 'password': 'vflow123'}
    x = requests.post(url, data=myobj)

    token = x.json()['access_token']

    curr_time = int(time.time() * 1000)
    start_time = curr_time - 1800000

    for key, value in units.items():
        endpoint = 'https://datakrewtech.com/api/iot_mgmt/orgs/3/projects/70/gateways/' + \
            str(key) + '/data_dump_index'
        headers = {'Authorization': f'Bearer {token}'}
        params = {'page_size': 1000, 'page_number': 1,
                  'to_date': curr_time, 'from_date': start_time}

        response = requests.get(endpoint, headers=headers, params=params)
        try:
            json_dump = response.json()
            # print(len(json_dump['data_dumps']))
        except json.JSONDecodeError:
            print('JSONDecodeError for ' + units[key])

        try:
            timestamp_epoch = json_dump['data_dumps'][0]['data']['timestamp']
            status[units[key]] = 'online'
        except IndexError:
            # json_dump['data_dumps'] is empty
            status[units[key]] = 'offline'

    return status, to

def send_email_gmail(to, status):
    yag = yagmail.SMTP('vft.test123@gmail.com', 'wvwdrfcwsaepgjhe')
    subject = "Status Report"
    body = "<h1>Unit Status Report</h1>"
    allOk = True
    for key, value in status.items():
        if value != 'online':
            allOk = False
            body += f'<p>{key}: {value}</p>'
    if allOk:
        body += '<p>All units are online</p>'
    # attachments = doc
    print(body)

    yag.send(to=to, subject=subject, contents=body)

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--interval', type=int, default=30, help='interval in minutes')
    args = parser.parse_args()
    

    starttime = time.time()
    while True:
        status, to = get_mads()
        allOk = True
        for key,value in status.items():
            if value == 'offline':
                send_email_gmail(to, status)
                allOk = False
                break
        if allOk:
            print('All units are online')
        else:
            print('Some units are offline')
        time.sleep((60.0 * args.interval) - ((time.time() - starttime) % (60.0 * args.interval)))