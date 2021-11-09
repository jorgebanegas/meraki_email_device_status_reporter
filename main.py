""" Copyright (c) 2021 Cisco and/or its affiliates.
This software is licensed to you under the terms of the Cisco Sample
Code License, Version 1.1 (the "License"). You may obtain a copy of the
License at
           https://developer.cisco.com/docs/licenses
All use of the material herein must be in accordance with the terms of
the License. All rights not expressly granted by the License are
reserved. Unless required by applicable law or agreed to separately in
writing, software distributed under the License is distributed on an "AS
IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
or implied. 
"""

import smtplib
from meraki import config
from texttable import Texttable
import meraki
import config
import requests
from tabulate import tabulate
import pandas as pd
from pretty_html_table import build_table
import datetime
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders

dashboard = meraki.DashboardAPI(api_key=config.meraki_key,output_log=False)

url = f"https://api.meraki.com/api/v1/organizations/{config.org_id}/devices/statuses"

payload = {}
headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "X-Cisco-Meraki-API-Key": f"{config.meraki_key}"
    }

devices = requests.request('GET', url, headers=headers, data=payload).json()

for device in devices:
    if device['lastReportedAt'] == None:
        device['lastReportedAt'] = ''

devices = sorted(devices, key=lambda d: d['lastReportedAt'], reverse=True)

temp_list = []
temp_device = []
device_link_list = []


# (downtime,usage)
networks = dashboard.organizations.getOrganizationNetworks(organizationId=config.org_id)

organization = dashboard.organizations.getOrganization(organizationId=config.org_id)

org_url = '<a href="' + organization['url'] + '">Dashboard Link</a><br>'

device_base_url = 'https://api.meraki.com/manage/dashboard/show?mac='

online_count = len(device)
offline_count = 0

for device in devices:
    temp_link = {}
    try:
        clients = dashboard.devices.getDeviceClients(device["serial"])
        client_count = len(clients)
    except:
        client_count = "N/A"

    try:
        resp = dashboard.devices.getDeviceLldpCdp(device["serial"])
        switch_port = str(resp['ports']['wired0']['lldp']['systemName'].split(" - ")[1]) + "/" + str(resp['ports']['wired0']['lldp']['portId'])
    except:
        switch_port = "N/A"
    try:
        date = ""
        date = date + device["lastReportedAt"].split("T")[0] + " "
        time = device["lastReportedAt"].split("T")[1].split(".")[0]
        time = datetime.datetime.strptime(time,'%H:%M:%S').strftime('%I:%M %p')
    except:
        time = ""
        date = ""

    for network in networks:
        if device["networkId"] == network["id"]:
            network_name = network["name"]
    if device["status"] != "online":
        offline_count = offline_count + 1
        device_link_list.append((device["serial"],device_base_url+device["mac"]))
        device_detail = dashboard.devices.getDevice(device["serial"])
        try:
            note = device_detail["notes"]
        except:
            note = " "
        temp_list.append([device["name"],network_name,note,device["model"],device["serial"], device["mac"],client_count,device["status"],switch_port,date+time,"placeholder"])
        
data = pd.DataFrame(temp_list,columns=["name","networkName","notes","model","serial", "mac","clients","status","switch/port","lastReportedAt","downtime(mins)"])

for device in device_link_list:
    data["serial"].replace({device[0]: '<a href="{}">{}</a>'.format(device[1], device[0])}, inplace=True)

html = build_table(data, 'blue_light', escape=False)

html = org_url + html 

# Add HTML/plain-text parts to MIMEMultipart message
# The email client will try to render the last part first
message = MIMEMultipart("alternative")

img1 = 'images/orange_dot.png'
img2 = 'images/green_dot.png'
img3 = 'images/red_dot.png'

if online_count == offline_count:
    chosen_image = img2
elif online_count - offline_count == 1:
    chosen_image = img1
else:
    chosen_image = img3

with open(chosen_image, 'rb') as f:
    # set attachment mime and file name, the image type is png
    mime = MIMEBase('image', 'png', filename='img1.png')
    # add required header data:
    mime.add_header('Content-Disposition', 'attachment', filename='img1.png')
    mime.add_header('X-Attachment-Id', '0')
    mime.add_header('Content-ID', '<0>')
    # read attachment file content into the MIMEBase object
    mime.set_payload(f.read())
    # encode with base64
    encoders.encode_base64(mime)
    # add MIMEBase object to MIMEMultipart object
    message.attach(mime)

html = '<p><img src="cid:0">' +  str(offline_count) + '/' + str(online_count) + ' APs online</p>' + html 

# Turn these into plain/html MIMEText objects
#part1 = MIMEText(org_url, "html")
part2 = MIMEText(html, "html")


message["Subject"] = "Meraki Device Report"
message["From"] = config.email_username
message["To"] = config.receiver_email

#message.attach(part1)
message.attach(part2)

text = MIMEText('<img src="cid:image1">', 'html')
message.attach(text)

email_user = config.email_username
email_password = config.email_password

sent_from = email_user
to = config.receiver_email


try:
    smtp_server = smtplib.SMTP_SSL(config.smtp_domain, 465)
    smtp_server.ehlo()
    smtp_server.login(config.email_username, config.email_password)
    smtp_server.sendmail(sent_from, to, message.as_string())
    smtp_server.close()
    print ("Email sent successfully!")
except Exception as ex:
    print ("Something went wrong….",ex)