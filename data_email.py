import win32com.client as client
import email
import imaplib
import re

import os
from datetime import datetime, timedelta
import mailparser
import csv

outlook = client.Dispatch("Outlook.Application").GetNameSpace("MAPI")

account = outlook.Folders['egutierrez-ntt@huntservices.pe']

inbox = account.Folders['Bandeja de entrada']

duo = inbox.Folders['Duo']

# yt_emails = [message for message in inbox.Items if message.Subject.startswith('[TICK')]

# for message in yt_emails:
#     print(message)

# yt_folder = inbox.folders.add('Tickets')

# for message in yt_emails:
#     message.Move(yt_folder)

# junk_messages = [message for message in duo.Items if 'authentication failures' in message.Body.lower()]

# print(len(junk_messages))

# for message in junk_messages:
#     print(message.SenderEmailAddress)

yt_duo = [message for message in inbox.Items if 'PI Notifications' in message.Subject]

ls = list()
ls2 = list()
name = ''
trigger = ''

for message in yt_duo:
    text = message.Body
    start = message.Body.find('Name: ') + 6
    end = message.Body.find(' Pwr', start)
    name = text[start:end]
    print(text[start:end])

    text2 = message.Body.strip()
    start2 = text2.find('Trigger Time:') + 14
    end2 = text2.find('Path', start2)
    trigger = text2[start2:end2]
    print(text2[start2:end2])






# with open('search_resuts.csv', 'w', newline='', encoding='utf-8') as f:
#     writer = csv.writer(f)
#     writer.writerow(['Name'],['Trigger'])
#     for message in ls:
#         writer.writerow([message])
