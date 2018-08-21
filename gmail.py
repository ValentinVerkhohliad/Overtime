#python3.5
from __future__ import print_function
import httplib2
import os
import threading

import base64
from apiclient import errors
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
errors = ''

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Python Quickstart'


def get_credentials():
    global errors
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        errors = errors + ('Storing credentials to ' + credential_path + '\n')
    return credentials


def ListMessagesWithLabels(service, user_id, label_ids=[]):
    global errors
    try:
        response = service.users().messages().list(userId=user_id,
                                                   labelIds=label_ids).execute()
        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])
        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId=user_id,
                                                     labelIds=label_ids,
                                                     pageToken=page_token).execute()
            messages.extend(response)
        return messages
    except errors.HttpError as error:
        errors = errors + ('An error occurred: %s' % error + '\n')


def GetAttachments(month, breakpoint, service, user_id, msg_id, prefix="C:\\reports\\logins\\"):
    global errors
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()
    for part in message['payload']['parts']:
        if part['filename']:
            if 'data' in part['body']:
                data = part['body']['data']
            else:
                att_id = part['body']['attachmentId']
                att = gmail_service.users().messages().attachments().get(userId=user_id, messageId=msg_id,id=att_id).execute()
                data = att['data']
            if month in part['filename']:
                file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                path = prefix+part['filename']
                with open(path, 'wb') as f:
                    f.write(file_data)
            if breakpoint in part['filename']:
                raise ZeroDivisionError


credentials = get_credentials()
http = credentials.authorize(httplib2.Http())
gmail_service = discovery.build('gmail', 'v1', http=http)
labels = gmail_service.users().labels().list(userId='me').execute()


def execute_(month):
    global errors
    try:
        if int(month.lstrip('.')) == 1:
            breakpoint = '.12'
        elif int(month.lstrip('.')) <= 10:
            breakpoint = '.0' + str((int(month.lstrip('.')) - 1))
        else:
            breakpoint = '.' + str((int(month.lstrip('.')) - 1))
        for label in labels['labels']:
            if label['name'] == 'py-cw1':
                py_cw1_label = label['id']
        messages = ListMessagesWithLabels(gmail_service, 'me', py_cw1_label)
        for msg in messages:
            m_id = msg['id']
            try:
                GetAttachments(month, breakpoint, gmail_service, 'me', m_id)
            except KeyError as e:
                if 'parts' in str(e):
                    errors = errors + ("No attachment in message\n ", m_id + '\n')
                else:
                    errors = errors + ("Something wrong in message\n ", m_id, e + '\n')
            except ZeroDivisionError:
                errors = errors + 'All is done\n'
                break
    except ValueError:
        errors = errors + 'Wrong month number\n'
    except PermissionError:
        errors = errors + 'Close all excel Aplications, ans Processes\n'
