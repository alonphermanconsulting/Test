from __future__ import print_function
from oauth2client.service_account import ServiceAccountCredentials
from xml.dom.minidom import parseString
import xmlrpclib
import os
import base64
import ConfigParser
import oauth2client
import httplib2
import pprint
import json
import re

from zipfile import ZipFile
from apiclient import discovery
from apiclient import errors
from oauth2client import client
from oauth2client import tools

try:
    import argparse

    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
#SCOPES = 'https://mail.google.com/'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Python Quickstart'
EMAIL_MESSAGE_DUMP_FILE = "emails.txt"

def findJsonValue(somejson, key):
    def val(node):
        # Searches for the next Element Node containing Value
        e = node.nextSibling
        while e and e.nodeType != e.ELEMENT_NODE:
            e = e.nextSibling
        return (e.getElementsByTagName('string')[0].firstChild.nodeValue if e
                else None)
    # parse the JSON as XML
    #foo_dom = parseString(somejson)
    foo_dom = parseString(xmlrpclib.dumps(somejson))
    # and then search all the name tags which are P1's
    # and use the val user function to get the value
    return [val(node) for node in foo_dom.getElementsByTagName('name')
            if node.firstChild.nodeValue in key]

def main():

    config = ConfigParser.RawConfigParser()
    config.read('config.properties')
    attachment_directory = config.get('general', 'attachment_directory')
    # if 'attachments' not in os.listdir(attachment_directory):
    #     os.mkdir(attachment_directory + r'\attachments')

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)

    # messages = ListMessagesWithLabel(service, 'me', 'INBOX')
    # if not messages:
    #     print('No messages found.')
    # else:
    #     for message in messages:
    #         message_details = service.users().messages().get(userId='me', id=message['id']).execute()

    #with open(EMAIL_MESSAGE_DUMP_FILE, 'wt') as out:
     #   res = json.dump(message_details, out, sort_keys=True, indent=4, separators=(',', ': '))

        # TOCHANGE
        # emailLines = open(EMAIL_MESSAGE_DUMP_FILE, "r")

    # find messageid
    emailLines = open("scratchpad.txt", "r")
    for line in emailLines:
        if re.match("(.*)id'(.*)", line):
            strippedLine = line.replace(' ', '').replace('\'id\':', '').replace(':uu', '').replace('\'', ''). \
                replace(',', '').replace('\n', '').replace('uu', '')
            messageId = strippedLine

    emailLines = open("scratchpad.txt", "r")

    fileName = ''
    attachmentId = ''
    for line in emailLines:
        if re.match("(.*)attachmentId(.*)", line):
            strippedLine = line.replace(' ', '').replace('\'attachmentId\':', '').replace('{', ''). \
                replace(',', '').replace('\n', '').replace('\'', '').replace('ubody', '').replace(':uu', '')
            attachmentId = strippedLine
        if re.match("(.*)filename(.*)", line) and not re.match("(.*)value(.*)", line) and not re.match("(.*)''(.*)",line):
            strippedLine = line.replace(' ', '').replace('\'filename\':', ''). \
                replace(',', '').replace('\n', '').replace('\'', '').replace('uu', '')
            fileName = strippedLine
        if fileName and attachmentId:
            att = service.users().messages().attachments().get(userId='me', messageId=messageId,id=attachmentId).execute()
            data = att['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            path = os.path.join(attachment_directory, fileName)
            if not os.path.isfile(path):
                with open(path, 'w') as f:
                    f.write(file_data)
                    # if path.endswith(".zip"):
                    #     zippedFile = ZipFile(path)
                    #     zippedFile.extractall(store_dir)
                    #     os.remove(path)
                    continue

          #  test = findJsonValue(message_details, 'attachmentId')
          #  pp = pprint.PrettyPrinter(indent=4)
          #  pp.pprint(message_details)
          #  SaveAttachments(service=service, user_id='me', store_dir=attachment_directory, msg_id=message['id'])


def get_credentials():
    # type: () -> object
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """

    # home_dir = os.path.expanduser('~')
    # credential_dir = os.path.join(home_dir,
    #                                'Google Drive/G/IT/Development/ETL/.credentials')  # JLD - customized path to working directory
    # credential_path = os.path.join(credential_dir, 'client_secret.json')
    # credentials = ServiceAccountCredentials.from_json_keyfile_name(credential_path, scopes=SCOPES)


    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir,
                                  'Google Drive/G/IT/Development/ETL/.credentials')  # JLD - customized path to working directory
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'client_secret_anchorpath.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run_flow(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def ListMessagesWithLabel(service, user_id, label_name):
    """List all Messages of the user's mailbox with label_ids applied.

    Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        label_ids: Only return Messages with these labelIds applied.

    Returns:
        List of Messages that have all required Labels applied. Note that the
        returned list contains Message IDs, you must use get with the
        appropriate id to get the details of a Message.
    """

    try:
        response = service.users().messages().list(userId=user_id, labelIds=label_name).execute()
        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId=user_id, pageToken=page_token).execute()
            messages.extend(response['messages'])

        return messages
    except errors.HttpError, error:
        print('An error occurred: ', error)


def SaveAttachments(service, user_id, msg_id, store_dir):
    """Get and store attachment from Message with given id.

    Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        msg_id: ID of Message containing attachment.
        store_dir: The directory used to store attachments.
    """
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()
        if message['payload']:
            # for header in message['payload']['headers']:
            #     if header['name'] == 'Subject':
            #         subject = header['value'].replace(':','').replace('/','').replace('"','')
            #         break
            if 'parts' in message['payload']:
                partsparent = message['payload']
                parts = partsparent['parts']
                for part in partsparent[parts]:
                    if part['filename']:
                        file_data = base64.urlsafe_b64decode(part['body']['data'].encode('UTF-8'))
                        path = ''.join([store_dir, part['filename']])
                        f = open(path, 'w')
                        f.write(file_data)
                        f.close()
                    if 'attachmentId' in part['body']:
                        att_id=part['body']['attachmentId']
                        att=service.users().messages().attachments().get(userId=user_id, messageId=msg_id,id=att_id).execute()
                        data=att['data']
                        file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                        path = os.path.join(store_dir, part['filename'])
                        if not os.path.isfile(path):
                            with open(path, 'w') as f:
                                f.write(file_data)
                            # if path.endswith(".zip"):
                            #     zippedFile = ZipFile(path)
                            #     zippedFile.extractall(store_dir)
                            #     os.remove(path)

    except errors.HttpError, error:
        print('An error occurred: ', error)

if __name__ == '__main__':
    main()
