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
import zipfile

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
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Python Quickstart'
EMAIL_MESSAGE_DUMP_FILE = "emails.txt"

def main():

    config = ConfigParser.RawConfigParser()
    config.read('config.properties')
    attachment_directory = config.get('general', 'attachment_directory')
    suffixes = config.get('general', 'emails_suffixes_to_ignore').split(",")
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)
    messages = ListMessagesWithLabel(service, 'me', 'INBOX')
    label = config.get('general', 'label')
    if not messages:
        print('No messages found.')
    else:
        for message in messages:
            message_details = service.users().messages().get(userId='me', id=message['id']).execute()
            messageString = json.dumps(message_details).replace('\\','')
            #print('message string', messageString)
            pp = pprint.PrettyPrinter(indent=4)
            #pp.pprint(message_details)

            attachmentIdMatch = re.match(r"(.*)attachmentId\": \"(.+?)\"(.*)", messageString)
            fromMatch = re.match(r"(.*)From\", \"value\": \"(.+?)\"(.*)", messageString)
            subjectMatch = re.match(r"(.*)Subject\", \"value\": \"(.+?)\"(.*)", messageString)
            returnPathMatch = re.match(r"(.*)Return-Path\", \"value\": \"(.+?)\"(.*)", messageString)

            if label not in messageString:
                print('not correct label - ignoring')
                continue

            if not returnPathMatch:
                print('WARNING no return path - ignoring')
                pp.pprint(message_details)
                continue

            if not subjectMatch:
                print('WARNING no subject - ignoring')
                pp.pprint(message_details)
                continue

            fromFound = fromMatch.group(2)
            subjectFound = subjectMatch.group(2)
            returnPathFound = returnPathMatch.group(2).replace('<', '').replace('>', '')
            if '@' not in returnPathFound:
                print('skipping message as no valid return path', messageSummary)
                continue
            messageSummary = "FROM " + fromFound + " RETURN PATH " + returnPathFound + " SUBJECT " +subjectFound
            emailSuffixToCheck = re.match(r"(.*)@(.*)", returnPathFound).group(2)
            if emailSuffixToCheck in suffixes:
                print('skipping message as email suffix is to be ignored', messageSummary)
                continue
            fileNameMatch = re.match(r"(.*)filename=\"(.+?)\"(.*)", messageString)
            if not fileNameMatch:
                print('skipping message as no attachment present', messageSummary)
                continue
            fileNameFound = fileNameMatch.group(2)
            attachmentIdFound = attachmentIdMatch.group(2)
            att = service.users().messages().attachments().get(userId='me', messageId=message['id'],id=attachmentIdFound).execute()

            data = att['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            path = os.path.join(attachment_directory, fileNameFound)
         # PUT BACK
            if not os.path.isfile(path) and not path.lower().endswith(".jpg") and not path.lower().endswith(".png") \
                    and not path.lower().endswith(".gif") and not path.lower().endswith(".ics"):
         #   if path.endswith(".zip"):
                with open(path, 'w') as f:
                    f.write(file_data)
                if path.endswith(".zip"):
                    try:
                        fh = open(path, 'rb')
                        z = zipfile.ZipFile(fh)
                        #zippedFile = ZipFile(path)
                        z.extractall(attachment_directory)
                        os.remove(path)
                        fh.close()
                    except zipfile.BadZipfile, error:
                        print('unzip error occurred: ', error)
            continue


def get_credentials():
    # type: () -> object
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """

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

if __name__ == '__main__':
    main()
