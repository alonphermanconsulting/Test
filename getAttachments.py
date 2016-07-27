from __future__ import print_function
import httplib2
import os
import base64
import oauth2client

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


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('.')
    credential_dir =  os.path.join(home_dir,'')  #,'Google Drive/G/IT/Development/ETL/.credentials') #JLD - customized path to working directory
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'client_secret.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def GetAttachments(service, user_id, msg_id, store_dir):
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

        for part in message['payload']['parts']:
            if part['filename']:
            
                print(part['filename'])
                att_id = part['body']['attachmentId']
                #print(att_id)
                raw = service.users().messages().attachments().get(userId=user_id, messageId=id, id=att_id).execute()
                file_data = base64.urlsafe_b64decode(raw['data'].encode('UTF-8'))

                path =  os.path.join(store_dir, part['filename'])
                print(path)
                f = open(path, 'w')
                f.write(file_data)
                f.close()

    except errors.HttpError, error:
        print('An error occurred: ',error)

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
        response = service.users().messages().list(userId=user_id, q='label:'+label_name).execute()
        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId=user_id, q='label:'+label_name, pageToken=page_token).execute()
            messages.extend(response['messages'])

        return messages
    except errors.HttpError, error:
        print('An error occurred: ', error)

def ListLabels(service, user_id):
    """Get a list all labels in the user's mailbox.

    Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.

        Returns:
    A list all Labels in the user's mailbox.
    """
    
    try:
        response = service.users().labels().list(userId=user_id).execute()
        labels = response['labels']
        for label in labels:
            #print( "Label id: {1} - Label name: {2}".format(label['id'], label['name']) )  
            print(label)
        return labels
    except errors.HttpError, error:
        print('An error occurred: ', error) 
        
def GetMessage(service, user_id, msg_id):
    """Get a Message with given ID.

    Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        msg_id: The ID of the Message required.

    Returns:
        A Message.
    """
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()

        print('Message snippet: ', message['snippet'])

        return message
    except errors.HttpError, error:
        print('An error occurred: ', error)
            
def main():
    """Shows basic usage of the Gmail API.

    Creates a Gmail API service object and outputs a list of label names
    of the user's Gmail account.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)
    
#    ListLabels(service, 'me')

    messages = ListMessagesWithLabel(service, 'me', 'ETL')

    if not messages:
        print('No messages found.')
    else:
        for message in messages:
            msg = GetMessage(service, 'me', message['id'])
            print(msg['labelIds'])
            
#     test_msg_id = '153c980f25738942'
#     home_dir = os.path.expanduser('~')
#     test_dir = os.path.join(home_dir,'Desktop/test')
#     results = GetAttachments(service, 'me', test_msg_id, test_dir)
         
if __name__ == '__main__':
    main()