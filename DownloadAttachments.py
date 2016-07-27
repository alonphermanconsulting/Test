# Something in lines of http://stackoverflow.com/questions/348630/how-can-i-download-all-emails-with-attachments-from-gmail
# Make sure you have IMAP enabled in your gmail settings.
# Right now it won't download same file name twice even if their contents are different.

import email, email.utils
import datetime, time
import getpass, imaplib
import os,re
import sys
import ConfigParser
import traceback

from rsa._version200 import from64

config = ConfigParser.RawConfigParser()
config.read('config.properties')
#detach_dir = '.'
attachment_directory=config.get('general','attachment_directory')
if 'attachments' not in os.listdir(attachment_directory):
    os.mkdir(attachment_directory + r'/attachments')

#userName = raw_input('Enter your GMail username:')
#passwd = getpass.getpass('Enter your password: ')
userName=config.get('general','gmail_username')
passwd=config.get('general','gmail_password')
emails_to_ignore=config.get('general','emails_to_ignore')
emails_to_ignore_list = emails_to_ignore.split(',')

try:
    imapSession = imaplib.IMAP4_SSL('imap.gmail.com')
    typ, accountDetails = imapSession.login(userName, passwd)
    if typ != 'OK':
        print 'Not able to sign in!'
        raise

    imapSession.select('[Gmail]/All Mail')
    typ, data = imapSession.search(None, 'ALL')
    if typ != 'OK':
        print 'Error searching Inbox.'
        raise

    # Iterating over all emails
    for msgId in data[0].split():
        typ, messageParts = imapSession.fetch(msgId, '(RFC822)')
        if typ != 'OK':
            print 'Error fetching mail.'
            raise

        emailBody = messageParts[0][1]
        mail = email.message_from_string(emailBody)
        #print 'Message:', mail['Subject']
        #print 'Raw Date:', mail['Date']
        #print 'From:', mail['From']
        fromEmail = mail['From']
        if ('<' in fromEmail):
            start = fromEmail.index('<')+1
            end = fromEmail.index('>')
            sender = fromEmail[start:end]
        else:
            sender = fromEmail

        for part in mail.walk():
            if part.get_content_maintype() == 'multipart':
                # print part.as_string()
                continue
            if part.get('Content-Disposition') is None:
                # print part.as_string()
                continue
            if sender in emails_to_ignore:
                print('ignoring email from ',sender)
                continue
            subject_line = mail['subject'].replace(':','')
            parsedDate = email.utils.parsedate_tz(mail['Date'])
            formattedDate = time.strftime("%Y-%m-%d %Hhrs%Mmin%Ssec", parsedDate[0:9])
            print 'subject_line:', subject_line
            print 'sender:', sender
            print 'formattedDate:', formattedDate
            print 'part.get_filename:', part.get_filename()

            if  part.get_filename() is None:
                continue

            fileName = 'MESSAGE[' + subject_line +']' + 'FROM[' + sender + ']' + 'SENT_DATE[' + formattedDate +']' + part.get_filename()
            #fileName = '@.MESSAGE[' + mail['subject'] + ']' + part.get_filename()
            if bool(fileName):
                filePath = os.path.join(attachment_directory, 'attachments', fileName)
                if not os.path.isfile(filePath):
                    print fileName
                    fp = open(filePath, 'w+')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
    imapSession.close()
    imapSession.logout()
except:
   # print ('Not able to download all attachments.', sys.exc_info())
    for frame in traceback.extract_tb(sys.exc_info()[2]):
        fname,lineno,fn,text = frame
        print "Error in %s on line %d" % (fname, lineno)