import pandas as pd

import numpy as np
import base64

import os
import pickle
# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode
# for dealing with attachment MIME types
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from mimetypes import guess_type as guess_mime_type

from IPython.display import Markdown

# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/']
our_email = 'username@gmail.com' #replace username with email


def gmail_authenticate():
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials availablle, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

# get the Gmail API service
service = gmail_authenticate()


def search_messages(service, query):
    result = service.users().messages().list(userId='me',q=query).execute()
    messages = [ ]
    if 'messages' in result:
        messages.extend(result['messages'])
    while 'nextPageToken' in result:
        page_token = result['nextPageToken']
        result = service.users().messages().list(userId='me',q=query, pageToken=page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
    return messages


def get_body(mail):
    try:
        return base64.urlsafe_b64decode(mail["payload"]['parts'][0]['body']['data'].encode("ASCII")).decode("utf-8")
    except:
        return base64.urlsafe_b64decode(mail["payload"]['body']['data'].encode("ASCII")).decode("utf-8")



def read_message(service, message_id):

    """
    This function takes Gmail API `service` and the given `message_id` and does the following:
        - Downloads the content of the email
        - Adds email basic information (To, From, Subject & Date) and plain/text parts to dictionary
    """

    info = {}

    msg = service.users().messages().get(userId='me', id=message_id['id'], format='full').execute()
    # parts can be the message body, or attachments

    info['snippet'] = msg['snippet']
    payload = msg['payload']
    headers = payload["headers"]


    if headers:

        for header in headers:

            name = header["name"]
            value = header["value"]

            if name.lower() == 'from':
                info['from'] = value

            if name.lower() == "to":
                info['to'] = value

            if name.lower() == "subject":
                info['Subject'] = value

            if name.lower() == "date":
                info['date'] = value

    info['body'] = get_body(msg)

    return info


#define subjects/words to remove/date
subjects = ['application', 'applying', 'applied', 'resume', 'position', 'Candidacy', 'Analyst']
additional_subjects = ['']
words_to_remove_in_search = ['']
emails_to_remove_from_search = ['jobs-noreply@linkedin.com', 'jobalerts-noreply@linkedin.com', 'noreply@glassdoor.com']
after_date = '2021/04/01'


words_to_remove_in_search = ' '.join(['-'+i for i in words_to_remove_in_search])
emails_to_remove_from_search = ' '.join(['-'+i for i in emails_to_remove_from_search])

subjects = ' '.join('subject: '+i for i in subjects)
additional_subjects = ' '.join('subject: '+i for i in additional_subjects)
all_subjects = "{" + subjects + " " + additional_subjects + "}"


results = search_messages(service, f"after: {after_date} {all_subjects} in:inbox {words_to_remove_in_search} from: {emails_to_remove_from_search}")

msg_content = []

for msg in results:
    msg_content.append(read_message(service, msg))

#organize in dataframe
msg_content = pd.DataFrame(msg_content)
msg_content = msg_content[['from', 'Subject', 'to', 'date', 'snippet', 'body']]



application_did_not_move_forward_words = ['unfortunately', 'move forward with other candidates', 'highly competitive', 'regrettably', 'Regrettably', 'Unfortunately', 'not in the position to move forward',
                                            'other candidates', 'more closely align', 'not the right fit', 'not to move forward']

application_move_forward_words = ['set up', 'initial conversation', 'call', 'discussing the position', 'grab some time']

did_not_move_forward_indexes = []

move_forward_indexes = []

for i in range(len(msg_content)):
    email_body = msg_content.loc[i, 'body']
    if any(x in email_body for x in application_did_not_move_forward_words):
        did_not_move_forward_indexes.append(i)

for i in range(len(msg_content)):
    email_body = msg_content.loc[i, 'body']
    if any(x in email_body for x in application_move_forward_words):
        move_forward_indexes.append(i)




did_not_move_forward = msg_content.iloc[did_not_move_forward_indexes].reset_index(drop=True)
move_forward = msg_content.iloc[move_forward_indexes].reset_index(drop=True)
applications = msg_content[~msg_content.index.isin(did_not_move_forward_indexes) & ~msg_content.index.isin(move_forward_indexes)].reset_index(drop=True)



print('# of job applications (using gmail):', len(applications))
display(applications)

print()

print("# didn't move forward:", len(did_not_move_forward))
display(did_not_move_forward)


print()

print("# move forward:", len(move_forward))
display(move_forward)
