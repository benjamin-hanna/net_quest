from __future__ import print_function
import pickle
import os.path
from pathlib import Path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from oauth2client import tools

import base64

from email.mime.text import MIMEText

import os

from googleapiclient import errors

try:
    import argparse

    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete the file token.pickle.
SCOPES = 'https://www.googleapis.com/auth/gmail.compose'
APPLICATION_NAME = 'net_quest'


def make_connection():

    client_secrets_file = Path("./credentials.json")


    credentials = None
    """
    The file token.pickle stores the user's access and refresh tokens, and is
    created automatically when the authorization flow completes for the first time.
    """
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            credentials = pickle.load(token)
            return credentials

    # If there are no (valid) credentials available, let the user log in.
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(client_secrets_file, SCOPES)
            credentials = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(credentials, token)
            print("initial connection established. re-run net_quest")
            exit(0)



def create_draft(service, user_id, message_body, subject):
  """Create and insert a draft email. Print the returned draft's message and id.
  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message_body: The body of the email message, including headers.
  Returns:
    Draft object, including draft id and message meta data.
  """
  try:
    message = {'message': message_body}
    draft = service.users().drafts().create(userId=user_id, body=message).execute()

    print("\nDraft created: ", subject, "\n")
    return draft

  except errors.HttpError as error:
    print('An error occurred: %s' % error)
    return None


def create_message(sender, to, subject, message_text):
    """Create a message for an email.
   Args:
     sender: Email address of the sender.
     to: Email address of the receiver.
     subject: The subject of the email message.
     message_text: The text of the email message.
   Returns:
     An object containing a base64url encoded email object.
   """
    message = MIMEText(message_text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_string().encode('UTF-8')).decode('ascii')}


def make_drafts(subject, message):

    credentials = make_connection()

    service = build('gmail', 'v1', credentials=credentials)

    msg = create_message('me', None, subject, message)

    create_draft(service, 'me', msg, subject)
