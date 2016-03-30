"""Send an email message from the user's account.
"""
import httplib2
import os


from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://mail.google.com/'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatability with Python 2.6
            credentials = tools.run(flow, store)
        print(('Storing credentials to ' + credential_path))
    return credentials


import base64
from email import encoders
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from apiclient import errors


def SendMessage(service, user_id, message):
  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
  try:
    message = (service.users().messages().send(userId=user_id, body=message)
               .execute())
    print(('Message Id: %s' % message['id']))
    return message
  except errors.HttpError as error:
    print(('An error occurred: %s' % error))


def CreateMessage(sender, to, subject, message_text):
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
  message['to'] = ", ".join(to)
  message['from'] = sender
  message['subject'] = subject
  return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}


def CreateMessageWithAttachment(
    sender, to, subject, message_text, file_dir, filename):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file_dir: The directory containing the file to be attached.
    filename: The name of the file to be attached.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEMultipart()
  message['to'] = ", ".join(to)
  message['from'] = sender
  message['subject'] = subject

  msg = MIMEText(message_text)
  message.attach(msg)

  path = os.path.join(file_dir, filename)

  content_type = 'application/octet-stream'
  main_type, sub_type = content_type.split('/', 1)

  fp = open(path, 'rb')
  msg = MIMEBase(main_type, sub_type)
  
  msg.set_payload(fp.read())
  encoders.encode_base64(msg)
  fp.close()
   
  msg.add_header('Content-Disposition', 'attachment', filename=filename)
  message.attach(msg)
  
  
  return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}


###TEST###
#credentials = get_credentials()
#http = credentials.authorize(httplib2.Http())
#service = discovery.build('gmail', 'v1', http=http)

#testmsg = CreateMessage('galperin.sergey@gmail.com','galperin.sergey@gmail.com', 'test subject', 'test text')
#testmsg = CreateMessageWithAttachment('galperin.sergey@gmail.com','galperin.sergey@gmail.com','test subject','test text','teachers_2016-01','Агапонова_2016-01.xls')
#SendMessage(service,'me',testmsg)