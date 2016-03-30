import sendemail as se
import base64
from email import encoders
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from apiclient import errors
from apiclient import discovery

import httplib2
import os

import oauth2client
from oauth2client import client
from oauth2client import tools

credentials = se.get_credentials()
http = credentials.authorize(httplib2.Http())
service = discovery.build('gmail', 'v1', http=http)

#testmsg = CreateMessage('galperin.sergey@gmail.com','galperin.sergey@gmail.com', 'test subject', 'test text')
testmsg = se.CreateMessageWithAttachment('galperin.sergey@gmail.com',['galperin.sergey@gmail.com','galperin.sergey222222@gmail.com'],'test subject','test text','teachers_2016-01','Агапонова_2016-01.xls')
se.SendMessage(service,'me',testmsg)
