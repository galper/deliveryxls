""" We have one .xls file and we have to split it to many files.
Each string to 1 file and then send to e-mails.

NB! save names of sheets: total, ф=н, all
    report name looks like blablaYYYY_MM 

"""

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
import pandas as pd
import os
import pickle

from pandas import ExcelWriter
import sys

def CreateMailList(tlist):
    
    mdict=dict()
    for teacher in tlist:
        print('Enter e - mail for ' + teacher)        
        mdict[teacher] = [str(x) for x in input('-->').split()]
    with open('mails.txt', 'wb') as handle:
        pickle.dump(mdict, handle)
    return mdict

def EditMailList(tlist):
    try:
        with open('mails.txt', 'rb') as handle:
               mdict = pickle.loads(handle.read())
               for teacher in tlist:
                   if teacher not in list(mdict.keys()):
                       print('No mail for {}.'.format(teacher))
                       mdict[teacher] = input('Enter e-mail -->')
    except: pass
    
    while True:
        print('Do you want to edit mails?(y/!y)')
        ans = input('-->')
        if ans =='y':
            continue
        else:
            break
        print('Teachers and mails:')
        print(mdict)
        print('Whose mail you want to change? ')
        cur_teacher = input('-->')
        if cur_teacher in tlist:
            print('Enter new mail')
            cur_mail = [str(x) for x in input('-->').split()]
            mdict[cur_teacher] = cur_mail
        else:
            print('No teacher with this surname in this report')
    with open('mails.txt', 'wb') as handle:
            pickle.dump(mdict, handle)
    return mdict        


credentials = se.get_credentials()
http = credentials.authorize(httplib2.Http()) #Authentification with gmail api
service = discovery.build('gmail', 'v1', http=http)  


while True:
    try:
        reportname = input("enter report name please (blablabla_YYYY-MM.xls): ")
    except ValueError:
            print("I don't understand you")
            continue
    if os.path.exists(reportname):
        print('I find a file. Just wait')
        break
    else:
        print('No file! Try Again: ')
        continue

    
date = reportname[-11:-4]     
folder = 'teachers' + '_' + date

df = pd.read_excel(reportname, sheetname=['ф=н','all','total'])

df['ф=н'].index = df['ф=н'][df['ф=н'].columns[0]]
df['ф=н'] = df['ф=н'].drop(df['ф=н'].columns[[0]],axis=1)
df['all'].index = df['all'][df['all'].columns[0]]
df['all'] = df['all'].drop(df['all'].columns[[0]],axis=1)

df2 = df['total']
df2 = df2.dropna(subset=df2.columns[[0]]) # Delete na from sheet
df2.index = df2[df2.columns[0]] # make index from first column
df2 = df2.drop(['Бонус','Общий итог'])
df2 = df2.drop(df2.columns[[0]],axis=1)
df['total'] = df2

teacherlist = df2.index

try:
    os.mkdir(folder)
except:
    print("Be careful! Teacher's folder already exist.")

if not os.path.exists('mails.txt'):
    print('Creating Mail List')
    maildict = CreateMailList(teacherlist)
else:
    print('Editing Mail List')
    maildict = EditMailList(teacherlist)

for teacher in teacherlist:
    file_path = folder + '/' + teacher +'_' + date + '.xls'
    if os.path.exists(file_path):
        print('File already exist!')
    else:   
        writer = ExcelWriter(file_path)
        try:
            df['ф=н'].xs(teacher).to_excel(writer,'ф=н')
        except:
            print('Empty sheet "ф=н" for {0}'.format(teacher))
        try:
            df['all'].xs(teacher).to_excel(writer,'all')
        except:
            print('Empty sheet "all" for {0}'.format(teacher))
        df['total'].xs(teacher).to_frame().T.to_excel(writer,'total')
        writer.save()
    testmsg = se.CreateMessageWithAttachment('galperin.sergey@gmail.com', maildict[teacher], 'Docs for ' + date, 'test text', folder,teacher +'_' + date + '.xls')
    se.SendMessage(service,'me',testmsg)

        
    
        

