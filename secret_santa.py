#!/usr/bin/env python
# coding: utf-8

# In[10]:


#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep  7 19:29:43 2022

@author: joris
"""
#%%
#Libraries specific to the mail api
#Full tutorial : https://developers.google.com/gmail/api/quickstart/python

from __future__ import print_function

import os.path
import base64
from email.message import EmailMessage

import google.auth
from google.auth.transport.requests import Request
from google.auth import credentials
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

#%%
#Libraries
import pandas as pd
import numpy as np
import random
#import smtplib
#from email.mime.text import MIMEText
#from email.mime.multipart import MIMEMultipart
#import requests
#import os

#%%
# If modifying these scopes, delete the file token.json.
SCOPES = ['https://mail.google.com/']

#%%
def example():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        results = service.users().labels().list(userId='me').execute()
        labels = results.get('labels', [])

        if not labels:
            print('No labels found.')
            return
        print('Labels:')
        for label in labels:
            print(label['name'])

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')

#%%
#if __name__ == '__main__':
#    example()

#%%
def gmail_send_message(to,content):
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())


    try:
        service = build('gmail', 'v1', credentials=creds)
        message = EmailMessage()

        message.set_content(content)

        message['To'] = to
        message['From'] = "sayanel.jb@gmail.com" #replace with your own mail
        message['Subject'] = "Secret Santa NO NAME"

        # encoded message
        encoded_message = base64.urlsafe_b64encode(message.as_bytes())             .decode()

        create_message = {
            'raw': encoded_message
        }
        # pylint: disable=E1101
        send_message = (service.users().messages().send
                        (userId="me", body=create_message).execute())
        print(F'Message Id: {send_message["id"]}')
    except HttpError as error:
        print(F'An error occurred: {error}')
        send_message = None
    return send_message


#%%
#Create the message
def createMsg(chosenPerson):
    return "Ho ho hoo ! Cette annee, tu offres un cadeau a : " + (chosenPerson)
    
#%%
#Extract info from excel file
def getExcelInfo(fileName):
    dataFrame = pd.read_excel(fileName)
    data = dataFrame.to_numpy()
    return data
   
#%%
#Create a function that return True when a least one of the names is free
def namesRemaining(nameDict):
    result = False
    for remain in nameDict.values():
        result = result or remain
    return result

#%%
#Manage the exclusion between people (mostly for couples)
def isAllowed(name1, name2):
    result = True
    if (name1 == "Alex" and name2 == "Julie" 
        or name1 == "Julie" and name2 == "Alex" 
        or name1 == "Fabien" and name2 == "Pauline" 
        or name1 == "Pauline" and name2 == "Fabien" 
        or name1 == "Evans" and name2 == "Lisa" 
        or name1 == "Lisa" and name2 == "Evans" 
        or name1 == "Samy" and name2 == "Lucie" 
        or name1 == "Lucie" and name2 == "Samy"):
            result = False
    
    return result   

    
#%%
#Process the selection
def selectNames(data):
    storage = dict() #Store the selected name for everyone
    names = list(data[:,0])
    #Dictionnary to check if the person is already picked for someone else
    namesFree = dict()
    for name in names:
        namesFree[name] = True
    
    #Use this loop to make sure that every name is picked in the end
    while namesRemaining(namesFree):
        #List of the indexes of the names (in initial order then shuffled)
        indexes = list()
        for i in range(np.size(names)):
            indexes.append(i)
        shuffledIndexes = indexes.copy()
        #we want to shuffle indexes, not names
        random.shuffle(shuffledIndexes)

        #Pick the names one by one
        for i in shuffledIndexes:
            occurences = data[i, 2:] #The number start at row 3
            maxOc = np.max(occurences) + 1
            #Create a list of possibilities to adjust weights according to the number of occurence
            possibilities = list()
            for j in range(np.size(names)): #Add verification to avoid multiple and self picking
                if(names[i] != names[j] and namesFree[names[j]]
                   and isAllowed(names[i], names[j])):
                    for _ in range(maxOc - occurences[j]):
                        possibilities.append(names[j])
            
            try:
                chosenName = random.choice(possibilities)
            except IndexError:
                print("Possibility list empty, clearing the names and restarting")
                for name in names:
                    namesFree[name] = True
                    storage[name] = ''
            else:
                namesFree[chosenName] = False
                storage[names[i]] = chosenName
                
    return storage


#%%
if __name__ == '__main__':
    #aquire data and isolate names and mails
    data = getExcelInfo("secret_santa_no_name.xlsx")
    names = data[:,0]
    mails = dict()
    for i in range(np.size(names)):
        mails[names[i]] = data[i,1]
    
    #choice for everyone
    tirage = selectNames(data)
    for name in tirage.keys():        
        print("Trying to send mail to : " + name)
        try:
            gmail_send_message(mails[name],createMsg(tirage[name]))
            print('Mail sent succesfully to %s'%(name)) 
        except Exception as ex:
            print('Something went wrong... : ', ex)




