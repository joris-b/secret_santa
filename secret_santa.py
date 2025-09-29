#!/usr/bin/env python
# coding: utf-8


#Libraries specific to the mail api
#Full tutorial : https://developers.google.com/gmail/api/quickstart/python

# Need to clarify all the imports
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

# Libraries used for reading the excel 
# and the random draw
import pandas as pd
import numpy as np
import random

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://mail.google.com/']

def mailExample():
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

def gmail_send_message(x_recipient, x_content):
    """
    Sends an email using the Gmail API to a specified recipient with the provided content.

    Args:
        x_recipient (str): The email address of the recipient.
        x_content (str): The content of the email message.

    Returns:
        dict or None: The response from the Gmail API if the message is sent successfully, or None if an error occurs.
    """
    if os.path.exists('token.json'):
        l_creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not l_creds or not l_creds.valid:
        if l_creds and l_creds.expired and l_creds.refresh_token:
            l_creds.refresh(Request())
        else:
            l_flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            l_creds = l_flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as l_token:
            l_token.write(l_creds.to_json())


    try:
        l_service = build('gmail', 'v1', credentials=l_creds)
        l_message = EmailMessage()

        l_message.set_content(x_content)

        l_message['To'] = x_recipient
        l_message['From'] = "your.email@gmail.com" #replace with your own mail
        l_message['Subject'] = "Secret Santa"

        # encoded message
        l_encoded_message = base64.urlsafe_b64encode(l_message.as_bytes()).decode()

        l_create_message = {
            'raw': l_encoded_message
        }
        # pylint: disable=E1101
        l_send_message = (l_service.users().messages().send
                        (userId="me", body=l_create_message).execute())
        print(F'Message Id: {l_send_message["id"]}')
    except HttpError as l_error:
        print(F'An error occurred: {l_error}')
        l_send_message = None
    return l_send_message

#Create the message
def createMsg(x_chosenPerson):
    return "Ho ho hoo ! Cette annee, tu offres un cadeau a : " + (x_chosenPerson)
    
#Extract info from excel file
def getExcelInfo(x_fileName):
        """
        Extracts two tables from the given Excel file: one from the 'occurrence' tab and one from the 'avoidance' tab.
        Raises a ValueError if either tab does not exist.

        Args:
            x_fileName (str): Path to the Excel file.

        Returns:
            tuple: (occurrence_table, avoidance_table) as numpy arrays.
        """
        try:
            l_xls = pd.ExcelFile(x_fileName)
        except Exception as l_err:
            raise ValueError(f"Could not open Excel file: {l_err}")

        l_required_sheets = ['occurrence', 'avoidance']
        l_missing_sheets = [l_sheet for l_sheet in l_required_sheets if l_sheet not in l_xls.sheet_names]
        if l_missing_sheets:
            raise ValueError(f"Missing required sheet(s): {', '.join(l_missing_sheets)}")

        l_occurrenceDf = pd.read_excel(l_xls, sheet_name='occurrence')
        l_avoidanceDf = pd.read_excel(l_xls, sheet_name='avoidance')

        l_occurrenceTable = l_occurrenceDf.to_numpy()
        l_avoidanceTable = l_avoidanceDf.to_numpy()
        return (l_occurrenceTable, l_avoidanceTable)
   
#%%
#Create a function that return True when a least one of the names is free
def namesRemaining(nameDict):
    result = False
    for remain in nameDict.values():
        result = result or remain
    return result

#Manage the exclusion between people (mostly for couples)
def isAllowed(name1, name2):
    result = True
    # Function to complete later according to 
    # tab avoidance in the excel file
    
    return result   

    
#%%
#Process the selection
def randomDraw(x_data):
    l_storage = dict() #Store the selected name for everyone
    l_names = list(x_data[:,0])
    #Dictionnary to check if the person is already picked for someone else
    l_namesFree = dict()
    for l_name in l_names:
        l_namesFree[l_name] = True
    
    #Use this loop to make sure that every name is picked in the end
    while namesRemaining(l_namesFree):
        #List of the indexes of the names (in initial order then shuffled)
        l_indexes = list()
        for l_i in range(np.size(l_names)):
            l_indexes.append(l_i)
        l_shuffledIndexes = l_indexes.copy()
        #we want to shuffle indexes, not names
        random.shuffle(l_shuffledIndexes)

        #Pick the names one by one
        for l_i in l_shuffledIndexes:
            l_occurences = data[l_i, 2:] #The number start at row 3
            l_maxOc = np.max(l_occurences) + 1
            # Create a list of possibilities to adjust weights according to the number of occurence
            l_possibilities = list()
            for l_j in range(np.size(l_names)): #Add verification to avoid multiple and self picking
                if(l_names[l_i] != l_names[l_j] and l_namesFree[l_names[l_j]]
                   and isAllowed(l_names[l_i], names[l_j])):
                    for _ in range(l_maxOc - l_occurences[l_j]):
                        l_possibilities.append(l_names[l_j])
            
            try:
                l_chosenName = random.choice(l_possibilities)
            except IndexError:
                print("Possibility list empty, clearing the names and restarting")
                for l_name in l_names:
                    l_namesFree[l_name] = True
                    l_storage[l_name] = ''
            else:
                l_namesFree[l_chosenName] = False
                l_storage[l_names[l_i]] = l_chosenName
                
    return l_storage


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
            # Commented during development phase
            # gmail_send_message(mails[name],createMsg(tirage[name]))
            print('Mail sent succesfully to %s'%(name)) 
        except Exception as ex:
            print('Something went wrong... : ', ex)




