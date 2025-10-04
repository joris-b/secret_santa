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

# --------------------------------------------------------------------
def checkAPI():
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

# --------------------------------------------------------------------
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
    
# --------------------------------------------------------------------
class RandomDraw():
    def __init__(self):
        self.m_occurrence = np.array([])
        self.m_avoidance = np.array([])
        self.m_names = []
        self.m_mails = dict()
        self.m_draw = dict()
        self.m_nameAvailable = dict()

    # --------------------------------------------------------------------
    def getExcelInfo(self, x_fileName):
            """
            Extracts two tables from the given Excel file: one from the 'occurrence' tab and one from the 'avoidance' tab.
            Raises a ValueError if either tab does not exist.

            Args:
                x_fileName (str): Path to the Excel file.

            Note:
                This function does not return anything. The extracted data is stored in the following object attributes:
                - self.m_occurrence
                - self.m_avoidance
                - self.m_names
                - self.m_mails
                - self.m_nameAvailable
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

            self.m_occurrence = l_occurrenceDf.to_numpy()
            self.m_avoidance = l_avoidanceDf.to_numpy()
            self.m_names = list(self.m_occurrence[:,0])
            self.m_mails = dict(zip(self.m_names, list(self.m_occurrence[:,1])))
            self.m_nameAvailable = {name: True for name in self.m_names}

    # --------------------------------------------------------------------
    def makePossibilityList(self, x_i):
        """
        Generates a list of possible names that can be drawn for the person at index x_i,
        taking into account occurrence and avoidance constraints.

        Args:
            x_i (int): Index of the person for whom to generate the possibility list.

        Returns:
            list: A list of possible names that can be drawn for the person at index x_i.
        """
        l_possibilities = []
        l_occurrences = self.m_occurrence[x_i, 2:]  # The number start at column 3
        l_maxOc = np.max(l_occurrences) + 1

        # Exclude names based on avoidance table
        l_avoidanceNames = self.m_avoidance[:,0]  # first column
        l_avoidanceExclude = set()
        for l_idx, l_name in enumerate(l_avoidanceNames):
            if l_name == self.m_names[x_i]:
                l_avoidanceExclude.add(self.m_avoidance[l_idx, 1])

        for l_j in range(len(self.m_names)):
            candidate = self.m_names[l_j]
            if (candidate != self.m_names[x_i]
                and self.m_nameAvailable[candidate]
                and candidate not in l_avoidanceExclude):
                for _ in range(l_maxOc - l_occurrences[l_j]):
                    l_possibilities.append(candidate)

        return l_possibilities


    # --------------------------------------------------------------------
    def process(self):
        l_finalDraw = dict() #Store the selected name for everyone
        
        # Use this loop to make sure that every name is picked in the end and restart in case of failure
        while any(self.m_nameAvailable.values()):
            # Reset the availability at the start of each attempt
            self.m_nameAvailable = {name: True for name in self.m_names}
            # Create a shuffled list of indexes for random draw order
            l_shuffledIndexes = list(range(len(self.m_names)))
            random.shuffle(l_shuffledIndexes)
            
            #Pick the names one by one
            for l_i in l_shuffledIndexes:
                # Create a list of possibilities to adjust weights according to the number of occurence
                l_possibilities = self.makePossibilityList(l_i)
                
                try:
                    l_chosenName = random.choice(l_possibilities)
                except IndexError:
                    print("Possibility list empty, clearing the names and restarting")
                    for l_name in self.m_names:
                        self.m_nameAvailable[l_name] = True
                        l_finalDraw[l_name] = ''
                else:
                    self.m_nameAvailable[l_chosenName] = False
                    l_finalDraw[self.m_names[l_i]] = l_chosenName
        return l_finalDraw


# --------------------------------------------------------------------
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




