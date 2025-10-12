#!/usr/bin/env python
# coding: utf-8


#Libraries specific to the mail api
#Full tutorial : https://developers.google.com/gmail/api/quickstart/python

# Need to clarify all the imports
from __future__ import print_function

# To locate the credentials file
import os.path

# Libraries used for the gmail API
import google.auth
from google.auth.transport.requests import Request
from google.auth import credentials
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.exceptions import RefreshError
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.message import EmailMessage
import base64

# Libraries used for reading the excel and the random draw
import pandas as pd
import numpy as np
import random

# ----------------------------------------------------
class MailManager:
    """
    A class to manage Gmail API interactions, including checking API access, configuring and sending emails.
    """

    # --------------------------------------------------------------------
    def __init__(self):
        self.m_scopes = ['https://mail.google.com/']
        self.m_creds = None
        self.m_from = ''
        self.m_subject = ''

    # --------------------------------------------------------------------
    def setMailAndSubject(self, x_mail, x_subject):
        """Sets the sender email and subject for the email message.

        Args:
            x_mail (str): The sender's email address.
            x_subject (str): The subject of the email.
        """
        self.m_from = x_mail
        self.m_subject = x_subject

    # --------------------------------------------------------------------
    def setCreds(self):
        """Sets up the credentials for Gmail API access."""

        self.m_creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            self.m_creds = Credentials.from_authorized_user_file('token.json', self.m_scopes)
        # If there are no (valid) credentials available, let the user log in.
        if not self.m_creds or not self.m_creds.valid:
            if self.m_creds and self.m_creds.expired and self.m_creds.refresh_token:
                try:
                    self.m_creds.refresh(Request())
                except RefreshError:
                    # Refresh token is invalid or revoked: force re-authentication
                    try:
                        os.remove('token.json')
                    except OSError as l_e:
                        print(f"Error removing token.json: {l_e}")
                    self.m_creds = None
            if not self.m_creds:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', self.m_scopes)
                self.m_creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(self.m_creds.to_json())
        return (self.m_creds and self.m_creds.valid)

    # --------------------------------------------------------------------
    def checkAPI(self):
        """Shows basic usage of the Gmail API.
        Lists the user's Gmail labels.
        """
        if not self.setCreds():
            raise RuntimeError("Could not set the credentials for the Gmail API")
        
        try:
            # Call the Gmail API
            l_service = build('gmail', 'v1', credentials=self.m_creds)
            l_results = l_service.users().labels().list(userId='me').execute()
            l_labels = l_results.get('labels', [])

            if not l_labels:
                print('No labels found.')
                return
            print('Labels:')
            for l_label in l_labels:
                print(l_label['name'])

        except HttpError as l_error:
            # TODO(developer) - Handle errors from gmail API.
            print(f'An error occurred: {l_error}')

    # --------------------------------------------------------------------
    def gmailSendMessage(self, x_recipient, x_content):
        """
        Sends an email using the Gmail API to a specified recipient with the provided content.

        Args:
            x_recipient (str): The email address of the recipient.
            x_content (str): The content of the email message.

        Returns:
            dict or None: The response from the Gmail API if the message is sent successfully, or None if an error occurs.
        """
        if not self.setCreds():
            raise RuntimeError("Could not set the credentials for the Gmail API")

        try:
            l_service = build('gmail', 'v1', credentials=self.m_creds)

            # Prepare the content
            l_message = EmailMessage()
            l_message.set_content(x_content)
            l_message['To'] = x_recipient
            l_message['From'] = self.m_from
            l_message['Subject'] = self.m_subject
            # encoded message
            l_encodedMessage = base64.urlsafe_b64encode(l_message.as_bytes()).decode()

            l_create_message = {
                'raw': l_encodedMessage
            }
            # pylint: disable=E1101
            l_sendMessage = (l_service.users().messages().send
                            (userId="me", body=l_create_message).execute())
            print(F'Message Id: {l_sendMessage["id"]}')
        except HttpError as l_error:
            print(F'An error occurred: {l_error}')
            l_sendMessage = None
        return l_sendMessage

# --------------------------------------------------------------------
class RandomDraw():
    """
    A class to manage the Secret Santa random draw process.

    This class handles the extraction of participant and avoidance data from Excel files,
    manages the state of available names, and generates possible assignments while respecting
    occurrence and avoidance constraints.

    Attributes:
        m_occurrence (np.ndarray): Array containing occurrence data from the Excel file.
        m_avoidance (np.ndarray): Array containing avoidance data from the Excel file.
        m_names (list): List of participant names.
        m_mails (dict): Dictionary mapping names to email addresses.
        m_draw (dict): Dictionary storing the final draw assignments.
        m_nameAvailable (dict): Dictionary tracking which names are still available for assignment.
    """

    # --------------------------------------------------------------------
    def __init__(self):
        self.m_occurrence = np.array([])
        self.m_avoidance = np.array([])
        self.m_names = []
        self.m_mails = dict()
        self.m_draw = dict()
        self.m_nameAvailable = dict()

    # --------------------------------------------------------------------
    def resetNameAvailability(self):
        """Resets the availability status of all participant names."""
        self.m_nameAvailable = {name: True for name in self.m_names}

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
            self.resetNameAvailability()

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
        """
        Assigns Secret Santa recipients to each participant based on occurrence and avoidance constraints.

        Returns:
            dict: A dictionary mapping each participant's name to their assigned recipient's name.
        """

        l_finalDraw = dict() #Store the selected name for everyone
        self.resetNameAvailability()
        
        # Use this loop to make sure that every name is picked in the end and restart in case of failure
        # Add a counter to avoid infinite loops
        l_attemptCounter = 0
        while any(self.m_nameAvailable.values()) and l_attemptCounter < 100:
            # Reset the availability at the start of each attempt
            self.resetNameAvailability()
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
            l_attemptCounter += 1
        
        if l_attemptCounter == 100:
            raise RuntimeError("Failed to complete the draw after 100 attempts. Please check the constraints.")
        return l_finalDraw


# --------------------------------------------------------------------
if __name__ == '__main__':
    draw = RandomDraw()
    mailer = MailManager()

    # Perform the draw
    draw.getExcelInfo("template.ods")
    finalDraw = draw.process()
    mails = draw.m_mails

    mailer.setMailAndSubject("your.email@gmail.com", "Secret Santa")
    # mailer.checkAPI() # Uncomment this line to check the API connection

    for name in finalDraw.keys():
        print(f"Trying to send mail to : {name}")
        try:
            # Commented during development phase
            content = f"Hello {name},\n\nYou have been chosen to give a gift to {finalDraw[name]}!\n\nHappy gifting!\n"
            mailer.gmailSendMessage(mails[name], content)
            print(f'Mail sent successfully to {name}')
        except Exception as ex:
            print(f'Something went wrong... : {ex}')




