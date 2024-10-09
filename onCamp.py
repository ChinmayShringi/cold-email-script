from google.oauth2 import service_account
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import json
import os
import pickle

# Constants
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
CREDENTIALS_FILE = 'creds/credentials.json'  # OAuth2 credentials file
TOKEN_FILE = 'creds/token.json'              # File to store the access and refresh tokens
JSON_EMAIL_FILE = 'companyEmail.json'  # JSON file with email information
SENT_FILE = 'sent.json'                # JSON file to keep track of sent emails
RESUME_FILE_NAME = 'chinmay_resume.pdf' # Resume attachment
RESUME_FILE = 'resumes/Chinmay_Shringi.pdf'    # Resume attachment

# Load email details from the JSON file
with open(JSON_EMAIL_FILE, 'r') as file:
    email_data = json.load(file)

# Load sent emails from the sent.json file
if os.path.exists(SENT_FILE):
    with open(SENT_FILE, 'r') as file:
        sent_emails = json.load(file)
else:
    sent_emails = []

# Load or refresh OAuth2 credentials
creds = None
if os.path.exists(TOKEN_FILE):
    with open(TOKEN_FILE, 'rb') as token:
        creds = pickle.load(token)

# If credentials are not available or are invalid, go through the OAuth flow
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except RefreshError:
            print("The token has expired, and needs re-authentication.")
    else:
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)

    # Save the credentials for the next run
    with open(TOKEN_FILE, 'wb') as token:
        pickle.dump(creds, token)

# Loop through each email entry in the JSON file
for email_details in email_data:
    to_email = email_details["to_email"]
    subject = email_details["subject"]
    body = email_details["body"]

    # Check if the email has already been sent
    if to_email in sent_emails:
        print(f'Email to {to_email} has already been sent. Skipping...')
        continue  # Skip sending the email if already sent

    try:
        # Create a MIMEMultipart object to represent the email
        msg = MIMEMultipart()
        msg['From'] = 'Chinmay Shringi <cs7810@nyu.edu>'
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Attach the resume file if it exists
        if os.path.exists(RESUME_FILE):
            attachment = MIMEBase('application', 'octet-stream')
            with open(RESUME_FILE, 'rb') as file:
                attachment.set_payload(file.read())
            encoders.encode_base64(attachment)
            attachment.add_header('Content-Disposition', f'attachment; filename={RESUME_FILE_NAME}')
            msg.attach(attachment)

        # Encode the email in base64
        raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()

        # Create a Gmail API service
        service = build('gmail', 'v1', credentials=creds)

        # Send the email
        message_body = {
            'raw': raw_message
        }
        sent_message = service.users().messages().send(userId='me', body=message_body).execute()

        print(f'Email sent successfully to {to_email}! Message Id: {sent_message["id"]}')

        # Append the sent email to the sent_emails list
        sent_emails.append(to_email)

    except HttpError as error:
        print(f'An error occurred: {error}')

    except RefreshError as refresh_error:
        print(f'Token expired, need to refresh or re-authenticate: {refresh_error}')

# Write updated sent emails back to the sent.json file
with open(SENT_FILE, 'w') as file:
    json.dump(sent_emails, file)
