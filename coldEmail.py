import time
import random  # Import random module for selecting random subject and body
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
SCOPES = ['https://www.googleapis.com/auth/gmail.send', 'https://www.googleapis.com/auth/gmail.modify']
CREDENTIALS_FILE = 'creds/credentials.json'  # OAuth2 credentials file
TOKEN_FILE = 'creds/token.json'              # File to store the access and refresh tokens
JSON_EMAIL_FILE = 'companyEmail.json'  # JSON file with email information
SENT_FILE = 'sent.json'                # JSON file to keep track of sent emails
RESUME_FILE = 'resumes/chinmay_ats_resume.pdf' # Resume attachment
RESUME_FILE_NAME = 'chinmay_resume.pdf' # Resume attachment

# Load email details from the JSON file
with open(JSON_EMAIL_FILE, 'r') as file:
    email_data = json.load(file)

subjects = email_data.get("subject", [])  # List of subjects
body_templates = email_data.get("bodyTemplates", [])  # List of body templates
contacts = email_data.get("contacts", [])
company_name = email_data.get("company_name", "").lower()  # Get the company name from JSON and convert to lowercase for comparison
email_greet = email_data.get("email_greet", [])
email_outro = email_data.get("email_outro", [])
email_links = email_data.get("email_links", "")

# To track skipped emails and the reason why
skipped_emails = []

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

# Function to check if company name matches email domain
def is_company_match(to_email, company_name):
    email_domain = to_email.split('@')[1].split('.')[0].lower()
    return company_name in email_domain

hasAsked = False

# Loop through each contact in the JSON file
for contact in contacts:
    to_email = contact.get("email")
    recipient_name = contact.get("name")
    # Skip if the email or name is missing
    if not to_email or not recipient_name:
        reason = 'Missing email or name'
        print(f'Skipping contact due to missing email or name: {contact}')
        skipped_emails.append({"email": to_email, "reason": reason})
        continue

    # Check if the email domain matches the company name
    if not hasAsked and not is_company_match(to_email, company_name):
        user_input = input(f"Email domain for {to_email} does not match company name '{company_name.capitalize()}'. Proceed? (Y/N): ")
        hasAsked = True
        if user_input.lower() != 'y':
            print(f"Skipping email to {to_email} based on user input.")
            skipped_emails.append({"email": to_email, "reason": "User opted not to send"})
            continue

    # Check if the email has already been sent
    if to_email in sent_emails:
        print(f'Email to {to_email} has already been sent. Skipping...')
        continue  # Skip sending the email if already sent

    try:
        # Randomly select a body template, outro, and greet
        body_template = random.choice(body_templates).get("body")
        outro = random.choice(email_outro)
        greet = random.choice(email_greet)

        # Replace placeholders with actual values
        body = body_template.replace("{{RECIPIENT_NAME}}", recipient_name)
        body = body.replace("{{EMAIL_OUTRO}}", outro)
        body = body.replace("{{COMPANY_NAME}}", company_name.capitalize())
        body = body.replace("{{EMAIL_GREET}}", greet)
        body = body.replace("{{EMAIL_LINKS}}", email_links)

        # Check if replacement was successful
        if any(placeholder in body for placeholder in ["{{RECIPIENT_NAME}}", "{{COMPANY_NAME}}", "{{EMAIL_OUTRO}}", "{{EMAIL_GREET}}", "{{EMAIL_LINKS}}"]):
            reason = 'Missing placeholders or failed replacements'
            print(f'Skipping email to {to_email} due to missing placeholders')
            skipped_emails.append({"email": to_email, "reason": reason})
            print(body)
            continue

        # Randomly select a subject from the list
        subject = random.choice(subjects)

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
        waiting_time = random.randint(30, 40)
        print(f'Sending email to {to_email} in {waiting_time} seconds...')
        time.sleep(waiting_time)
        sent_message = service.users().messages().send(userId='me', body=message_body).execute()

        print(f'Email sent successfully to {to_email}! Message Id: {sent_message["id"]}')

        # Write to sent.json file immediately after sending
        if to_email != 'chinmayusa4@gmail.com':
            sent_emails.append(to_email)
            with open(SENT_FILE, 'w') as file:
                json.dump(sent_emails, file)
            print(f"Recorded sent email to {to_email} in {SENT_FILE}")

        # Now attempt to delete the email from the 'Sent' folder
        try:
            service.users().messages().delete(userId='me', id=sent_message['id']).execute()
            print(f'Email deleted from Sent folder for {to_email}.')
        except HttpError as delete_error:
            print(f'Failed to delete email for {to_email} from Sent folder: {delete_error}')

    except HttpError as error:
        print(f'An error occurred: {error}')

    except RefreshError as refresh_error:
        print(f'Token expired, need to refresh or re-authenticate: {refresh_error}')

# Print the list of emails that were skipped and the reason
if skipped_emails:
    print("\nEmails that were not sent and the reason:")
    for skipped in skipped_emails:
        print(f"Email: {skipped['email']}, Reason: {skipped['reason']}")
