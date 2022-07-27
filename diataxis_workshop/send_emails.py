"""
Sending emails to participants via email.

Steps:
- Read the Google Sheets, Acceptance List tab
- Iterate through everything.
- if Column J == "accept" and column K is empty:
    then send the acceptance email, and update column K with timestamp
- if Column J == "waitlist" and column K is empty:
    then send the waitlist email, and update column L with timestamp
- if Column J == "accept" and column K is not empty:
    then do nothing. It means the person has been invited
- if Column J == "waitlist" and column L is not empty:
    then do nothing. It means the person has been waitlisted
- if Column J == "accept" and column L is not empty:
    then send the acceptance email (same logic as #1 above). this means person was changed from waitlist to accept.

How to use:

- python3 -m pip install -U pip
- python3 -m pip install -U google-api-python-client google-auth-httplib2 google-auth-oauthlib

In Google cloud console, create the service key, save it as credentials.json in the same directory

Enable the products: GMail, Google Sheets, Google Calendar

Require read and write on GMail, Sheets, and Calendar


"""

# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/calendar.events",
]


import base64
from email.mime.text import MIMEText
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from string import Template

EMAIL_FROM = "Mariatta <mariatta@python.org>"
EMAIL_CC = ""  # FILL IT IN

SUBJECT_CONFIRMED = "Attendance confirmation for Python Docs Diátaxis Workshop"
SUBJECT_WAITLISTED = "Waitlisted for Python Docs Diátaxis Workshop"

SHEETS_ID = ""  # the Google Sheets ID
RANGE = "Acceptance list"  # the tab name

STATUS_ACCEPT = "accept"
STATUS_WAITLIST = "waitlist"
CALENDAR_EVENT_ID_PART1 = ""  # The Google Calendar Event ID
CALENDAR_EVENT_ID_PART2 = ""  # The Google Calendar Event ID


class EmailSender:
    def __init__(self):
        self.creds = self.get_credentials()
        self.email_service = build("gmail", "v1", credentials=self.creds)
        self.sheets_service = build("sheets", "v4", credentials=self.creds)
        self.calendar_service = build("calendar", "v3", credentials=self.creds)

    def get_credentials(self):
        creds = None
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES
                )
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(creds.to_json())

        return creds

    def get_confirmed_html_content(self, to_name):
        with open("confirmation_template.html", "r") as f:
            src = Template(f.read())
            text = src.substitute(name=to_name)
            return text

    def get_waitlisted_html_content(self, to_name):
        with open("waitlist_template.html", "r") as f:
            src = Template(f.read())
            text = src.substitute(name=to_name)
            return text

    def send_confirmation_email(self, to_name, to_address):
        """Send the email"""
        html = self.get_confirmed_html_content(to_name)
        subject = f"(test) {SUBJECT_CONFIRMED}"
        return self.send_email(subject, to_address, html)

    def send_waitlist_email(self, to_name, to_address):
        """Send the email"""
        html = self.get_waitlisted_html_content(to_name)
        subject = f"(test) {SUBJECT_WAITLISTED}"
        return self.send_email(subject, to_address, html)

    def send_email(self, subject, to_address, email_html_message):

        try:

            email_message = MIMEText(email_html_message, "html")
            email_message["To"] = to_address
            email_message["From"] = EMAIL_FROM
            email_message["cc"] = EMAIL_CC
            email_message["Subject"] = subject

            encoded_message = base64.urlsafe_b64encode(
                email_message.as_bytes()
            ).decode()

            create_message = {
                "raw": encoded_message,
            }

            send_message = (
                self.email_service.users()
                .messages()
                .send(userId="me", body=create_message)
                .execute()
            )
            print(f'Message Id: {send_message["id"]}')
        except HttpError as error:
            print(f"An error occurred: {error}")
            send_message = None
        return send_message

    def process_sheets(self):
        sheet = self.sheets_service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SHEETS_ID, range=RANGE).execute()
        values = result.get("values", [])

        index = 1
        for row in values:
            self.process_row(index, row)
            index += 1

    def process_row(self, index, row):
        timestamp = row[0]
        email_address = row[1]
        name = row[2]
        try:
            confirmation_status = row[9]
        except IndexError:
            confirmation_status = ""

        try:
            confirmation_sent = row[10]
        except IndexError:
            confirmation_sent = False

        try:
            waitlist_sent = row[11]
        except IndexError:
            waitlist_sent = False

        if confirmation_status == STATUS_ACCEPT:
            if not confirmation_sent:
                print(f"{name}, CONFIRMED, Sending email")
                email_sent = self.send_confirmation_email(
                    to_name=name, to_address=email_address
                )
                self.update_sheets_value(cell_index=f"K{index}", data=email_sent)
                self.add_attendee(email_address, CALENDAR_EVENT_ID_PART1)
                self.add_attendee(email_address, CALENDAR_EVENT_ID_PART2)

            else:
                print(f"{name} noop")
        elif confirmation_status == STATUS_WAITLIST:
            if not waitlist_sent:
                print(f"{name}, WAITLIST, Sending email")
                email_sent = self.send_waitlist_email(
                    to_name=name, to_address=email_address
                )
                self.update_sheets_value(cell_index=f"L{index}", data=email_sent)
            else:
                print(f"{name} noop")

    def update_sheets_value(self, cell_index, data):
        """"""
        range = f"{RANGE}!{cell_index}"
        body = {"values": [[data["id"]]]}
        result = (
            self.sheets_service.spreadsheets()
            .values()
            .update(
                spreadsheetId=SHEETS_ID,
                range=range,
                valueInputOption="USER_ENTERED",
                body=body,
            )
            .execute()
        )

    def add_attendee(self, email_address, event_id):
        event = (
            self.calendar_service.events()
            .get(calendarId="primary", eventId=event_id)
            .execute()
        )
        if not event.get("attendees"):
            event["attendees"] = [{"email": email_address}]
        else:
            event["attendees"].append({"email": email_address})

        updated_event = (
            self.calendar_service.events()
            .update(
                calendarId="primary",
                eventId=event_id,
                body=event,
                sendNotifications=True,
            )
            .execute()
        )


if __name__ == "__main__":
    email_sender = EmailSender()
    email_sender.process_sheets()
