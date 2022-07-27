"""
Sync calendar and attendance sheets.

Go through attendee response in Google Calendar.
If attendee accepted, mark it as such in the Google Sheets
If attendee declined, mark it as such in the Google Sheets

How to use:

- python3 -m pip install -U pip
- python3 -m pip install -U google-api-python-client google-auth-httplib2 google-auth-oauthlib


In Google cloud console, create the service key, save it as credentials.json in the same directory

Enable the products: Google Sheets, Google Calendar

Require read write on Google Sheets, and readonly on Calendar


"""

# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/calendar.events",
]


import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


SHEETS_ID = ""  # the Google Sheets ID
RANGE = "Acceptance list"  # the tab name

EVENT_CONFIRMED = "confirmed"
EVENT_DECLINED = "declined"

CALENDAR_EVENT_ID_PART1 = ""  # The Google Calendar Event ID
CALENDAR_EVENT_ID_PART2 = ""  # The Google Calendar Event ID

EVENT_IDS = [CALENDAR_EVENT_ID_PART1, CALENDAR_EVENT_ID_PART2]


class CalendarSync:
    def __init__(self):
        self.creds = self.get_credentials()
        self.sheets_service = build("sheets", "v4", credentials=self.creds)
        self.calendar_service = build("calendar", "v3", credentials=self.creds)
        self.attendees = {}
        self.load_sheets()

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

    def process_calendar(self):
        for event in EVENT_IDS:
            event = (
                self.calendar_service.events()
                .get(calendarId="primary", eventId=event)
                .execute()
            )
            attendees = event.get("attendees")
            for attendee in attendees:
                status = ""
                if attendee["responseStatus"] == "accepted":
                    status = EVENT_CONFIRMED
                elif attendee["responseStatus"] == "declined":
                    status = EVENT_DECLINED

                index = self.get_attendee_index(attendee["email"])
                if status in [EVENT_CONFIRMED, EVENT_DECLINED] and index > 0:
                    self.update_sheets_value(f"O{index}", status)
                    print(attendee["email"], index, status)
                else:
                    print(attendee["email"], index, "noop")

    def load_sheets(self):
        sheet = self.sheets_service.spreadsheets()

        result = sheet.values().get(spreadsheetId=SHEETS_ID, range=RANGE).execute()
        values = result.get("values", [])

        index = 1
        for row in values:
            self.process_row(index, row)
            index += 1

    def get_attendee_index(self, email_address):
        if self.attendees.get(email_address):
            return self.attendees[email_address]
        return -1

    def process_row(self, index, row):
        email_address = row[1]
        self.attendees[email_address] = index

    def update_sheets_value(self, cell_index, data):
        """"""
        range = f"{RANGE}!{cell_index}"
        body = {"values": [[data]]}
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


if __name__ == "__main__":
    email_sender = CalendarSync()
    email_sender.process_calendar()
