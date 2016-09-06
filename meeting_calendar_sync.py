import httplib2
import os
import sys
from datetime import datetime
import pytz
from optparse import OptionParser
from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
from datetime import datetime
import calendar
import xml.etree.ElementTree as ET

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

CALENDAR_RSS_URL = "http://intranet.innismaggiore.net/calendar/feeds/meetings/dave/"
MEETING_CALENDAR_ID = "h7h8ij3d5ohrtm7cdljbubjaig@group.calendar.google.com"
CLIENT_SECRET_FILE = 'calendar.json'
SCOPES = 'https://www.googleapis.com/auth/calendar'
APPLICATION_NAME = 'Work Calendar Sync'

class Event(object):
    def __init__(self, element):
        self.title = element.find('title').text
        self.description = element.find('description').text
        self.guid = element.find('guid').text
        self.date = self.parse_date(element.find('pubDate').text)

    def parse_date(self, date_str):
        (day_name, rest) = date_str.split(',')
        pieces = rest.strip().split(' ')
        (d,m,y,t) = (pieces[0], datetime.strptime(pieces[1], '%b').month, pieces[2], pieces[3])
        is_dst = self.date_isdst(y, m, d) 
        # intranet doesn't handle dst properly on events
        tz_offset = '-0400' if is_dst else '-0500'
        return '{}-{}-{}T{}{}'.format(y, m, d, t, tz_offset)

    def date_isdst(self, y, m, d):
        return bool(pytz.timezone('America/New_York').dst(datetime(int(y), int(m), int(d)), is_dst=None))

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.googleapi')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar_sync.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def get_calendar_service():
     credentials = get_credentials()
     http = credentials.authorize(httplib2.Http())
     return discovery.build('calendar', 'v3', http=http)

def get_work_calendar_events():
    http = httplib2.Http()
    resp, content = http.request(CALENDAR_RSS_URL)
    assert resp.status == 200
    root = ET.fromstring(content)
    events = root.findall('channel/item')
    all_events = [Event(e) for e in events]
    unsynced_events = []
    # see which of these are in the calendar already
    for event in all_events:
        g_cal_event = get_google_calendar_event(event)
        if not g_cal_event:
            unsynced_events.append(event)

    return unsynced_events


def get_google_calendar_event(event):
    guid = event.guid
    service = get_calendar_service()

    results = service.events().list(calendarId=MEETING_CALENDAR_ID, singleEvents=True, privateExtendedProperty="guid={}".format(guid)).execute()
    calendar_events = results.get('items', [])
    return calendar_events

def sync_events():
    work_events = get_work_calendar_events()
    if len(work_events) == 0:
        print "Nothing to sync"
    else:
        for event in work_events:
            print "Syncing event '{}'".format(event.title)
            create_event(event)

def create_event(event):
    service = get_calendar_service()
    event = {
        'summary': event.title,
        'description': event.description,
        'start': {
            'dateTime': event.date,
            'timeZone': 'America/New_York',
         },
        'end': {
            'dateTime': event.date,
            'timeZone': 'America/New_York',
         },
        'extendedProperties': {
            'private': {
                'guid': event.guid
            }
        }
    }

    event = service.events().insert(calendarId=MEETING_CALENDAR_ID, body=event).execute()
        
if __name__ == "__main__":
    sync_events()
