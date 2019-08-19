"""
Created on Wed Feb 27 13:07:52 2019

@author: Robert Piazza
"""

from __future__ import absolute_import, division, unicode_literals, print_function

import math
import numpy as np
import pandas as pd


import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import requests



import smtplib

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

#keys to be scrubbed if shared:
darksky_token = 'lowercase letters and numbers'

todoist_token = 'lowercase letters and numbers'

rescuetime_token = 'uppercase and lowercase letters and numbers'

#%% Weather
#Weather at home or work:
from datetime import datetime as dt
desired_location = 'Home'
locations_dict = {'Home': '47.000000,-122.000000',
             'Work': '47.000000,-122.000000',
             'Hawaii': '21.000000,-157.000000', 
             'Parents': '34.000000,-117.000000'}

weather_site = 'https://api.darksky.net/forecast/'+darksky_token+'/'+locations_dict[desired_location]
location_weather = requests.get(weather_site).json()
daily_weather = location_weather['daily']['data'][0]
weather_report = '' #This will become the full weather report text and we add as new information is pulled
#print(location_weather['daily']['summary'])
weather_report += location_weather['daily']['summary']+'\n'
weather_date = dt.fromtimestamp(daily_weather['time']).strftime('%a')
agenda_date  = dt.fromtimestamp(daily_weather['time']).strftime('%a %b-%d')
daily_summary = daily_weather['summary']
weather_report += 'Today(%s), %s High of %s at %s, low of %s at %s.\n' % (weather_date,
daily_summary, 
int(daily_weather['temperatureHigh']),
dt.fromtimestamp(daily_weather['temperatureHighTime']).strftime('%H:%M'),
int(daily_weather['temperatureLow']),
dt.fromtimestamp(daily_weather['temperatureLowTime']).strftime('%H:%M'))

try:
    for alert in location_weather['alerts']:
        #print(alert['description'])
        weather_report += alert['description'] + '\n'
except:
    #print('No Alerts')
    weather_report += 'No Alerts\n'
#print(weather_report)

#%% Calendar
"""Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server()
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)


desired_calendars = ['Todoist', '[MyEmail]@gmail.com', 'Work', '[MySportsTeamsCalendar]']

# Call the Calendar API
service = build('calendar', 'v3', credentials=creds)
calendars = pd.DataFrame(service.calendarList().list(pageToken = None).execute()['items'])

now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
agenda_date  = dt.fromtimestamp(daily_weather['time']).strftime('%a %b-%d')
compiled_events = pd.DataFrame()
##For each calendar, call next 100 events- 100 chosen because unlikely to have more than that over two days
for cal in desired_calendars:
    cal_id = calendars.loc[calendars['summary']==cal,'id'].tolist()[0]
    #print(cal, calId)

    
    #print('Getting the upcoming 100 events')
    events_result = service.events().list(calendarId = cal_id, timeMin=now,
                                        maxResults=100, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])
    compiled_events = compiled_events.append(pd.DataFrame(events),sort = True)
    #if not events:
        #print('No upcoming events found.')
        #events =[]
    #for event in events:
        #start = event['start'].get('dateTime', event['start'].get('date'))
        #print(start, event['summary'])
        
compiled_events = compiled_events.reset_index(drop=True)

## extract times from event information    
for index, time in enumerate(compiled_events['start']):
    try:
        #if a date is given, it's for local time, adjust to standard UTC start
        event_start = (pd.to_datetime(time['date'])+pd.Timedelta(hours = 8)).tz_localize('UTC')
    except:
        event_start = (pd.to_datetime(time['dateTime'])).tz_localize('UTC')
    #print(time, event_start.tz_convert('US/Pacific').strftime('%a %m.%d %I:%M %p'))
    compiled_events.loc[index, 'newStart']=event_start
    compiled_events.loc[index, 'Time'] = event_start.tz_convert('US/Pacific').strftime('%I:%M %p')

sorted_events = compiled_events.sort_values('newStart') 
today = pd.to_datetime(pd.Timestamp("today").strftime("%m/%d/%Y")).tz_localize('US/Pacific')
time_range = pd.date_range(
        today, 
        today+pd.Timedelta(days=2), 
        freq = 'D')
todays_events_bool = np.logical_and(np.logical_and(sorted_events['newStart']<time_range[1],
                              sorted_events['newStart']>time_range[0]),
                              np.logical_not(sorted_events.duplicated(['newStart','summary'])))

#Calendar Part
#print("Today,", pd.Timestamp("today").strftime('%m.%d.%a:'))
#today_header = "Today, " + pd.Timestamp("today").strftime('%m.%d.%a:')
if not sorted_events.loc[todays_events_bool,['Time', 'summary']].empty:
    today_events = sorted_events.loc[todays_events_bool,['Time', 'summary']].to_html(index=False, header = False).replace('<table border="1" class="dataframe">','<table class="table table-striped">')

else:
    today_events = 'No Events'
#print(today_events)
  


#%%Todoist

from todoist import TodoistAPI
import requests

api = TodoistAPI(todoist_token)
#api.sync()
#print(api.state['projects'])

#projects = api.state['projects']
items = api.state['items']
activity = api.activity.get()
#labels = api.state['labels']
labels_JSON = requests.get("https://beta.todoist.com/API/v8/labels", headers={"Authorization": "Bearer %s" % todoist_token}).json()

labels = pd.DataFrame(labels_JSON)

projects = pd.DataFrame(requests.get(
    "https://beta.todoist.com/API/v8/projects",
    params={
        #"project_id": 123
        #"label_id":179764
        #"filter": '(today|overdue) & @Work'
        #"lang": string
        
    },
    headers={
        "Authorization": "Bearer %s" % todoist_token
    }).json())
projects.columns = ['project_comment_count', 'project_id', 'project_indent', 'project_name', 'project_order']

    
number_completed_today = api.completed.get_stats()['days_items'][0]['total_completed']
# max of 50 items can be called at once
completed_limit = 30
completed = []
for completed_page in range(math.ceil(number_completed_today/completed_limit)):
    if completed_page == 0:
        page = None
    else:
        page = completed_limit * completed_page - 1
    completed += api.completed.get_all(limit = completed_limit, offset = page, since = today.strftime('%Y-%m-%dT00:00'))['items']#until = '2019-02-28T08:00'
completed_no_id = pd.DataFrame(completed)


#print('Completed Today:\n')
if not completed_no_id.empty:
    complete_tasks = (completed_no_id.join(
            projects.set_index('project_id'), 
            on = 'project_id').sort_values(
                    ['completed_date'])).reset_index(drop=True).loc[:,['content', 'project_name']]
    
    dupes = np.logical_not(complete_tasks.duplicated(['content','project_name']))
    complete_tasks_table = complete_tasks.loc[dupes,['content', 'project_name']].to_html(index=False, header = False).replace('<table border="1" class="dataframe">','<table class="table table-striped">')
    #print(complete_tasks.loc[dupes,['content', 'project_name']])
else:
    #print('No Tasks Completed Today')
    complete_tasks_table = 'No Tasks Completed Today'
    
    
#%% Rescuetime
#This integration requires a premium subscription
    
rescuetime_data = 'https://www.rescuetime.com/anapi/data'
rescuetime_date = pd.Timestamp("today").strftime("%Y-%m-%d")
rescuetime_data = requests.get(
    "https://www.rescuetime.com/anapi/data",
    params={\
        'key': rescuetime_token,
        'perspective': 'rank', #rank, interval or person
        'restrict_kind': 'activity',  #overview, category, activity, productivity, efficiency, document
        'interval': 'hour',
        'restrict_begin': rescuetime_date,
        'restrict_end': rescuetime_date,
        'format': 'json'
        #"filter": '(today|overdue) & @Work'
        #"lang": string
        
    }
    ).json()

rescuetime_df = pd.DataFrame(rescuetime_data['rows'], columns = rescuetime_data['row_headers'])
#rescuetime_df.loc[:,'Date']= pd.to_datetime(rescuetime_df.Date).dt.strftime('%H:%M') #only for interval
rescuetime_df.loc[:,'Minutes']=round(rescuetime_df.loc[:,'Time Spent (seconds)']/60,1)
rescuetime_table = rescuetime_df.loc[rescuetime_df['Time Spent (seconds)']>6,['Activity', 'Minutes']].to_html(index=False, header = False).replace('<table border="1" class="dataframe">','<table class="table table-striped">')

# In[71]:


html_string = '''
<html>
    <head>
        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.1/css/bootstrap.min.css">
        <style>body{ margin:0 100; background:whitesmoke; }</style>
    </head>
    <body>
        <h1>Daily Completions</h1>

        <!-- *** Section 1 *** --->
        <h2>Events Today</h2>
        ''' + today_events + '''
        
        <!-- *** Section 2 *** --->
        <h2>Tasks</h2>
        <h3> Tasks Completed Today: </h3>
        ''' + complete_tasks_table + '''
        
        <!-- *** Section 4 *** --->
        <h2>Time</h2>
        ''' + rescuetime_table + '''
    </body>
</html>'''


# In[72]:


#f = open('report.html','w')
#f.write(html_string)
#f.close()


# In[73]:


from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# me == my email address
# you == recipient's email address
me = "[MyEmailAddress]@gmail.com"
you = "[MyEmailAddress]@gmail.com"


# Create message container - the correct MIME type is multipart/alternative.
msg = MIMEMultipart('alternative')
msg['Subject'] = '[MyName]\'s Daily Journal Items for '+agenda_date
msg['From'] = me
msg['To'] = you

# Create the body of the message (a plain-text and an HTML version).
text = 'htmlFailed'
html = html_string

# Record the MIME types of both parts - text/plain and text/html.
part1 = MIMEText(text.encode('utf-8'), 'plain', 'utf-8')
part2 = MIMEText(html.encode('utf-8'), 'html', 'utf-8')

# Attach parts into message container.
# According to RFC 2046, the last part of a multipart message, in this case
# the HTML message, is best and preferred.
msg.attach(part1)
msg.attach(part2)

# Send the message via local SMTP server.
#s = smtplib.SMTP('localhost')
smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
smtpObj.ehlo()
smtpObj.starttls()
smtpObj.login('[MyEmail]@gmail.com', '[App-specific password]')

# sendmail function takes 3 arguments: sender's address, recipient's address
# and message to send - here it is sent as one string.
smtpObj.sendmail(me, you, msg.as_string())
smtpObj.quit()
