"""
Created on Wed Feb 27 13:07:52 2019

@author: Robert Piazza
"""

from __future__ import absolute_import, division, unicode_literals, print_function

import pandas as pd
import numpy as np

import datetime
from datetime import datetime as dt

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import requests

import smtplib

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']


#%% Dark Sky translations
weather_icons = {'Sun':'☀️', 
                 'Sun Behind Cloud':'⛅',
                 'Cloud With Lightning and Rain':'⛈️',
                 'Umbrella With Rain Drops':'☔',
                 'Cloud':'☁️',
                 '🌨️ Cloud With Snow': '🌨️ Cloud With Snow',
                 'Cloud With Lightning': '🌩️',
                 'Tornado': '🌪️',
                 'Snowflake':'❄️',
                 'Thermometer':'🌡️',
                 'Foggy': '🌁',
                 'Fog':'🌫️',
                 'High Voltage':'⚡',
                 'Sunset':'🌇',
                 'Sunrise':'🌅',
                 'Dashing Away':'💨',
                 'Night With Stars':'🌃',
                 'Full Moon': '🌕',
                 'New Moon': '🌑',
                 'Last Quarter Moon': '🌗',
                 'Waning Crescent Moon':'🌘',
                 'Waxing Crescent Moon':'🌒',
                 'Waxing Gibbous Moon':'🌔',
                 'First Quarter Moon':'🌓',
                 'Waning Gibbous Moon':'🌖'
                 }
moon_translate = {0:weather_icons['New Moon'],
                  0.125:weather_icons['Waxing Crescent Moon'],
                  0.25:weather_icons['First Quarter Moon'],
                  0.375:weather_icons['Waxing Gibbous Moon'],
                  0.5: weather_icons['Full Moon'],
                  0.625:weather_icons['Waning Gibbous Moon'],
                  0.75:weather_icons['Last Quarter Moon'],
                  0.875:weather_icons['Waning Crescent Moon'],
                  1.0:weather_icons['New Moon']
                  }
icon_translate={'clear-day': weather_icons['Sun'],
                'clear-night':weather_icons['Full Moon'], 
                'rain':weather_icons['Umbrella With Rain Drops'],
                'snow': weather_icons['Snowflake'],
                'sleet': weather_icons['Umbrella With Rain Drops'],
                'wind': weather_icons['Dashing Away'],
                'fog': weather_icons['Fog'],
                'cloudy': weather_icons['Cloud'],
                'partly-cloudy-day': weather_icons['Sun Behind Cloud'],
                'partly-cloudy-night': weather_icons['Sun Behind Cloud']
                }
#%% Weather

#Weather at home or work:
desired_location = 'Home'
locations_dict = {'Home': ['47.000000,-122.000000', 'America/Los_Angeles'],
             'Work': ['47.000000,-122.000000', 'America/Los_Angeles'],
             'Hawaii': ['21.000000,-157.000000', 'Pacific/Honolulu'],
             'Parents': ['34.000000,-117.000000', 'America/Los_Angeles']}
tz = locations_dict[desired_location][1]
darksky_token = '32characters_lowercaseletters&#s'
weather_site = 'https://api.darksky.net/forecast/'+darksky_token+'/'+locations_dict[desired_location][0]
location_weather = requests.get(weather_site).json()
daily_weather = location_weather['daily']['data'][0]
 
#This will become the full weather report text and we add as new information is pulled
weather_report = '<li>Next 7 days: '+location_weather['daily']['summary']+'\n</li>'

weather_day = dt.fromtimestamp(daily_weather['time']).strftime('%A')
agenda_date  = dt.fromtimestamp(daily_weather['time']).strftime('%a %b-%d')
weather_icon = icon_translate[daily_weather['icon']]


daily_summary = daily_weather['summary']+' '+weather_icon

weather_report += '<li>Today(%s), %s\n</li>' % (weather_day, daily_summary) 


#May not be any alerts but check for them
try:
    for alert in location_weather['alerts']:
        #print(alert['description'])
        weather_report += '<li>'+alert['description'] + '\n</li>'
except:
    #print('No Alerts')
    weather_report += '<li>No Alerts\n</li>'

#Create weather events table for incorporating into combined calendar event list
weather_times = {'Sunset':daily_weather['sunsetTime'],
                 'Sunrise':daily_weather['sunriseTime'],
                 'TempLow':daily_weather['temperatureLowTime'],
                 'TempHigh':daily_weather['temperatureHighTime']}
try:
    rain_event_chance = '%s%% chance of %s.' % (str(int(daily_weather['precipProbability']*100)),
                                                     daily_weather['precipType'])
    weather_times['Precip'] = daily_weather['precipIntensityMaxTime']
except:
    rain_event_chance = ''
#Create an ordered agenda of the weather events of the day instead of by type    

weather_event_text = {'Sunset':'🌇',
                      'Sunrise':'🌅',
                      'TempLow':'🌡️⬇️ of '+str(int(daily_weather['temperatureLow']))+'°F.',
                      'TempHigh':'🌡️⬆️ of '+str(int(daily_weather['temperatureHigh']))+ '°F.',
                      'Precip':rain_event_chance}
weather_ordered = sorted(weather_times, key=weather_times.get)

weather_event_list = []
for weather_event in weather_ordered:
    weather_time = pd.to_datetime(weather_times[weather_event], unit = 's').tz_localize('UTC').strftime('%Y-%m-%dT%H:%M:%SZ')
    weather_event_list += [{'summary':weather_event_text[weather_event], 'start': {'dateTime': weather_time}, 'timeZone':tz}]

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


desired_calendars = ['Todoist', '[myemailaddress]@gmail.com', 'Work', '[My Sports Team]']

# Call the Calendar API
service = build('calendar', 'v3', credentials=creds)
calendars = pd.DataFrame(service.calendarList().list(pageToken = None).execute()['items'])

now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
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
    #add time zone to each event for later use
    for event in events:
        event['timeZone']=events_result['timeZone']
        
    compiled_events = compiled_events.append(pd.DataFrame(events),sort = True)

#Add weather events to agenda
compiled_events = compiled_events.append(pd.DataFrame(weather_event_list),sort = True)
       
#reset index
compiled_events = compiled_events.reset_index(drop=True)

## extract times from event information    
for index, time in enumerate(compiled_events['start']):
    try:
        #if a date is given, it's for local time, adjust to standard UTC start
        event_start = (pd.to_datetime(time['date'])).tz_localize(compiled_events.loc[index,'timeZone'])
    except:
        event_start = (pd.to_datetime(time['dateTime'])).tz_localize('UTC')
        
    compiled_events.loc[index, 'newStart']=event_start
    compiled_events.loc[index, 'Time'] = event_start.tz_convert(tz).strftime('%I:%M %p')

sorted_events = compiled_events.sort_values('newStart') 
today = pd.to_datetime(pd.Timestamp("today").strftime("%m/%d/%Y")).tz_localize(tz)
time_range = pd.date_range(
        today, 
        today+pd.Timedelta(days=2), 
        freq = 'D')
todays_events_bool = np.logical_and(np.logical_and(sorted_events['newStart']<time_range[1],
                              sorted_events['newStart']>pd.Timestamp("today").tz_localize(tz)),
                              np.logical_not(sorted_events.duplicated(['newStart','summary'])))
tomorrows_events_bool = np.logical_and(np.logical_and(sorted_events['newStart']<time_range[2],
                              sorted_events['newStart']>time_range[1]),
                              np.logical_not(sorted_events.duplicated(['newStart','summary'])))

#Calendar Part

today_header = "Today, " + pd.Timestamp("today").strftime('%m.%d.%a:')
if not sorted_events.loc[todays_events_bool,['Time', 'summary']].empty:
    today_events = sorted_events.loc[todays_events_bool,['Time', 'summary']].to_html(index=False, header = False).replace('<table border="1" class="dataframe">','<table class="table table-striped">')

else:
    today_events = 'No Events'
    
tomorrow_header = "\nTomorrow," + (pd.Timestamp("today")+pd.Timedelta(hours=24)).strftime('%m.%d.%a:')
if not sorted_events.loc[tomorrows_events_bool,['Time', 'summary']].empty:
    tomorrow_events = sorted_events.loc[tomorrows_events_bool,['Time', 'summary']].to_html(index=False, header = False).replace('<table border="1" class="dataframe">','<table class="table table-striped">')

else:
    tomorrow_events = 'No Events'

#%%Todoist

todoist_token = '[lowercase letters and numbers obtained from having a premium account and going to https://todoist.com/prefs/integrations]'

#from todoist.api import TodoistAPI
import requests

#api = TodoistAPI(todoist_token)
#api.sync()
#print(api.state['projects'])

#projects = api.state['projects']
#items = api.state['items']
#labels = api.state['labels']
try:
    labels_JSON = requests.get("https://beta.todoist.com/API/v8/labels", 
                           headers={"Authorization": "Bearer %s" % todoist_token}).json()
    #write to a file in case further ones fail
    with open('Labels', 'wb') as fp:
        pickle.dump(labels_JSON, fp)
except:
    with open ('labels', 'rb') as fp:
        labels_JSON = pickle.load(fp)
    
labels = pd.DataFrame(labels_JSON)

try:
    projects_JSON = requests.get("https://beta.todoist.com/API/v8/projects",
                                params={
                                    #"project_id": 123
                                    #"label_id":179764
                                    #"filter": '(today|overdue) & @Work'
                                    #"lang": string
                                    
                                },
                                headers={
                                    "Authorization": "Bearer %s" % todoist_token
                                }).json()
    with open('projects', 'wb') as fp:
        pickle.dump(projects_JSON, fp)
except:
    with open ('projects', 'rb') as fp:
        projects_JSON = pickle.load(fp)
projects = pd.DataFrame(projects_JSON)
projects.columns = ['project_comment_count', 'project_id', 'project_indent', 'project_name', 'project_order']


work_tasks_no_id = pd.DataFrame(requests.get(
    "https://beta.todoist.com/API/v8/tasks",
    params={
        #"project_id": 123
        #"label_id":179764
        "filter": '(today|overdue) & @Work'
        #"lang": string
        
    },
    headers={
        "Authorization": "Bearer %s" % todoist_token
    }).json())

other_tasks_no_id = pd.DataFrame(requests.get(
    "https://beta.todoist.com/API/v8/tasks",
    params={
        #"project_id": 123
        #"label_id":179764
        "filter": '(today|overdue) & !@Work'
        #"lang": string
        
    },
    headers={
        "Authorization": "Bearer %s" % todoist_token
    }).json())

tomorrow_tasks_no_id = pd.DataFrame(requests.get(
    "https://beta.todoist.com/API/v8/tasks",
    params={
        #"project_id": 123
        #"label_id":179764
        "filter": "tomorrow"
        #"lang": string
        
    },
    headers={
        "Authorization": "Bearer %s" % todoist_token
    }).json())
#print('Due  for Work Today:\n')
if not work_tasks_no_id.empty:
    work_tasks = (work_tasks_no_id.join(
            projects.set_index('project_id'), 
            on = 'project_id').sort_values(
                    ['project_order', 'order'])).loc[:,['content', 'project_name']]
    #print(work_tasks)
    work_tasks_table = work_tasks.to_html(index=False, header = False).replace('<table border="1" class="dataframe">','<table class="table table-striped">')
else:
    #print('No Tasks for Work Today')
    work_tasks_table = 'No Tasks for Work Today'
#print('Other Tasks Due Today:\n')
if not other_tasks_no_id.empty:
    other_tasks = (other_tasks_no_id.join(
            projects.set_index('project_id'), 
            on = 'project_id').sort_values(
                    ['project_order', 'order'])).loc[:,['content', 'project_name']]
    #print(other_tasks)
    other_tasks_table = other_tasks.to_html(index=False, header = False).replace('<table border="1" class="dataframe">','<table class="table table-striped">')
else:
    #print('No Tasks elsewhere Today')
    other_tasks_table = 'No Tasks elsewhere Today'
    
#print('Tomorrow:\n')
if not tomorrow_tasks_no_id.empty:
    tomorrow_tasks = (tomorrow_tasks_no_id.join(
            projects.set_index('project_id'), 
            on = 'project_id').sort_values(
                    ['project_order', 'order'])).loc[:,['content', 'project_name']]
    #print(tomorrow_tasks)
    tomorrow_tasks_table = tomorrow_tasks.to_html(index=False, header = False).replace('<table border="1" class="dataframe">','<table class="table table-striped">')
else:
    #print('No Tasks Tomorrow')
    tomorrow_tasks_table = 'No Tasks Tomorrow'
    


# In[71]:


html_string = '''
<html>
    <head>
        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.1/css/bootstrap.min.css">
        <style>body{ margin:0 100; background:whitesmoke; }</style>
    </head>
    <body>
        <div class="page-header">
        <h1>Daily Agenda</h1>
        <p>Your daily agenda Robert</p>
        </div>
        <!-- *** Section 1 *** --->
        <div class="page-header">
            <h1>Weather '''+weather_icon+'''</h1>
        </div>        
        ''' + weather_report + '''
        <!-- *** Section 2 *** --->
        <div class="page-header">
            <h1> ''' + today_header + ''' </h1>
        </div>        
        ''' + today_events + '''
        <h3> Work Tasks: </h3>
        ''' + work_tasks_table + '''
        <br>
        ''' + moon_translate[round(daily_weather['moonPhase']*8)/8]+ '''
        </br>
        <div class="page-header">
            <h1> ''' + tomorrow_header + ''' </h1>
        </div>
        ''' + tomorrow_events + '''
        <div class="page-header">
            <h3> Tasks: </h3>
        </div>        
        ''' + tomorrow_tasks_table +'''
        <!-- *** Section 3 *** --->
        <h2> Other Tasks Today: </h2>
        ''' + other_tasks_table +'''

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
me = "[myemailaddress]@gmail.com"
you = "[myemailaddress]@gmail.com"

# Create message container - the correct MIME type is multipart/alternative.
msg = MIMEMultipart('alternative')
msg['Subject'] = '[MyName]\'s Daily Agenda for '+agenda_date
msg['From'] = me
msg['To'] = you

# Create the body of the message (a plain-text and an HTML version).
text = 'Daily Agenda'
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
smtpObj.login('[MyEmailAddress]@gmail.com', '[app-specific-password]')

# sendmail function takes 3 arguments: sender's address, recipient's address
# and message to send - here it is sent as one string.
smtpObj.sendmail(me, you, msg.as_string())
smtpObj.quit()

"""
On windows, I run Task Scheduler, 
Trigger: Daily 5:30 am 
Action: Start a Program: Program/Script: C:\ProgramData\Anaconda3\pythonw.exe
Add arguments: AgendaMaker.py
Start in: [directory that holds this file, for me it's C:\Dropbox\Python\Agenda]
