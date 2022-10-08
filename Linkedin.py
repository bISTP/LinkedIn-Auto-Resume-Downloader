from sys import exit
import os
import pickle
import time
import shutil
import pandas as pd
from random import uniform
import re
import base64
from datetime import datetime
import wget
from datetime import datetime
from dateutil.parser import parse
from pytz import timezone
from math import ceil
from bs4 import BeautifulSoup
from urllib.error import HTTPError, URLError
from googleapiclient.errors import HttpError
from google.auth.exceptions import RefreshError

from Google import Create_Service

# To Build Service
CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://mail.google.com/']

# Authorization, Building Service and Dumping in a Pickle File.
try:
  service = pickle.load(open('service.pkl', 'rb'))
  service
except:
  service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
  with open(path + 'service.pkl', 'wb' ) as f:
    pickle.dump(service, f)
    f.close()

# Creating Log file and Reading it
if os.path.exists('log.xlsx'):
    log = pd.read_excel('log.xlsx')
else:
    log = pd.DataFrame(columns=['Thread ID', 'Date', 'Recieved', 'File Name', 'Job Post'])

# Start and End Date
input_date = input('Enter Start Date in dd mm yyyy hh Format: ')
output_date = input('Enter End Date in dd mm yyyy hh Format: ')

if input_date == '':
  with open('info.txt', 'r') as f:
      input_date = f.read()
      f.close()
else:
  pass

input_date = datetime.strptime(input_date, "%d %m %Y %H").replace(tzinfo=timezone('Asia/Kolkata'))

if output_date == '':
  output_date = datetime.today()
  output_date = output_date.astimezone(timezone('Asia/Kolkata'))
else:
  output_date = datetime.strptime(output_date, "%d %m %Y %H").replace(tzinfo=timezone('Asia/Kolkata'))

# Checking if Resumes already been downloaded for a given Time Window
log1 = log.copy()
log1['Date Recieved'] = log1['Date Recieved'].apply(lambda x: datetime.strptime(x, "%d-%m-%Y %H:%M").replace(tzinfo=timezone('Asia/Kolkata')))

index_to_drop = list(log1[(log1['Date Recieved'] >= input_date)&(log1['Date Recieved'] <= output_date)].index)

# What to do with them
if index_to_drop:
  print(f'''Resume Has been Downloaded upto {log1['Date Recieved'].max().strftime('%d %B %Y %H')} \n
            There are {len(index_to_drop)} Resumes between {input_date.strftime('%d %B %Y %H')} and {output_date.strftime('%d %B %Y %H')}\n
            Do you want to Drop the Log for {len(index_to_drop)} Resumes and Download them Again ?''')
  
  is_drop = input('Enter Y for Yes and N for No :')

  if is_drop.lower() == 'y':
    log.drop(index=index_to_drop, inplace=True)
    log = log.reset_index(drop=True)
  else:
    print('Exiting the Code...')
    exit()

else:
  pass

try:
  threads_list = service.users().threads().list(userId='me', maxResults=500, q="from:jobs-listings@linkedin.com").execute()
except RefreshError:
  service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
  with open('service.pkl', 'wb' ) as f:
    pickle.dump(service, f)
    f.close()

  threads_list = service.users().threads().list(userId='me', maxResults=500, q="from:jobs-listings@linkedin.com").execute()

try:
  nextPageToken = threads_list['nextPageToken']
  iter_counts = ceil(threads_list['resultSizeEstimate']/len(threads_list['threads']))
except KeyError:
  iter_counts = 1
  pass

count = 0
for i in range(iter_counts):
  #try:
  if i == 0:
    print(f'i = {i}, pagetoken = {nextPageToken}')

  else:
    threads_list = service.users().threads().list(userId='me', maxResults=500, q="from:jobs-listings@linkedin.com", pageToken=nextPageToken).execute()
    nextPageToken = threads_list['nextPageToken']
    print(f'i = {i}, pagetoken = {nextPageToken}')
  
  for thread in list(threads_list['threads']):
    thread_id = thread['id']
  
    try:

      message = service.users().messages().get(userId='me', id=thread_id).execute()
      date_text = message['payload']['headers'][1]['value'].split(';')[1].split(',')[1].strip()
      date_text = parse(date_text).astimezone(timezone('Asia/Kolkata'))
      date_text_str = date_text.strftime("%d-%m-%Y %H:00")

      if date_text > input_date and date_text < output_date:

        payload = message['payload']
        data = payload['parts'][1]['body']['data']
        data = data.replace("-","+").replace("_","/")
        decoded_data = base64.b64decode(data)
        soup = BeautifulSoup(decoded_data, 'html.parser')

        
        for link in soup.find_all('a'):
          if 'download_resume' in link.get('href'):
            url = link.get('href')
            break

        try:
          file_name = wget.download(url)

          time.sleep(uniform(1, 2))

          snippet = message['snippet']

          try:
            job_post = snippet[snippet.index(',')+2 : snippet.index(', has a new applicant')].replace('(', '').replace(')', '').replace(' ', '_')
            post_path = 'Downloaded_Files/' + job_post

            if os.path.exists(post_path):
              pass
            else:
              os.makedirs(post_path)
            
            for file in os.listdir():
              if file_name.lower() in file.lower():

                if file_name.lower() == file.lower():
                  pass
                elif '.pdf' not in file_name and '.doc' not in file_name:
                  file_name = file_name + file[len(file)-file[::-1].index('.')-1:]
                
                try:
                  shutil.move(file_name, post_path)
                  log = log.append({'Thread ID':thread_id, 'Date Recieved':date_text_str,	'File Name':file_name,	'Job Post':job_post}, ignore_index=True)

                  print(f'{count}. {file_name} : Downloaded in {job_post}')
                  count += 1
                except OSError:
                  pass
              else:
                pass
          except ValueError:
            pass
        except (HTTPError, URLError) as url_errors:
          pass
      else:
        break
    except HttpError:
      pass

print(f'Completed!!! {count} Files Downloaded.')
log.to_excel('log.xlsx', index=False)
with open('info.txt', 'w') as f:
  f.write(datetime.now().strftime("%d %m %Y %H"))
  f.close()