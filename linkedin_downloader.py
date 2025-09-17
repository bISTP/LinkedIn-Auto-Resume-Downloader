import os
import pickle
import time
import shutil
import pandas as pd
from random import uniform
import re
import base64
from datetime import datetime
import requests  # Modern alternative to wget
from dateutil.parser import parse
from pytz import timezone
from math import ceil
from bs4 import BeautifulSoup
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
import logging

# --- Configuration ---
CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.modify'] # Readonly is sufficient for downloading, modify for marking as read if desired
TOKEN_PICKLE_FILE = 'token.pickle'
LOG_FILE_NAME = 'log.xlsx'
INFO_FILE_NAME = 'info.txt'
DOWNLOAD_DIR = 'Downloaded_Resumes' # Changed from 'Downloaded_Files' for clarity
TIMEZONE = 'Asia/Kolkata' # Define timezone centrally

# --- Logging Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Helper Functions ---
def get_gmail_service():
    """Authenticates and returns the Gmail API service."""
    creds = None
    if os.path.exists(TOKEN_PICKLE_FILE):
        with open(TOKEN_PICKLE_FILE, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError:
                logging.error("Failed to refresh token. Re-authenticating.")
                creds = None # Force re-authentication
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PICKLE_FILE, 'wb') as token:
            pickle.dump(creds, token)
    return build(API_NAME, API_VERSION, credentials=creds)

def get_most_recent_download_date():
    """Reads the last successful download date from info.txt."""
    if os.path.exists(INFO_FILE_NAME):
        with open(INFO_FILE_NAME, 'r') as f:
            date_str = f.read().strip()
            if date_str:
                return datetime.strptime(date_str, "%d %m %Y %H").replace(tzinfo=timezone(TIMEZONE))
    return None

def update_most_recent_download_date(dt_obj):
    """Writes the current run's end date to info.txt."""
    with open(INFO_FILE_NAME, 'w') as f:
        f.write(dt_obj.strftime("%d %m %Y %H"))

def load_or_create_log(log_path):
    """Loads the log DataFrame or creates a new one."""
    if os.path.exists(log_path):
        return pd.read_excel(log_path)
    else:
        return pd.DataFrame(columns=['Thread ID', 'Date Received', 'File Name', 'Job Post'])

def parse_email_date(date_header_value):
    """Parses the date string from email headers."""
    # Example format: 'Mon, 13 Nov 2023 10:30:00 +0530 (IST)'
    try:
        # Split at the first semicolon to isolate the relevant part
        date_str_part = date_header_value.split(';')[0].strip()
        # Handle cases where timezone might be in parentheses or at the end
        if '(' in date_str_part and ')' in date_str_part:
            date_str_part = re.sub(r'\s*\([^)]*\)\s*', '', date_str_part).strip() # Remove text in parentheses

        dt_obj = parse(date_str_part).astimezone(timezone(TIMEZONE))
        return dt_obj
    except Exception as e:
        logging.warning(f"Could not parse date '{date_header_value}': {e}")
        return None

def download_file(url, destination_folder):
    """Downloads a file from a URL using requests."""
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status() # Raise an HTTPError for bad responses (4xx or 5xx)

        # Try to get filename from headers, otherwise from URL
        if 'content-disposition' in response.headers:
            fname = re.findall('filename="(.+)"', response.headers['content-disposition'])
            file_name = fname[0] if fname else url.split('/')[-1]
        else:
            file_name = url.split('/')[-1]

        # Clean up filename (e.g., remove query parameters if any)
        file_name = file_name.split('?')[0]

        destination_path = os.path.join(destination_folder, file_name)

        with open(destination_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        return destination_path
    except requests.exceptions.RequestException as e:
        logging.error(f"Error downloading file from {url}: {e}")
        return None

# --- Main Logic ---
def main():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    # 1. Authentication
    service = get_gmail_service()
    logging.info("Gmail service authenticated.")

    # 2. Load Log
    log_df = load_or_create_log(LOG_FILE_NAME)
    logging.info(f"Loaded {len(log_df)} entries from {LOG_FILE_NAME}.")

    # 3. Get Date Range
    input_date_str = input(f'Enter Start Date in dd mm yyyy hh (e.g., 01 01 2023 00). Leave blank to use last run date ({get_most_recent_download_date() if get_most_recent_download_date() else "N/A"}): ')
    output_date_str = input('Enter End Date in dd mm yyyy hh (Leave blank for current date and time): ')

    if not input_date_str:
        start_date = get_most_recent_download_date()
        if start_date is None:
            logging.error("No start date provided and no previous run date found. Exiting.")
            return
        logging.info(f"Using last run date as start date: {start_date.strftime('%d %m %Y %H')}")
    else:
        try:
            start_date = datetime.strptime(input_date_str, "%d %m %Y %H").replace(tzinfo=timezone(TIMEZONE))
        except ValueError:
            logging.error("Invalid start date format. Please use dd mm yyyy hh.")
            return

    if not output_date_str:
        end_date = datetime.now(timezone(TIMEZONE))
        logging.info(f"Using current date and time as end date: {end_date.strftime('%d %m %Y %H')}")
    else:
        try:
            end_date = datetime.strptime(output_date_str, "%d %m %Y %H").replace(tzinfo=timezone(TIMEZONE))
        except ValueError:
            logging.error("Invalid end date format. Please use dd mm yyyy hh.")
            return

    if start_date >= end_date:
        logging.error("Start date cannot be greater than or equal to end date. Exiting.")
        return

    logging.info(f"Processing emails from {start_date.strftime('%d %B %Y %H')} to {end_date.strftime('%d %B %Y %H')}.")

    # 4. Check for existing downloads in the range
    log_df['Date Received'] = pd.to_datetime(log_df['Date Received'], errors='coerce').dt.tz_localize(TIMEZONE, errors='coerce')
    existing_downloads_in_range = log_df[(log_df['Date Received'] >= start_date) & (log_df['Date Received'] <= end_date)]

    if not existing_downloads_in_range.empty:
        max_logged_date = existing_downloads_in_range['Date Received'].max().strftime('%d %B %Y %H')
        print(f'''Resume(s) have been downloaded up to {max_logged_date}.
There are {len(existing_downloads_in_range)} resumes between {start_date.strftime('%d %B %Y %H')} and {end_date.strftime('%d %B %Y %H')}.
Do you want to re-download these?''')

        is_drop = input('Enter Y for Yes (re-download) and N for No (skip): ').lower()
        if is_drop == 'y':
            log_df = log_df[~((log_df['Date Received'] >= start_date) & (log_df['Date Received'] <= end_date))].reset_index(drop=True)
            logging.info(f"Dropped {len(existing_downloads_in_range)} entries from log for re-download.")
        else:
            logging.info("Skipping re-download. Exiting.")
            return

    # 5. Fetch Gmail Threads
    downloaded_count = 0
    page_token = None
    query = f"from:jobs-listings@linkedin.com after:{int(start_date.timestamp())} before:{int(end_date.timestamp())}"
    logging.info(f"Gmail API query: {query}")

    while True:
        try:
            threads_response = service.users().threads().list(
                userId='me',
                maxResults=500,
                q=query,
                pageToken=page_token
            ).execute()
        except RefreshError:
            logging.warning("Token expired during fetch, attempting re-authentication.")
            service = get_gmail_service()
            threads_response = service.users().threads().list(
                userId='me',
                maxResults=500,
                q=query,
                pageToken=page_token
            ).execute()
        except Exception as e:
            logging.error(f"Error fetching threads from Gmail API: {e}")
            break

        threads = threads_response.get('threads', [])
        if not threads:
            logging.info("No more threads found or all threads processed.")
            break

        logging.info(f"Fetched {len(threads)} threads. Total processed: {downloaded_count + len(threads)} (approx).")

        for thread_info in threads:
            thread_id = thread_info['id']

            # Skip if already logged and not chosen for re-download
            if thread_id in log_df['Thread ID'].values and is_drop != 'y':
                logging.debug(f"Thread {thread_id} already logged. Skipping.")
                continue

            try:
                message = service.users().messages().get(userId='me', id=thread_id, format='full').execute()
                headers = message['payload']['headers']
                date_header = next((h['value'] for h in headers if h['name'].lower() == 'date'), None)
                
                if not date_header:
                    logging.warning(f"Could not find 'Date' header for thread {thread_id}. Skipping.")
                    continue

                email_date = parse_email_date(date_header)

                if not email_date:
                    logging.warning(f"Skipping thread {thread_id} due to unparsable date.")
                    continue

                # Filter by date range (already done in query, but good to double check)
                if not (start_date <= email_date <= end_date):
                    logging.debug(f"Thread {thread_id} date {email_date} is outside the specified range. Skipping.")
                    continue

                # Extract email content
                # This part is highly dependent on LinkedIn's email structure.
                # The original code assumes payload['parts'][1]['body']['data'].
                # A more robust approach might iterate through parts or check MIME types.
                email_body_data = None
                if 'parts' in message['payload']:
                    for part in message['payload']['parts']:
                        if part['mimeType'] == 'text/html' and 'body' in part and 'data' in part['body']:
                            email_body_data = part['body']['data']
                            break
                        elif part['mimeType'] == 'multipart/alternative': # Common for HTML/Plaintext emails
                            for sub_part in part.get('parts', []):
                                if sub_part['mimeType'] == 'text/html' and 'body' in sub_part and 'data' in sub_part['body']:
                                    email_body_data = sub_part['body']['data']
                                    break
                        if email_body_data:
                            break
                elif 'body' in message['payload'] and 'data' in message['payload']['body']:
                    email_body_data = message['payload']['body']['data']


                if not email_body_data:
                    logging.warning(f"Could not find email body data for thread {thread_id}. Skipping.")
                    continue

                decoded_data = base64.urlsafe_b64decode(email_body_data).decode('utf-8', errors='ignore')
                soup = BeautifulSoup(decoded_data, 'html.parser')

                # Find resume download link
                resume_url = None
                for link in soup.find_all('a', href=True):
                    if 'download_resume' in link.get('href', ''):
                        resume_url = link['href']
                        break

                if not resume_url:
                    logging.warning(f"No resume download link found for thread {thread_id}. Skipping.")
                    continue

                # Extract Job Post Name from snippet
                snippet = message.get('snippet', '')
                job_post = "Unknown_Job_Post"
                match = re.search(r'has a new applicant for (.*?)(?: \(.+\))?,', snippet) # More robust regex
                if match:
                    job_post_raw = match.group(1).strip()
                    job_post = re.sub(r'[^\w\-_\. ]', '', job_post_raw).replace(' ', '_') # Sanitize for folder name
                
                destination_folder = os.path.join(DOWNLOAD_DIR, job_post)
                os.makedirs(destination_folder, exist_ok=True)

                # Download file
                downloaded_file_path = download_file(resume_url, destination_folder)

                if downloaded_file_path:
                    filename = os.path.basename(downloaded_file_path)
                    # Add to log
                    new_entry = pd.DataFrame([{
                        'Thread ID': thread_id,
                        'Date Received': email_date.strftime("%Y-%m-%d %H:%M:%S%z"), # ISO format with timezone
                        'File Name': filename,
                        'Job Post': job_post
                    }])
                    log_df = pd.concat([log_df, new_entry], ignore_index=True)
                    downloaded_count += 1
                    logging.info(f"{downloaded_count}. Downloaded '{filename}' for '{job_post}' to '{destination_folder}'")
                    time.sleep(uniform(0.5, 1.5)) # Be polite to servers
                else:
                    logging.warning(f"Failed to download resume from {resume_url} for thread {thread_id}.")

            except Exception as e:
                logging.error(f"Error processing thread {thread_id}: {e}", exc_info=True)
                # Continue to next thread even if one fails

        page_token = threads_response.get('nextPageToken', None)
        if not page_token:
            break

    logging.info(f"Completed! Downloaded {downloaded_count} new resume files.")

    # 6. Save Log and Update Info File
    log_df['Date Received'] = pd.to_datetime(log_df['Date Received']).dt.tz_convert(TIMEZONE).dt.strftime("%d-%m-%Y %H:%M") # Format for Excel
    log_df.to_excel(LOG_FILE_NAME, index=False)
    update_most_recent_download_date(end_date)
    logging.info(f"Log updated in '{LOG_FILE_NAME}'. Last run date saved in '{INFO_FILE_NAME}'.")

if __name__ == '__main__':
    main()
