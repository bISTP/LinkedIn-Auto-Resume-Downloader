# Linkedin-Auto-Resume-Downloader

This program automatically downloads the resume of candidates who applied to a specific job position posted by the recruiter on Linkedin using GmailAPI.
It downloads the Resumes and save them in different folder, job post wise, in case of multiple job postings by same recruiter.
It also maintains an Excel File Record for all the Resumes Downloaded.
*I didn't encapsulated the code inside functions but if anyone wants it, Let me know.*
Note: Linkedin might change its format, so if this doesn't work, let me know.

# Requirements
* You need to make sure that as a recruiter, you are getting mail from Linkedin every time someone applies for a job post.
* Enable GmailAPI from Google Developer Console using the same account you are recieving mail from Linkedin.
* Generate OAuth2 Credentials File and Rename it as clients_secret.json


 # Usage: There are two ways you can run this program:
 ## 1. Google Colab (Cloud):
 * Just create a Folder Named **Linkedin** inside the root directory of Google Drive.
 * Put client_secret.json and Google.py inside that folder.
 * Copy this [Colab File](https://colab.research.google.com/drive/14jm44-I8FlYbiafMlSdWVuXzAgA18Tzz?usp=sharing) into your drive.
 * Then just run all.
 * The Directory Should Look like this:
  .
  └── Linkedin/
      ├── client_secret.json
      └── Google.py

 ## 2. Using Local Machine
 * Just put Linkedin.py, client_secret.json and Google.py inside any folder.
 * Install the packages from requirements.txt or just `pip install -r requirements.txt`
 * Run Linkedion.py
* The Directory Should Look like this:
  .
  ├── client_secret.json
  ├── Google.py
  └── Linkedin.py
 
 # How it Works:
 * It will ask for authorization only the first time, A browser tab will be opened so make sure you are logged in with your main google account.
 * It has Two Inputs: Start Date and End Date. They both defines the time window for which you want to download the Resumes.
 * The format is "dd mm yyyy hh". For Example: 10 10 2022 14 means 10th October 2022 2PM.
 * If Start Date left empty, It will pick up the last date time you ran this program to ensure the continuity.
 * If End Date left empty, It will just take current date time.
 * The time zone used is 'Asia/Kolkata', you can replace it to your timezone by searching and replacing the string.
 * Note: There is a simpler inbuilt date filter in api request which can be given directly in the query but it is not precised to the "HOUR" like the one I used here. (if you need it, let me know)
 * Resumes are going to be downloaded inside Downloaded_Files Folder, separted by different sub folders based on the Job Post.
 * An excel file will be generated named *log.xlsx* which will keep all the records about the downloaded resumes.
 
