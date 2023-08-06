# ApplicationAssistant
A bot meant to aid in applications as it cross-references a tracking application with your inbox, automatically updating the spreadsheet and sending you a Push notification to your phone! See my demo [here!](https://www.youtube.com/watch?v=MTVr7DEyxlI)

# INTRODUCTION

I am quite excited about this project and seeing how its been useful for me, I’d like to show you how you can set it up on your own computer!

For reference, I am using python3 on a Macbook with Monterey OS so the upcoming tutorial will cater to how I have set it up for myself. I believe it is still possible to use with differing setups, you may just have to stray from these instructions.

Prerequisites: Python3 installed on your computer, an email account and a Google Sheet that you would like to track. 
> [!NOTE]
> Your tracking sheet must have a column for: employers names, interview status, offer status and interview date.

# SET-UP

 Let’s get you set up with all the external APIs so that you can then just make a couple of edits to the source code.

First, we need to get you set up on **Google Cloud Platform**. 

* Create a [free account](https://www.youtube.com/watch?v=m5hwU0jD0qc)
* Create a [new project](​​https://developers.google.com/workspace/guides/create-project) 
* Fill in the relevant fields. Most importantly, list yourself as a test user to allow yourself to “test” your application.
* Search for the Google Drive API and enable it
* Search for the Google Sheets API and enable it.
* Create a service account for your [project](https://cloud.google.com/iam/docs/service-accounts-create)
* Now just make sure to share your Google Sheet that you are tracking with your service account so that it is able to interface with it.

Now, let’s get to started with the **gspread** library (pip install it first in your command line). Please follow the instructions under For Bots: Using a Service Account.

Now let’s get to set up with **PushOver** so that you can send push notifications to your phone. 
> [!NOTE]
> Unfortunately I was unable to find any completely free push notification services. I found that PushOver’s pricing scheme was the most fair with a one-time $5 USD fee to use it (and you get a one month trial to test it out).

* Sign up at [PushOver](https://pushover.net/)
* Download the PushOver app on your phone and sign into your account
* Register the application at [https://pushover.net/apps/build](https://pushover.net/apps/build)
* Copy down your api key and your user key for use later!

Now we are able to download and edit the source code found under **ApplicationAssistant_public.py**. 
> [!NOTE]
> Please install the python libraries found at the top of the file if you do not have them already (likely using pip install). Let’s edit the source code!

* Line 18: please put the name that your Google Sheet document is named (it can be in any folder in your Google Drive)
* Lines 33-37: please update the index ([]) with the appropriate names of your columns
* Lines 56 & 57: update with your email and password (note: you may need to create an app password for your email)
* Line 60: Replace imap.mail.yahoo.com with the [appropriate host name](https://knowledge.hubspot.com/email-notifications/how-can-i-find-my-email-servers-imap-and-smtp-information)
* Line 174: Update with the name of the folder in your inbox where you want to store the emails you have (personally, I just made a new folder for this)
  * This is so that you don’t keep getting notifications for emails that have already been processed but are still in your inbox
* Lines 231, 236 & 240: Change the letter in those lines to correspond
* Line 248: Update this line to have your api key for the project you made for PushOver
* Line 249: Update this line to have your user key from PushOver 

Following up on editing the source code, we use **Automator** to have the script run automatically on our Mac upon starting our computer.

* Open the Automator app
* Select Application
* Select Run shell script
* Add the lines (if you are using zsh terminal environment):
  * source $HOME/.zshrc
  * python <path>/<to>/ApplicationAssistant.py
* Save this app
* Go to your System Preferences, then Users and Group and then Login items
* Then add the new app you built.

That’s it! If you have any questions feel free to contact my through sean.pourgoutzidis@yahoo.ca or my [website](https://www.seanpourgoutzidis.com/) 
