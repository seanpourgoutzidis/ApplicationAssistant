#IMPORT STATEMENTS
import gspread #Import gspread to work with Google Sheets
import imaplib #For use reading my inbox
import email
from email.header import decode_header
import webbrowser
import os
import http.client, urllib
import re

#Global Variables
pattern_uid = re.compile(r'\d+ \(UID (?P<uid>\d+)\)')

#Gets the list of employers from your application spreadsheet
def getEmployers ():
    gc = gspread.service_account() #Get service account that has been set up

    sh = gc.open("NAME OF YOUR SPREADSHEET")

    spreadsheetData = sh.sheet1.get_all_records()

    employers = []
    
    employerRowDict = {}
    employerInterviewResponseDict = {}
    employerOfferResponseDict = {}
    
    #It starts reading from the second row in the file
    rowNum = 2

    #Get a list of all the emplyers
    for row in spreadsheetData:
        name = row['COLUMN WHERE YOU STORE EMPLOYER NAMES']
        
        employerRowDict[row['COLUMN WHERE YOU STORE EMPLOYER NAMES']] = rowNum
        employerInterviewResponseDict[row['COLUMN WHERE YOU STORE EMPLOYER NAMES']] = row['COLUMN WHERE YOU STORE INTERVIEW STATUS']
        employerOfferResponseDict[row['COLUMN WHERE YOU STORE EMPLOYER NAMES']] = row['COLUMN WHERE YOU STORE OFFER STATUS']
        
        if (" " in name):
            name = name.split(' ')
        
        employers.append(name)
        rowNum += 1
        
    return employers, employerRowDict, employerInterviewResponseDict, employerOfferResponseDict

#ApplicationAssistant App Password for Yahoo: zflrozxeizsabncd

#Reads your inbox and compares to given list of employers it should look for
def readEmails(employers, employerInterviewResponseDict, emailOfferResponseDict):

    emailFound = False
    employersFound = set()

    #Credentials for logging into email
    user = 'YOUR EMAIL'
    password = 'YOUR APP PASSWORD'

    #Sign into email
    M = imaplib.IMAP4_SSL('imap.mail.yahoo.com', 993) #OR WHATEVER EMAIL YOU USE! I KNOW GMAIL AND OUTLOOK WORK AS WELL!
    M.login(user,password)
    
    #Select the inbox
    status,messages = M.select(mailbox = 'Inbox', readonly = False)
    
    #Get the email ids of the inbox
    resp, items = M.search(None, 'All')
    email_ids  = items[0].split()

    numEmailsToCheck = 5

    # total number of emails
    messages = int(messages[0])
    
    #This represents the body of the email
    body = ""

    #Code adapted from: https://www.thepythoncode.com/code/reading-emails-in-python
    for i in range(messages, messages-numEmailsToCheck, -1):
        # fetch the email message by ID
        res, msg = M.fetch(str(i), "(RFC822)")
        
        #Was there an employer found in this iteration?
        employerFoundThisIteration = False
        
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                try:
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        # if it's a bytes, decode to str
                        subject = subject.decode(encoding)
                except:
                    subject = "(No Subject)"
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)

                # if the email message is multipart
                if msg.is_multipart():
                    
                    #Convert body to a list to store all the emails raw text
                    body = list()
                    
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body.append(part.get_payload(decode=True).decode())
                        except:
                            pass
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    
                #For multipart emails, join into one string
                if (type(body) == list):
                    body = " ".join(body)
                
                #Tracks the current iteration so that one email is not processed multiple times
                currIteration = -1
                
                #If the employer name is within the email Subject or From fields
                for employer in employers:
                    #For multi-word employer names
                    if (type(employer) == list):
                        for word in employer:
                            if word in subject or word in From:
                                emailFound = True
                                employerFoundThisIteration = True
                                name  = " ".join(employer)
                                employersFound.add(name)
                                
                                #If we have not received a conclusive interview response, update the interview response with this email
                                if (employerInterviewResponseDict[name] == '' or employerInterviewResponseDict[name] == '?'):
                                    employerInterviewResponseDict[name] = messageEvaluationInterview(body)
                                    currIteration = i
                                    
                                #If we have had an interview with the people with the employer, update the offer instead
                                elif (employerInterviewResponseDict[name] == 'Y' and (employerOfferResponseDict[name] == '?' or employerOfferResponseDict[name] == '') and i != currIteration):
                                    employerOfferResponseDict[name] = messageEvaluationOffer(body)

                    #For single word employer names
                    else:
                        if employer in subject or employer in From:
                            emailFound = True
                            employerFoundThisIteration = True
                            employersFound.add(employer)
                            
                            #If we have not received a conclusive interview response, update the interview response with this email
                            if (employerInterviewResponseDict[employer] == '' or employerInterviewResponseDict[employer] == '?'):
                                    employerInterviewResponseDict[employer] = messageEvaluationInterview(body)
                            #If we have had an interview with the people with the employer, update the offer instead
                            elif (employerInterviewResponseDict[employer] == 'Y' and (employerOfferResponseDict[employer] == '?' or employerOfferResponseDict[employer] == '')):
                                    employerOfferResponseDict[employer] = messageEvaluationOffer(body)
        
        #If we found an employer
        if (employerFoundThisIteration):
            
            
            #Code Adapted from: https://stackoverflow.com/questions/3527933/move-an-email-in-gmail-with-python-and-imaplib
            
            #Move the email to our specific applications folder
            latest_email_id = email_ids[i - 1]
            resp, data = M.fetch(latest_email_id, "(UID)")
            match = pattern_uid.match(data[0].decode("utf-8"))
            msg_uid = match.group('uid')
            result = M.uid('COPY', msg_uid, '/FOLDER WHERE YOU WANT TO STORE YOUR APPLICATIONS')

            #Remove the email from our inbox
            if result[0] == 'OK':
                mov, data = M.uid('STORE', msg_uid , '+FLAGS', '(\Deleted)')
                M.expunge()
            

    # close the connection and logout
    M.close()
    M.logout()
    
    return emailFound,employersFound,employerInterviewResponseDict,employerOfferResponseDict

#Evaluate the message to see if we have been offered an interview, declined or unsure
def messageEvaluationInterview (body):
    positiveInterview = {'interview', 'meet', 'schedule', 'move forward', 'accept' , 'selected', 'get to know', 'next stage'}
    negativeInterview = {'move in another direction', 'not move forward', 'not to move forward', 'reject', 'decline', 'not selected', 'proceed without', 'filled the position', 'position is filled'}
    
    selectedForInterview = '?'
    
    for keyTerm in positiveInterview:
        if keyTerm in body:
            selectedForInterview = 'Y'
            
    for keyTerm in negativeInterview:
        if keyTerm in body:
            selectedForInterview = 'N'
            
    return selectedForInterview

#Evaluate the message to see if we have been accepted, declined or unsure
def messageEvaluationOffer (body):
    positiveOffer = {'pleased to offer', 'belong', 'welcome', 'love to', 'have been chosen', 'accept'}
    negativeOffer = {'decline', 'reject', 'not offer', 'move in another direction', 'not move forward', 'not selected', 'proceed without', 'position is filled'}
    
    selectedForPosition = '?'
    
    for keyTerm in positiveOffer:
        if keyTerm in body:
            selectedForPosition = 'Y'
            
    for keyTerm in negativeOffer:
        if keyTerm in body:
            selectedForPosition = 'N'
            
    return selectedForPosition

#Update our spreadsheet with what we have found with our emails
def updateTracker (employersFound, employerRowDict, employerInterviewResponseDict, employerOfferResponseDict):
    gc = gspread.service_account() #Get service account that has been set up

    sh = gc.open("Application Tracker (ApplicationAssistant)")
    
    #Update the relevant cells
    for employer in employersFound:
        #CHANGE LETTER TO WHAT THE COLUMN THAT STORES THE INTERVIEW STATUS CORRESPONDS TO IN YOUR TABLE
        cell = 'D' + str(employerRowDict[employer])
        sh.sheet1.update(cell, employerInterviewResponseDict[employer]);
        
        if (employerInterviewResponseDict[employer] == 'Y'):
            #CHANGE LETTER TO WHAT THE COLUMN THAT STORES THE INTERVIEW DATE CORRESPONDS TO IN YOUR TABLE
            cell = 'E' + str(employerRowDict[employer])
            sh.sheet1.update(cell, 'SEE EMAIL!!!');
        
        #CHANGE LETTER TO WHAT THE COLUMN THAT STORES THE OFFER STATUS CORRESPONDS TO IN YOUR TABLE
        cell = 'F' + str(employerRowDict[employer])
        sh.sheet1.update(cell, employerOfferResponseDict[employer]);

#Send an SMS Notification to the PushOver application
def sendMessage (message):
    conn = http.client.HTTPSConnection("api.pushover.net:443")
    conn.request("POST", "/1/messages.json",
      urllib.parse.urlencode({
        "token": "YOUR APPLICATION TOKEN FROM PUSHOVER",
        "user": "YOUR USER CODE FROM PUSHOVER",
        "message": message,
      }), { "Content-type": "application/x-www-form-urlencoded" })
    conn.getresponse()

#Construct text message to send
def constructMessage (employersFound, employerInterviewResponseDict, employerOfferResponseDict):
    
    #Convert set to list
    employersFound = list(employersFound)
    
    #Construct the message
    message = "[ApplicationAssistant] You have recently received a response to your applications to: "
    
    for num in range(len(employersFound)):
        
        #Modifier stores if the response was positive, negative or unsure
        modifier = '(?)'
        
        #Set the modifier
        if (employerInterviewResponseDict[employersFound[num]] == 'Y'):
            if (employerOfferResponseDict[employersFound[num]] == 'Y' or employerOfferResponseDict[employersFound[num]] == ''):
                modifier = '(+)'
            elif (employerOfferResponseDict[employersFound[num]] == 'N'):
                modifier = '(-)'
            else:
                modifier = '(?)'
        elif (employerInterviewResponseDict[employersFound[num]] == 'N'):
            modifier = '(-)'
        
        #Assemble the message
        if (len(employersFound) == 1):
            message = message + employersFound[num] + modifier
        elif (num < (len(employersFound) - 1)):
            message = message + employersFound[num] + modifier +", "
        else:
            message = message + "and " + employersFound[num]
    
    message = message + "!"
    
    return message

#'Main Function'
if __name__ == '__main__':
    
    #Read our application tracking spreadsheet
    employers, employerRowDict, employerInterviewResponseDict, employerOfferResponseDict = getEmployers()
    
    #Read my inbox
    sendText, employersFound, employerInterviewResponseDict, employerOfferResponseDict = readEmails(employers, employerInterviewResponseDict, employerOfferResponseDict)
    
    #Update our tracking spreadsheet after cross-referencing with inbox
    updateTracker(employersFound, employerRowDict, employerInterviewResponseDict, employerOfferResponseDict)
    
    #If we have received a message from an employer, send push notification to my phone
    if sendText:
        sendMessage(constructMessage(employersFound, employerInterviewResponseDict, employerOfferResponseDict))
