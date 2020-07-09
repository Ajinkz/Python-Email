#!/usr/bin/env python
# coding: utf-8

# Read mails from Inbox using Python
import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import time


# account credentials
username = input('Email: ')
password = input('Password: ')


# number of top emails to fetch
N = input("Number of top emails to fetch: ")
if N is None or N==0:
    N = 3


# create an IMAP4 class with SSL, use your email provider's IMAP server
imap = imaplib.IMAP4_SSL("outlook.office365.com") # gmail= "imap.gmail.com"
# authenticate
imap.login(username, password)


# select a mailbox (in this case, the inbox mailbox)
# use imap.list() to get the list of mailboxes
status, messages = imap.select("INBOX")


# total number of emails
messages = int(messages[0])
messages


# - Fetch the mails from inbox,
# - decode the email subject and body
# - Iterating over mail contemts, we decode the encoded data
# - create a new folder inside root directory with subject name and fetch attachments into it, if present
# 


for i in range(messages-4, messages-N-4, -1):
    # fetch the email message by ID
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # parse a bytes email into a message object
            msg = email.message_from_bytes(response[1])
            
            # decode the email subject
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                # if it's a bytes, decode to str
                subject = subject.decode()
            
            # email sender
            from_ = msg.get("From")
            print("Subject:", subject)
            print("From:", from_)
            
            # if the email message is multipart
            if msg.is_multipart():
               
                # iterate over email parts
                for part in msg.walk():
                    
                    # extract content type of email
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        
                        # print text/plain emails and skip attachments
                        print(body)
                    
                    elif "attachment" in content_disposition:
                        
                        # download attachment
                        filename = part.get_filename()
                        if filename:
                            if not os.path.isdir(subject):
                                
                                # make a folder for this email (named after the subject)
                                os.mkdir(subject)
                            filepath = os.path.join(subject, filename)
                            
                            # download attachment and save it
                            open(filepath, "wb").write(part.get_payload(decode=True))
            else:
                # extract content type of email
                content_type = msg.get_content_type()
                
                # get the email body
                body = msg.get_payload(decode=True).decode()
                if content_type == "text/plain":
                    
                    # print only text email parts
                    print(body)
            
            if content_type == "text/html":
                
                # if it's HTML, create a new HTML file and open it in browser
#                 if not os.path.isdir(subject):
                    
                # make a folder for this email (named after the subject)
                dirname = subject[:20] + str(round(time.time()))
                os.mkdir(dirname)

                filename = subject[:20] + str(round(time.time()))+ ".html"
                filepath = os.path.join(dirname, filename)
                print(filepath)
                
                # write the file
                open(filepath, "w").write(body)
                
                # open in the default browser
                webbrowser.open(filepath)

            print("="*100)


# close the connection and logout
imap.close()
imap.logout()
