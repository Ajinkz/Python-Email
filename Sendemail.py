#!/usr/bin/env python
# coding: utf-8

# Send emails using Python

import smtplib
import os
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


#sender
mid='sender@xyz.com'
pwd='xxxxxxxxxxxxxxxx'
#receiver
rec='receiver@xyz.com'
sender_email_address = mid
sender_email_password = pwd
receiver_email_address = rec


#subject line
email_subject_line = 'New test msg alert...!!!'


# Creates a multipart/* type message
msg = MIMEMultipart()
msg['From'] = sender_email_address
msg['To'] = receiver_email_address
msg['Subject'] = email_subject_line


# Create the body of the message (a plain-text and an HTML version).
text = "Hi! User😀, \nThis is plain text msg."
html = """<html>
  <head></head>
  <body>
  <hr>
    <p>Hi User,
     This is html based content<a href="http://ajinkz.github.io">link</a> .
     <img src="https://img.shields.io/badge/Test-Email-blue">
    </p>
    <hr>
  </body>
</html>
"""

# Record the MIME types of both parts - text/plain and text/html.
part1 = MIMEText(text, 'plain')
part2 = MIMEText(html, 'html')

# Attach parts into message container.
# According to RFC 2046, the last part of a multipart message, in this case
# the HTML message, is best and preferred.
msg.attach(part1)
msg.attach(part2)


#adding attachment to email
file_name = "banner.jpg"
part3 = MIMEBase('application', "octet-stream")
part3.set_payload(open(file_name, "rb").read())
encoders.encode_base64(part3)
part3.add_header('Content-Disposition', 'attachment;filename="banner.jpg"')
msg.attach(part3)


# Convert the msg data into string format
email_content = msg.as_string()


#Create TLS sesion and login into mail
server = smtplib.SMTP('smtp.office365.com:587') #gmail = "smtp.gmail.com:587"
server.starttls()
server.login(sender_email_address, sender_email_password)


# sending mail
server.sendmail(sender_email_address, receiver_email_address, email_content)
server.quit()



