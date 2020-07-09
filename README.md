![Working with emails using Python](/emails.JPG)


# Working with emails using Python
Send, receive, read, download emails and its contents using python


## Sending emails
| Service provider 	| Host 	| Port 	|
|:-:	|:-:	|:-:	|
| office.com 	| smtp.office365.com 	| 587 	|
| gmail.com 	| smtp.gmail.com 	| 587 	|

## Libraries

```python
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
```

---

## Reading emails
| Service provider 	| Host 	| Port 	|
|:-:	|:-:	|:-:	|
| office.com | outlook.office365.com | 993 |
| gmail.com | imap.gmail.com | 993 |

## Libraries

```python
import os
import time
import imaplib
import email
from email.header import decode_header
import webbrowser
```
