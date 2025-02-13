import os
import extract_msg

# msg = extract_msg.Message("path_to_your_msg_file.msg")
msg = extract_msg.Message('/mnt/c/TestFiles/SR132650.msg')

# Subject of the email
subject = msg.subject
print(f"Subject: {subject}")

# Sender's email address
sender = msg.sender
print(f"Sender: {sender}")

# Recipient's email address
recipient = msg.to
print(f"Recipient: {recipient}")

# Date the email was sent
date = msg.date
print(f"Date: {date}")

# Email body (HTML or plain text)
body = msg.body
print(f"Body: {body}")