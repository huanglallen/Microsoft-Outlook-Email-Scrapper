import extract_msg
import glob
import pandas as pd
import os
import re
from PyPDF2 import PdfReader
from openpyxl import load_workbook

# Define folder path
folder_path = '/mnt/c/TestFiles/'
msg_files = glob.glob(f'{folder_path}*.msg')

# Regular expression to match emails
email_regex = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'

# Function to extract emails from a PDF file
def extract_emails_from_pdf(pdf_path):
    emails = []
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text = page.extract_text()
            if text:
                emails += re.findall(email_regex, text)
    except Exception as e:
        print(f"Error extracting emails from PDF {pdf_path}: {e}")
    return emails

# List to store email data
email_data = []

# Function to filter out emails from @origene.com
def filter_emails(emails):
    return [email for email in emails if '@origene.com' not in email.lower()]

# Loop through each .msg file
for msg_file in msg_files:
    # Load the .msg file
    try:
        msg = extract_msg.Message(msg_file)
    except Exception as e:
        print(f"Error processing {msg_file}: {e}")
        continue

    # Extract email details
    subject = msg.subject
    sender = msg.sender
    date = msg.date
    body = msg.body  # Extract email body text
    
    # Extract emails from the body of the email and filter out @origene.com
    body_emails = re.findall(email_regex, body)
    body_emails = filter_emails(body_emails)

    # Extract emails from attachments (PDFs) and filter out @origene.com
    attachment_emails = []
    if msg.attachments:
        for attachment in msg.attachments:
            if attachment.longFilename.endswith('.pdf'):
                # Save the attachment to the folder path
                attachment_path = os.path.join(folder_path, attachment.longFilename)
                with open(attachment_path, 'wb') as f:
                    f.write(attachment.data)  # Save the attachment data as a binary file
                attachment_emails += extract_emails_from_pdf(attachment_path)
        
        # Filter out @origene.com emails from the attachment emails
        attachment_emails = filter_emails(attachment_emails)
    
    # Combine emails from body and attachments
    all_emails = body_emails + attachment_emails

    # Extract just the filename without the extension (e.g., 'SR132642')
    file_name = os.path.splitext(os.path.basename(msg_file))[0]
    
    # Append data as a dictionary to the list
    email_data.append({
        'Date': date,
        'File': file_name,
        'Subject': subject,
        'Sender': sender,
        'Emails': "; ".join(set(all_emails))  # Join emails as a semicolon-separated string
    })

# Create a Pandas DataFrame from the list
df = pd.DataFrame(email_data)

# Write DataFrame to Excel
output_file = '/mnt/c/TestFiles/emails_output.xlsx'
try:
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Data has been written to {output_file}")
    
    # Load the workbook to adjust column widths
    wb = load_workbook(output_file)
    ws = wb.active

    # Iterate through each column to adjust the width
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name (e.g., 'A')
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)  # Add some padding
        ws.column_dimensions[column].width = adjusted_width

    # Save the workbook with adjusted column widths
    wb.save(output_file)

except PermissionError as e:
    print(f"Permission Error: {e}")
except Exception as e:
    print(f"Error writing file: {e}")
