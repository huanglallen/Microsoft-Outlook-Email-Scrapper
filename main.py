import extract_msg
import glob
import pandas as pd
import os
import re
from PyPDF2 import PdfReader
from openpyxl import load_workbook
from datetime import datetime

folder_path = '/mnt/c/TestFiles/'
msg_files = glob.glob(f'{folder_path}*.msg')
email_regex = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
email_data = []

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

def filter_emails(emails):
    keywords_to_filter = ['origene', 'support', 'sale', 'product', 'purchas', 'order', 'account', 'pay', 'bill', 'buy', 'track', 'team', 'custom', 'info', 'ship', 'suppl', 'invoic', 'help', 'admin', 'subscribe', 'reply', 'confirm', 'exped', 'procure', 'service', 'financ', 'trade', 'notif', 'communica', 'data', 'stock', 'contact', 'quote', 'market', 'vendor', 'po-', 'po_', 'po@', 'ap-', 'ap_', 'ap@']
    
    filtered_emails = [
        email for email in emails
        if not any(keyword in email.lower() for keyword in keywords_to_filter)
    ]
    return filtered_emails

def format_date(date_string):
    try:
        parsed_date = datetime.strptime(date_string, '%a, %d %b %Y %H:%M:%S %z')
        formatted_date = parsed_date.strftime('%I:%M %p, %d %b %Y')
        return formatted_date
    except ValueError as e:
        print(f"Error formatting date: {e}")
        return date_string

def main():
    for msg_file in msg_files:
        try:
            msg = extract_msg.Message(msg_file)
        except Exception as e:
            print(f"Error processing {msg_file}: {e}")
            continue
        
        body = msg.body
        if not isinstance(body, str):
            print(f"Skipping file {msg_file} due to invalid body content.")
            continue
        
        subject = msg.subject
        sender = msg.sender
        date = msg.date

        # Ensure the sender's email is extracted from the sender string
        sender_email = re.findall(email_regex, sender)
        if sender_email:
            sender_email = sender_email[0]
        else:
            sender_email = sender

        formatted_date = format_date(date)
        body_emails = re.findall(email_regex, body)
        
        # Extract emails from attachments (PDFs)
        attachment_emails = []
        if msg.attachments:
            for attachment in msg.attachments:
                if attachment.longFilename and attachment.longFilename.endswith('.pdf'):
                    attachment_path = os.path.join(folder_path, attachment.longFilename)
                    with open(attachment_path, 'wb') as f:
                        f.write(attachment.data)
                    attachment_emails += extract_emails_from_pdf(attachment_path)

        all_emails = body_emails + attachment_emails
        if sender_email and sender_email not in all_emails:
            all_emails.append(sender_email)

        # Apply the filter to the combined list of emails after adding the sender's email
        all_emails = filter_emails(all_emails)

        # Extract just the filename (without the path and extension)
        file_name = os.path.splitext(os.path.basename(msg_file))[0]

        email_data.append({
            'Date': formatted_date,
            'File': file_name,
            'Subject': subject,
            'Sender': sender,
            'Emails': "; ".join(set(all_emails))
        })

    # Create a Pandas DataFrame from the list
    df = pd.DataFrame(email_data)

    # Remove rows where the "Emails" column has 0 items
    df = df[df['Emails'].str.strip().astype(bool)]

    # Write DataFrame to Excel
    output_file = '/mnt/c/TestFiles/emails_output.xlsx'
    try:
        df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"Data has been written to {output_file}")
        
        # Load the workbook to adjust column widths
        wb = load_workbook(output_file)
        ws = wb.active

        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter 
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 1)
            ws.column_dimensions[column].width = adjusted_width

        wb.save(output_file)

    except PermissionError as e:
        print(f"Permission Error: {e}")
    except Exception as e:
        print(f"Error writing file: {e}")

if __name__ == "__main__":
    main()
