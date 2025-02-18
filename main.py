import extract_msg
import glob
import pandas as pd
import os  # Import os to manipulate file paths

# Define folder path
folder_path = '/mnt/c/TestFiles/'
msg_files = glob.glob(f'{folder_path}*.msg')

# List to store email data
email_data = []

# Loop through each .msg file
for msg_file in msg_files:
    # Load the .msg file
    msg = extract_msg.Message(msg_file)

    # Extract email details
    subject = msg.subject
    sender = msg.sender
    date = msg.date
    
    # Extract just the filename without the extension (e.g., 'SR132642')
    file_name = os.path.splitext(os.path.basename(msg_file))[0]
    
    # Append data as a dictionary to the list
    email_data.append({
        'File': file_name,  # Use the modified file name without extension
        'Subject': subject,
        'Sender': sender,
        'Date': date
    })

# Create a Pandas DataFrame from the list
df = pd.DataFrame(email_data)

# Write DataFrame to Excel
output_file = '/mnt/c/TestFiles/emails_output.xlsx'
df.to_excel(output_file, index=False, engine='openpyxl')

print(f"Data has been written to {output_file}")
