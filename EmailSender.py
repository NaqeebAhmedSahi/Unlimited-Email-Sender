import smtplib
import os
import pandas as pd
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import re
# naqeebahmedsahi@gmail.com
#App pssword=gkrx faby icld zhif
# Function to check if an email is valid
def is_valid_email(email):
    pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return re.match(pattern, email) is not None

# Function to send email
def send_email(to_email, subject, body, from_email, password, attachment_path=None):
    try:
        # Create the email
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Attach file if specified
        if attachment_path:
            attachment_name = os.path.basename(attachment_path)
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(open(attachment_path, 'rb').read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={attachment_name}')
            msg.attach(part)

        # Connect to the SMTP server
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(from_email, password)
            server.send_message(msg)
        print(f"Email sent to {to_email}")
        return True
    except Exception as e:
        print(f"Failed to send email to {to_email}. Error: {e}")
        return False

# Function to process all Excel files in a folder
def process_all_excels_in_folder(folder_path, email_column_index, subject, body, from_email, password, attachment_path=None):
    # Path for JSON tracking file
    json_file_path = os.path.join(folder_path, 'email_tracking.json')

    # Load or initialize email tracking JSON file
    if os.path.exists(json_file_path):
        try:
            with open(json_file_path, 'r') as f:
                email_tracking = json.load(f)
        except json.JSONDecodeError:
            print("JSON file is empty or corrupted. Initializing with an empty tracking dictionary.")
            email_tracking = {}
    else:
        email_tracking = {}

    # Process each Excel file in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx'):
            excel_path = os.path.join(folder_path, file_name)
            print(f"Processing file: {file_name}")
            send_emails_from_excel(excel_path, email_column_index, subject, body, from_email, password, email_tracking, attachment_path)

    # Save email tracking status
    with open(json_file_path, 'w') as f:
        json.dump(email_tracking, f)

# Function to send emails from an Excel file
def send_emails_from_excel(excel_path, email_column_index, subject, body, from_email, password, email_tracking, attachment_path=None):
    df = pd.read_excel(excel_path)
    # Ensure the column index is within bounds
    if email_column_index >= len(df.columns):
        print(f"Column index {email_column_index} is out of bounds for file {excel_path}")
        return

    email_column = df.iloc[:, email_column_index]

    for email in email_column:
        if not is_valid_email(email):
            print(f"Invalid email address: {email}")
            continue

        # Check if the email has already been sent
        if email in email_tracking and email_tracking[email] == 'sent':
            print(f"Email already sent to: {email}")
            continue

        # Send email and update tracking
        if send_email(email, subject, body, from_email, password, attachment_path):
            email_tracking[email] = 'sent'
        else:
            email_tracking[email] = 'failed'

# Main function to get user input and start processing
def main():
    # Get user inputs
    choice = input("Press 1 to select the path manually or 2 to use the current folder: ")
    if choice == '1':
        folder_path = input("Enter the path to the folder containing Excel files: ")
    elif choice == '2':
        folder_path = os.getcwd()  # Current working directory
    else:
        print("Invalid choice.")
        return

    email_column_index = int(input("Enter the column number (0-based index) containing email addresses: "))
    from_email = input("Enter The Email : ")
    password = input("Enter The Email Password: ")
    subject = input("Enter The Email Subject: ")

    print("Enter The Email Body (Press Enter twice to finish):")
    body = ""
    while True:
        line = input()
        if line == "":
            break
        body += line + "\n"

    attachment_choice = input("Do you want to attach a file? (yes/no): ").strip().lower()
    if attachment_choice == 'yes':
        attachment_path = input("Enter the full path of the file to attach: ")
    else:
        attachment_path = None

    process_all_excels_in_folder(folder_path, email_column_index, subject, body.strip(), from_email, password, attachment_path)

if __name__ == "__main__":
    main()
