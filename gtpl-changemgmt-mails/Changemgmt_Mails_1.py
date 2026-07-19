import win32com.client
import pandas as pd
from bs4 import BeautifulSoup
import os
import time
import traceback
from pathlib import Path


# Start timing
start_time = time.time()

# Configuration
EXCEL_FILE_PATH = Path.cwd() / "Excel_File" / "Latest_B2B.xlsx"
CC_RECIPIENTS = "'Tilok Choudhury' <tilok.choudhury@gtpl.net>; 'Amit Patel' <amit.patel@gtpl.net>"

# Create Outlook application instance
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    folder = namespace.Folders["My Outlook Data File(1)"].Folders["Inbox"].Folders["Changemgmt_Airtel"]
except Exception as e:
    print(f"Error initializing Outlook: {e}")
    print(traceback.format_exc())  # Print the traceback
    exit()

# Load Excel files
try:
    df_identity_to_name = pd.read_excel(EXCEL_FILE_PATH, sheet_name="Lookup-Sheet", usecols=[3, 4, 5, 6])
    df_name_to_email = pd.read_excel(EXCEL_FILE_PATH, sheet_name="Mail-IDs", usecols=[0, 1])
except Exception as e:
    print(f"Error loading Excel files: {e}")
    print(traceback.format_exc())  # Print the traceback
    exit()

# Email Signature
# signature_folder = os.path.expanduser(r'~\\AppData\\Roaming\\Microsoft\\Signatures')
pwd = os.getcwd()
signature_file = "Vinayak.htm"
signature_path = os.path.join(pwd, "Sign", signature_file)

try:
    with open(signature_path, 'r', encoding='utf-8') as file:
        html_signature = file.read()
except Exception as e:
    print(f"Error reading signature file: {e}")
    print(traceback.format_exc())  # Print the traceback
    html_signature = "Thanks & Regards\n\nGTPL Broadband Pvt Ltd.\nEnterprise-NOC\nTelephone:-+91-7966779500\nEmail: support.ahmedabad@gtpl.net"  # Fallback to an empty signature

# Function to extract a section of text between two delimiters
def extract_section(body, start_delim, end_delim):
    start_index = body.find(start_delim)
    if start_index == -1:
        return None
    start_index += len(start_delim)
    end_index = body.find(end_delim, start_index)
    if end_index == -1:
        end_index = len(body)
    return body[start_index:end_index].strip()

# Function to send email
def send_email(subject, body, recipients):
    try:
        mail = outlook.CreateItem(0)  # 0: Mail Item
        mail.Subject = subject
        mail.HTMLBody = body
        mail.To = "; ".join(recipients)
        mail.CC = CC_RECIPIENTS
        mail.Importance = 2  # 2: High Importance
        mail.Send()
    except Exception as e:
        print(f"Error sending email to {recipients}: {e}")
        print(traceback.format_exc())  # Print the traceback

processed_emails = 0
sent_emails = 0

# Process each email in the folder
for item in folder.Items:
    try:
        if item.UnRead:  # Process only unread emails
            processed_emails += 1
            mail_body = item.Body
            start_greet = mail_body.index("With respect")
            end_greet = mail_body.index("Details of Maintenance")
            greet = mail_body[start_greet:end_greet-1]
            start_sub = item.Subject.index("CRQ")
            end_sub = item.Subject.index('Planned Maintenance activity')
            crq = item.Subject[start_sub:end_sub-1]

            if "Postponed" in greet or "Completed successfully" in greet:
                print("\nSkipped --> " + crq + "\n")
                print("======================================\n")
                item.UnRead = True
                item.Save()
                continue

            # Extract the identity table
            identity_table = extract_section(mail_body, "Impacted LSI : Party Name", "We request you")

            # Extract Maintenance Details Table
            email_body_table = item.HTMLBody
            soup = BeautifulSoup(email_body_table, 'html.parser')
            maintenance_table = soup.find('table')

            # Checking for Maintenance Table and Identity Table
            if not maintenance_table or not identity_table:
                item.UnRead = True
                item.Save()
                continue

            # Parse the identity table
            identity_lines = identity_table.split('\n')
            for line in identity_lines:
                parts = line.split(":")
                if len(parts) != 2:
                    continue
                identity_number = parts[0].strip()

                # Check if the identity number starts with "1Old_" and remove it if present
                if identity_number.startswith("1Old_"):
                    identity_number = identity_number[5:]  # Remove "1Old_" prefix

                # Lookup company name in Excel sheet
                matched_row = df_identity_to_name[df_identity_to_name['AIRTEL Circuit ID'].astype(str) == identity_number]
                if matched_row.empty:
                    continue

                matched_ckt = matched_row['AIRTEL Circuit ID'].values[0]
                recipient_name = matched_row['Customer Name'].astype(str).values[0].strip()

                # Check if recipient_name is empty or NA
                if not recipient_name:
                    print(f"\nRecipient name is not found for matched CKT ID: {matched_ckt}.\n")
                    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                    continue

                # Convert the recipient_name to lower case for matching only
                recipient_name_lower = recipient_name.lower()

                df_name_to_email['Customer Name'] = df_name_to_email['Customer Name'].str.strip()
                matched_emails = df_name_to_email[df_name_to_email['Customer Name'].str.lower() == recipient_name_lower]

                # If recipient Email-ID not found
                if matched_emails['Email-ID'].isna().all():
                    print(f"\nCKT ID: {matched_ckt} with Customer Name: {recipient_name} has no Email-ID in 'Mail-IDs' Sheet.\n")
                    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                    continue

                # Making list of matched_emails
                email_list = matched_emails['Email-ID'].tolist()

                # Making string from Locations of matched_ckt
                a_end = matched_row["Location From"].values[0]
                b_end = matched_row["Location To"].values[0]
                if pd.isna(a_end) or a_end == '' or a_end == 'NA':
                    a_end = None
                location = f'{b_end} ILL' if a_end is None else f'{a_end} to {b_end}'

                if email_list:
                    # Construct email body
                    email_body = f"""
                    <p style="color: #2F5496;"><b>Dear Team,</b></p>
                    <p style="color: #2F5496;"><b>Below link(s) shall go down due to planned maintenance activity.</b></p>
                    <p style="color: #2F5496;"><b>Please find below activity details and consider the link down time as mentioned.</b></p>
                    <p style="color: #2F5496;"><b>CKT ID: {matched_ckt}</b></p>
                    <p style="color: #2F5496;"><b>Location: {location}</b></p>
                    <br>
                    {maintenance_table}
                    <br><br>
                    {html_signature}"""

                    # Send Email with exception handling for each matched_ckt
                    try:
                        send_email(f"GTPL || CKT ID: {matched_ckt} || Location: {location} || PLANNED ACTIVITY", email_body, email_list)
                        sent_emails += 1
                        print(f"\nCKT ID: {matched_ckt} --> {crq}")
                        print(F"Customer Name: {recipient_name}")
                        print(f"Location: {location}")
                        print(f"Mail sent to {recipient_name} at {email_list}")
                    except Exception as e:
                        print(f"Failed to send email for CKT ID {matched_ckt}: {e}")
                        print(traceback.format_exc())  # Print the traceback

                    print("-------------------------------------------------------")
                    print("-------------------------------------------------------")

            item.UnRead = False
            item.Save()
    except Exception as e:
        print(f"Error processing email: {e}")
        print(traceback.format_exc())  # Print the traceback

# End timing
end_time = time.time()
execution_time = end_time - start_time

print(f"\n{processed_emails} Emails processed successfully.\n")
print(f"\n{sent_emails} Emails sent successfully.\n")
print(f"\nTotal execution time: {execution_time:.2f} seconds\n")



