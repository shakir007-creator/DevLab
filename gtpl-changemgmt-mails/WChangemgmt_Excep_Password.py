# import win32com.client
# import pandas as pd
# from bs4 import BeautifulSoup
# import os
# import time
# import traceback
# import msoffcrypto
# import io
# import getpass
# from pathlib import Path

# # Default password if user input is empty or incorrect
# DEFAULT_PASSWORD = "admin@123"

# # Start timing
# start_time = time.time()

# # Configuration
# EXCEL_FILE_PATH = Path.cwd() / "Excel_File" / "Latest_B2B.xlsx"
# CC_RECIPIENTS = "'Tilok Choudhury' <tilok.choudhury@gtpl.net>; 'Amit Patel' <amit.patel@gtpl.net>;"

# # Function to check if the Excel file is password protected
# def is_password_protected(file_path):
#     """Check if the Excel file is password protected by attempting to load it with openpyxl."""
#     try:
#         pd.read_excel(file_path, sheet_name="Lookup-Sheet", usecols=[3, 4, 5, 6], engine="openpyxl")
#         return False  # No exception means the file is not password protected
#     except ValueError as e:
#         if "is not a zip file" in str(e):
#             return True  # Likely password protected
#         else:
#             raise e

# # Function to decrypt the Excel file if it is password protected
# def decrypt_excel(file_path, password):
#     """Decrypt the password-protected Excel file."""
#     try:
#         with open(file_path, "rb") as f:
#             office_file = msoffcrypto.OfficeFile(f)
#             office_file.load_key(password=password)  # Provide the password
#             decrypted = io.BytesIO()
#             office_file.decrypt(decrypted)  # Decrypt the file into a BytesIO object
#             decrypted.seek(0)  # Rewind the stream to the beginning
#         return decrypted
#     except msoffcrypto.exceptions.InvalidKeyError:
#         return None  # Invalid password
#     except Exception as e:
#         raise e

# # Function to read the Excel file (with or without password protection)
# def read_excel(file_path, password=None):
#     """Read the Excel file (decrypted or not) into a pandas DataFrame."""
#     try:
#         if password:
#             decrypted_file = decrypt_excel(file_path, password)
#             if decrypted_file is None:
#                 return None  # Password was incorrect
#             df = pd.read_excel(decrypted_file, engine="openpyxl")  # Load into pandas DataFrame
#         else:
#             df = pd.read_excel(file_path, engine="openpyxl")  # Load directly if not password-protected
#         return df
#     except Exception as e:
#         return None

# # Function to extract a section of text between two delimiters
# def extract_section(body, start_delim, end_delim):
#     start_index = body.find(start_delim)
#     if start_index == -1:
#         return None
#     start_index += len(start_delim)
#     end_index = body.find(end_delim, start_index)
#     if end_index == -1:
#         end_index = len(body)
#     return body[start_index:end_index].strip()

# # Function to send email
# def send_email(subject, body, recipients):
#     try:
#         mail = outlook.CreateItem(0)  # 0: Mail Item
#         mail.Subject = subject
#         mail.HTMLBody = body
#         mail.To = "; ".join(recipients)
#         mail.CC = CC_RECIPIENTS
#         mail.Importance = 2  # 2: High Importance
#         mail.Send()
#     except Exception as e:
#         print(f"Error sending email to {recipients}: {e}")
#         print(traceback.format_exc())  # Print the traceback

# # Function to get the password from the user (with retry mechanism)
# def get_password():
#     """Prompt the user for a password or use the default if input is empty or incorrect."""
#     password = getpass.getpass(prompt="Please enter the password for the Excel file (or press Enter to use the default password): ")
    
#     # If no password is provided, fall back to the default password
#     if not password:
#         password = DEFAULT_PASSWORD
    
#     return password

# # Get the current working directory
# current_directory = os.getcwd()

# # Full path to the file
# file_path = os.path.join(current_directory, EXCEL_FILE_PATH)

# # Get user password (with retry mechanism for wrong password)
# password = get_password()

# # Retry loop in case the user provided an incorrect password
# max_attempts = 3  # Number of attempts before default password is used
# attempts = 0
# df_identity_to_name = None
# df_name_to_email = None

# # Start retry loop for reading Excel file
# while attempts < max_attempts:
#     try:
#         if is_password_protected(file_path):
#             df_identity_to_name = read_excel(file_path, password)
#             df_name_to_email = read_excel(file_path, password)
#         else:
#             df_identity_to_name = pd.read_excel(file_path, sheet_name="Lookup-Sheet", usecols=[3, 4, 5, 6], engine="openpyxl")
#             df_name_to_email = pd.read_excel(file_path, sheet_name="Mail-IDs", usecols=[0, 1], engine="openpyxl")
#         break
#     except Exception as e:
#         print(f"Error reading Excel file with provided password: {e}")
#         if attempts < max_attempts - 1:
#             print("Please try again.")
#             password = get_password()  # Ask for the password again
#         attempts += 1

# # If all attempts failed, use the default password
# if df_identity_to_name is None and df_name_to_email is None and attempts >= max_attempts:
#     print(f"All {max_attempts} attempts failed. Falling back to default password.")
#     df_identity_to_name = read_excel(file_path, DEFAULT_PASSWORD)
#     df_name_to_email = read_excel(file_path, DEFAULT_PASSWORD)

# # Create Outlook application instance
# try:
#     outlook = win32com.client.Dispatch("Outlook.Application")
#     namespace = outlook.GetNamespace("MAPI")
#     folder = namespace.Folders["support.ahmedabad@gtpl.net"].Folders["Inbox"].Folders["Changemgmt_Airtel"]
# except Exception as e:
#     print(f"Error initializing Outlook: {e}")
#     print(traceback.format_exc())  # Print the traceback
#     exit()

# # Email Signature
# # signature_folder = os.path.expanduser(r'~\\AppData\\Roaming\\Microsoft\\Signatures')
# pwd = os.getcwd()
# signature_file = "Sign.htm"
# signature_path = os.path.join(pwd, "Sign", signature_file)

# try:
#     with open(signature_path, 'r', encoding='utf-8') as file:
#         html_signature = file.read()
# except Exception as e:
#     print(f"Error reading signature file: {e}")
#     print(traceback.format_exc())  # Print the traceback
#     html_signature = "Thanks & Regards\n\nGTPL Broadband Pvt Ltd.\nEnterprise-NOC\nTelephone:-+91-7966779500\nEmail: support.ahmedabad@gtpl.net"  # Fallback to an empty signature

# # Proceed with processing emails if the Excel files are read successfully
# if df_identity_to_name is not None and df_name_to_email is not None:
#     processed_emails = 0
#     sent_emails = 0

#     # Process each email in the folder
#     for item in folder.Items:
#         try:
#             if item.UnRead:  # Process only unread emails
#                 processed_emails += 1
#                 mail_body = item.Body
#                 start_greet = mail_body.index("With respect")
#                 end_greet = mail_body.index("Details of Maintenance")
#                 greet = mail_body[start_greet:end_greet-1]
#                 start_sub = item.Subject.index("CRQ")
#                 end_sub = item.Subject.index('Planned Maintenance activity')
#                 crq = item.Subject[start_sub:end_sub-1]

#                 if "Postponed" in greet or "Completed successfully" in greet:
#                     print("\nSkipped --> " + crq + "\n")
#                     print("======================================\n")
#                     item.UnRead = True
#                     item.Save()
#                     continue
                
#                 # Extract the identity table
#                 identity_table = extract_section(mail_body, "Impacted LSI : Party Name", "We request you")
#                 email_body_table = item.HTMLBody
#                 soup = BeautifulSoup(email_body_table, 'html.parser')
#                 maintenance_table = soup.find('table')

#                 # Checking for Maintenance Table and Identity Table
#                 if not maintenance_table or not identity_table:
#                     item.UnRead = True
#                     item.Save()
#                     continue

#                 # Parse the identity table
#                 identity_lines = identity_table.split('\n')
#                 for line in identity_lines:
#                     parts = line.split(":")
#                     if len(parts) != 2:
#                         continue
#                     identity_number = parts[0].strip()

#                     # Lookup company name in Excel sheet
#                     matched_row = df_identity_to_name[df_identity_to_name['AIRTEL Circuit ID'].astype(str) == identity_number]
#                     if matched_row.empty:
#                         continue

#                     matched_ckt = matched_row['AIRTEL Circuit ID'].values[0]
#                     recipient_name = matched_row['Customer Name'].astype(str).values[0].strip()
#                     # If recipient name not found
#                     if pd.isna(recipient_name) or recipient_name == '':
#                         print(f"\nCKT ID: {matched_ckt} has no Customer Name in 'Lookup-Sheet'.\n")
#                         continue
#                     df_name_to_email['Customer Name'] = df_name_to_email['Customer Name'].str.strip()
#                     matched_emails = df_name_to_email[df_name_to_email['Customer Name'] == recipient_name]
#                     # If recipient Email-ID not found
#                     if matched_emails['Email-ID'].isna().all():
#                         print(f"\nCKT ID: {matched_ckt} with Customer Name: {recipient_name} has no Email-ID in 'Mail-IDs' Sheet.\n")
#                         continue
#                     email_list = matched_emails['Email-ID'].tolist()
#                     a_end = matched_row["Location From"].values[0]
#                     b_end = matched_row["Location To"].values[0]
#                     if pd.isna(a_end) or a_end == '' or a_end == 'NA':
#                         a_end = None
#                     location = f'{b_end} ILL' if a_end is None else f'{a_end} to {b_end}'

#                     if email_list:
#                         # Construct email body
#                         email_body = f"""
#                         <p style="color: #2F5496;"><b>Dear Team,</b></p>
#                         <p style="color: #2F5496;"><b>Below link(s) shall go down due to planned maintenance activity.</b></p>
#                         <p style="color: #2F5496;"><b>Please find below activity details and consider the link down time as mentioned.</b></p>
#                         <p style="color: #2F5496;"><b>CKT ID: {matched_ckt}</b></p>
#                         <p style="color: #2F5496;"><b>Location: {location}</b></p>
#                         <br>
#                         {maintenance_table}
#                         <br><br>
#                         {html_signature}"""

#                         # Send Email with exception handling for each matched_ckt
#                         try:
#                             send_email(f"GTPL || CKT ID: {matched_ckt} || Location: {location} || PLANNED ACTIVITY", email_body, email_list)
#                             sent_emails += 1
#                             print(f"\nCKT ID: {matched_ckt} --> {crq}")
#                             print(F"Customer Name: {recipient_name}")
#                             print(f"Location: {location}")
#                             print(f"Mail sent to {recipient_name} at {email_list}")
#                         except Exception as e:
#                             print(f"Failed to send email for CKT ID {matched_ckt}: {e}")
#                             print(traceback.format_exc())  # Print the traceback

#                         print("-------------------------------------------------------")
#                         print("-------------------------------------------------------")

#                 item.UnRead = False
#                 item.Save()
#         except Exception as e:
#             print(f"Error processing email: {e}")
#             print(traceback.format_exc())  # Print the traceback

#     # End timing
#     end_time = time.time()
#     execution_time = end_time - start_time

#     print(f"\n{processed_emails} Emails processed successfully.\n")
#     print(f"\n{sent_emails} Emails sent successfully.\n")
#     print(f"\nTotal execution time: {execution_time:.2f} seconds\n")
# else:
#     print("Failed to read the Excel file after multiple attempts.")


import win32com.client
import pandas as pd
from bs4 import BeautifulSoup
import os
import time
import traceback
from pathlib import Path
import msoffcrypto
import io
import getpass

# Default password if user input is empty or incorrect
DEFAULT_PASSWORD = "Noc@123"

# Start timing
start_time = time.time()

# Configuration
EXCEL_FILE_PATH = Path.cwd() / "Excel_File" / "Latest_B2B.xlsx"
CC_RECIPIENTS = "'Tilok Choudhury' <tilok.choudhury@gtpl.net>; 'Amit Patel' <amit.patel@gtpl.net>"

# Create Outlook application instance
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    folder = namespace.Folders["support.ahmedabad@gtpl.net"].Folders["Inbox"].Folders["Changemgmt_Airtel"]
except Exception as e:
    print(f"Error initializing Outlook: {e}")
    print(traceback.format_exc())  # Print the traceback
    exit()

# Function to check if the Excel file is password protected
def is_password_protected(file_path):
    """Check if the Excel file is password protected by attempting to load it."""
    try:
        pd.read_excel(file_path, sheet_name="Lookup-Sheet", usecols=[3, 4, 5, 6], engine="openpyxl")
        return False  # No exception means the file is not password protected
    except ValueError as e:
        if "is not a zip file" in str(e):
            return True  # Likely password protected
        else:
            raise e

# Function to decrypt the Excel file if it is password protected
def decrypt_excel(file_path, password):
    """Decrypt the password-protected Excel file."""
    try:
        with open(file_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=password)  # Provide the password
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)  # Decrypt the file into a BytesIO object
            decrypted.seek(0)  # Rewind the stream to the beginning
        return decrypted
    except msoffcrypto.exceptions.InvalidKeyError:
        return None  # Invalid password
    except Exception as e:
        raise e

# Function to read the Excel file (with or without password protection)
def read_excel(file_path, password=None):
    """Read the Excel file (decrypted or not) into a pandas DataFrame."""
    try:
        if password:
            decrypted_file = decrypt_excel(file_path, password)
            if decrypted_file is None:
                return None  # Password was incorrect
            df = pd.read_excel(decrypted_file, engine="openpyxl")  # Load into pandas DataFrame
        else:
            df = pd.read_excel(file_path, engine="openpyxl")  # Load directly if not password-protected
        return df
    except Exception as e:
        return None

# Load Excel files (will be handled after checking if the file is password protected)
df_identity_to_name = None
df_name_to_email = None

# Check if the Excel file is password-protected
if is_password_protected(EXCEL_FILE_PATH):
    print("The Excel file is password protected.")
    
    # Prompt user for password (with retry mechanism)
    max_attempts = 3
    attempts = 0
    while attempts < max_attempts:
        password = getpass.getpass("Please enter the password for the Excel file: ")
        df_identity_to_name = read_excel(EXCEL_FILE_PATH, password)
        df_name_to_email = read_excel(EXCEL_FILE_PATH, password)
        if df_identity_to_name is not None and df_name_to_email is not None:
            break
        else:
            print(f"Incorrect password. You have {max_attempts - attempts - 1} attempts left.")
            attempts += 1
    if df_identity_to_name is None or df_name_to_email is None:
        print("Failed to open the Excel file after multiple attempts. Using default password.")
        password = DEFAULT_PASSWORD
        df_identity_to_name = read_excel(EXCEL_FILE_PATH, password)
        df_name_to_email = read_excel(EXCEL_FILE_PATH, password)
else:
    print("The Excel file is not password protected.")
    df_identity_to_name = pd.read_excel(EXCEL_FILE_PATH, sheet_name="Lookup-Sheet", usecols=[3, 4, 5, 6])
    df_name_to_email = pd.read_excel(EXCEL_FILE_PATH, sheet_name="Mail-IDs", usecols=[0, 1])

# Email Signature
pwd = os.getcwd()
signature_file = "Sign.htm"
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
