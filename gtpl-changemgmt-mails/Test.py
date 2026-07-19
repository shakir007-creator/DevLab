# import os
# from openpyxl import Workbook

# # Specify the directory and file path
# dir_path = 'D:\\Shared\\Changemgmt\\Excel_File'
# file_name = 'Test.xlsx'
# file_path = os.path.join(dir_path, file_name)

# # Ensure the directory exists
# os.makedirs(dir_path, exist_ok=True)  # Creates the directory if it doesn't exist

# # Create the workbook and write to the file
# wb = Workbook()
# ws = wb.active
# data = ['Alice', 'Bob', 'Charlie']

# # Write data to column 'A' (starting from row 1)
# for idx, value in enumerate(data, start=1):
#     ws[f'A{idx}'] = value

# # Save the workbook at the specified path
# wb.save(file_path)


import win32com.client
import os
import pandas as pd

# Access Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
    
# Access the specific folder: support.ahmedabad@gtpl.net > Inbox > Changemgmt_Airtel
folder = namespace.Folders["support.ahmedabad@gtpl.net"].Folders["Inbox"].Folders["Changemgmt_Airtel"]

# Get all emails in the folder
messages = folder.Items

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

# Function to extract CKT IDs from the email body using the extract_section function
def extract_ckts_from_body(email_body):
    # Extract the section containing CKT IDs
    identity_table = extract_section(email_body, "Impacted LSI : Party Name", "We request you")
    if not identity_table:
        return []

    identity_lines = identity_table.split('\n')
    ckt_ids = []
    for line in identity_lines:
        parts = line.split(":")
        if len(parts) != 2:
            continue
        identity_number = parts[0].strip()
        ckt_ids.append(int(identity_number))
    return ckt_ids

# Path to the existing Excel file
file_path = r'D:\Shared\Changemgmt\Excel_File\Test.xlsx'

# Prepare a list to hold the CKT data
ckt_data = []

# Process each email
for message in messages:
    if message.Unread:  # Only process unread emails
        mail_body = message.Body
        start_greet = mail_body.index("With respect")
        end_greet = mail_body.index("Details of Maintenance")
        greet = mail_body[start_greet:end_greet-1]
        start_sub = message.Subject.index('CRQ: “') + len('CRQ: “')
        end_sub = message.Subject.index('” Planned Maintenance activity')
        crq = message.Subject[start_sub:end_sub]
        date1 = message.Subject.index("scheduled on ") + len("scheduled on ")
        date2 = message.Subject.index('-202')
        activity_date = message.Subject[date1:date2].strip()
        
        # Skip emails with "Postponed" or "Completed successfully" in the subject or body
        if "Postponed" in greet or "Completed successfully" in greet:
            print(f"\nSkipped --> {crq}\n")
            print("======================================\n")
            message.UnRead = True
            message.Save()
            continue

        # Extract CKT IDs from the email body
        ckt_ids = extract_ckts_from_body(mail_body)
        
        if ckt_ids:
            # Add the extracted CKT IDs to the list
            ckt_data.extend(ckt_ids)

        # Create a DataFrame to write to Excel
        df = pd.DataFrame(ckt_data, columns=["CKT_IDs"])  # Set column name as "CKT_IDs"

        # If the file exists, append data to the sheet; otherwise, create a new file
        if os.path.exists(file_path):
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                # Check if 'Sheet1' exists, if not, create it
                try:
                    writer.sheets['Sheet']
                except KeyError:
                    # If Sheet1 does not exist, create it
                    # df.to_excel(writer, sheet_name='Sheet1', index=False)
                    print("Sheet Doesn't Exists")
                else:
                    sheet = writer.sheets['Sheet']  # Access the existing sheet
                    max_row = sheet.max_row  # Get the max row to append the data
                    start_row = max_row + 3  # Start appending from the next row
                    sheet[f'C{start_row}'].value = f"CRQ: {crq} ({activity_date})"
                    for idx, value in enumerate(ckt_data, start=start_row):
                        sheet[f'C{idx+1}'].value = value  # Write data to column 'C' (CKT_IDs)
        else:
            # If the file does not exist, create it and add data to 'Sheet1'
            df.to_excel(file_path, index=False, sheet_name='Sheet')

        # Mark the email as read
        message.UnRead = True
        ckt_data.clear()

print(f"Data written and saved to the Excel file: {file_path}")
