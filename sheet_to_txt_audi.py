from datetime import date
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Path to your service account key file
SERVICE_ACCOUNT_FILE = 'keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Google Sheets API credentials
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Google Sheets details
SAMPLE_SPREADSHEET_ID = '1zp_8pqbY_Lds4lcqEFdObeza8DA18GWB8OnetXyeMmQ'

# Construct file path dynamically
file_dir = os.path.join(os.getenv("USERPROFILE"), "Desktop", "Kamini")
file_path = os.path.join(file_dir, "audi_report.txt")
print(f"File path already exists: {file_path}")

# Ensure the directory exists
if not os.path.exists(file_dir):
    os.makedirs(file_dir)
    print(f"Directory created: {file_dir}")

# Step 1: Write the file header
try:
    with open(file_path, 'w') as fh:
        fh.write("=================\n")
        fh.write("Audi Leads \n")
        fh.write("=================\n\n\n")
    print("File header written successfully.")
except Exception as e:
    print(f"Error writing file header: {e}")
    exit()

# Step 2: Fetch data from Google Sheets
try:
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='ADW!A1:Q10000').execute()
    values = result.get('values', [])
    print("Google Sheets data fetched successfully.")
except Exception as e:
    print(f"Error fetching data from Google Sheets: {e}")
    exit()

# Step 3: Process rows and write leads to the file
leadnum = 1
leads_found = False

try:
    for row in values:
        if len(row) > 15:  # Ensure the row has enough columns
            today_date = str(date.today())

            # Check for today's date and delivery status
            if row[1] == today_date and row[15] == "DELIVERED":
                leads_found = True
                with open(file_path, 'a') as fh:
                    # Add "(DIGITAL)" and the date for the leads being processed
                    fh.write(f"*LEAD NO {leadnum}, (DIGITAL) {today_date}*\n")
                    fh.write("--------------------------\n")
                    leadnum += 1

                l1 = [
                    "SR.", "DATE", "Customer Name", "Mobile No", "Location", "Model",
                    "Test Drive", "Callback", "Present Car", "Finance", "Occupation",
                    "Budget", "Remark"
                ]

                # Write row details
                for i, cell in enumerate(row):
                    if i < len(l1):
                        cell_output = str(cell)
                        with open(file_path, 'a') as fh:
                            if i in {2, 3, 4, 6, 9, 10, 11}:
                                fh.write(f"{l1[i]}: {cell_output}\n")
                            elif i in {5, 7, 8, 12}:
                                fh.write(f"*{l1[i]}: {cell_output}*\n")

                with open(file_path, 'a') as fh:
                    fh.write("\n\n")

    # Final log
    if leads_found:
        print(f"Leads populated successfully in: {file_path}")
    else:
        print("No leads found to populate in the file.")
except Exception as e:
    print(f"Error processing and writing leads: {e}")