import gspread
from google.oauth2.service_account import Credentials

# Define the scope
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Path to the downloaded JSON key file
creds_path = "C:\\Users\\phams\\Downloads\\mythic-evening-425602-k5-bcaf0ca4be0f.json"

# Create a Credentials object from the JSON key file
creds = Credentials.from_service_account_file(creds_path, scopes=scope)

# Authorize the client
client = gspread.authorize(creds)

# Open the Google Sheet by title
spreadsheet = client.open("check_routes")

# Select the first sheet
sheet = spreadsheet.sheet1

# Write to the first sheet
sheet.update_cell(2,1,"Write successful")

# Get all values from the sheet
data = sheet.get_all_values()

# Print the data
for row in data:
    print(row)
