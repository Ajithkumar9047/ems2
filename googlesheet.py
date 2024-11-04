import gspread
from google.auth import exceptions
from google.oauth2 import service_account
import pandas as pd
from datetime import datetime

current_date = datetime.now().strftime('%d/%m/%Y')
current_time = datetime.now().strftime('%I:%M %p')


# Set up Google Sheets API credentials
credentials = service_account.Credentials.from_service_account_file(
    'local-dialect-398506-9f8890b4a8e1.json',
    scopes=['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
)

# Authenticate with Google Sheets
gc = gspread.authorize(credentials)

# Specify the Google Sheets document by its sheet ID
sheet_id = "1xuhOqumuvpjmuwKe40t09Fwj8gbxbJ2-vPwQGa2e4Ls"

# Open the Google Sheets document by its sheet ID
spreadsheet = gc.open_by_key(sheet_id)

# Select a specific worksheet
worksheet = spreadsheet.worksheet("Master_Count _Table")

# Update a specific cell in the last row of the "ColumnB" column
column_name = "ColumnB"
new_value = "Updated Value"

# Find the last row in the worksheet
existing_data = worksheet.get_all_values()
df = pd.DataFrame(existing_data)
last_row = len(existing_data) + 1

# Write the data to all columns in the last row
data_to_write = [current_date, current_time, "unavailable"]
worksheet.insert_row(data_to_write, last_row)

# Update a specific cell in the last row of "ColumnB"
worksheet.update_cell(last_row, df.columns.get_loc(column_name) + 1, new_value)

# Print the updated data
updated_data = worksheet.get_all_values()
print(f"Updated data: {updated_data}")
