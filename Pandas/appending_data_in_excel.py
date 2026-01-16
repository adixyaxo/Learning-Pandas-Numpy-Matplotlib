# In this programme we will look at how to append data in an excel sheet
import pandas as pd
import openpyxl
import os

file_path = "data/file.xlsx"

    
# New data to add
new_data = {
    'User_ID': [106],
    'Name': ['Rahul'],
    'Car_Number': ['DL-12-AB-1234'],
    'Entry_Time': ['11:00 AM'],
    'Status': ['Parked']
}

df = pd.DataFrame(new_data)

if not os.path.isfile(file_path):
    # If file doesn't exist, create it with headers
    df.to_excel(file_path, index=False)

# 1. Load the existing workbook to find the last row
book = openpyxl.load_workbook(file_path)
sheet = book.active
start_row = sheet.max_row

# 3. Append data using 'overlay' mode
# 'if_sheet_exists="overlay"' prevents creating a new sheet 'Sheet11'
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    df.to_excel(
        writer, #
        index=False, 
        header=False,      # Don't repeat the header!
        startrow=start_row # Start writing after the last filled row
    )