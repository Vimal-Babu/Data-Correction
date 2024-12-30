import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Define file path
file_path = "DATA VW-KLM (1).xlsx"

# Load the Excel file
sheets = pd.read_excel(file_path, sheet_name=None)

# Function to validate phone numbers
def validate_phone(phone):
    if pd.isna(phone):
        return False
    return str(phone).isdigit() and len(str(phone)) == 10

# Function to validate email IDs
def validate_email(email):
    if pd.isna(email):
        return False
    email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(email_regex, str(email)))

# Iterate over each sheet
output_sheets = {}
for sheet_name, df in sheets.items():
    df.columns = df.columns.str.strip().str.upper() 
    # Validate phone numbers
    if 'CONTACT NO' not in df.columns:
        print(f"'CONTACT NO' not found in sheet '{sheet_name}'. Available columns: {df.columns.tolist()}")
        continue
    
    df['Phone_Valid'] = df['CONTACT NO'].apply(validate_phone)
    df['Phone2_Valid'] = df['CONTACT NO 2'].apply(validate_phone)

    # Validate email IDs
    df['Email_Valid'] = df['EMAIL ID'].apply(validate_email)

    # Check mandatory fields
    mandatory_columns = ['CUSTOMER NAME', 'CONTACT NO', 'EMAIL ID']
    df['Mandatory_Filled'] = df[mandatory_columns].notna().all(axis=1)

    # Highlight invalid fields
    for index, row in df.iterrows():
        if not row['Phone_Valid']:
            df.at[index, 'CONTACT NO'] = f"INVALID: {row['CONTACT NO']}"
        if not row['Email_Valid']:
            df.at[index, 'EMAIL ID'] = f"INVALID: {row['EMAIL ID']}"

    output_sheets[sheet_name] = df

# Save the cleaned data to a new Excel file with color coding
with pd.ExcelWriter("validated_output.xlsx", engine='openpyxl') as writer:
    for sheet_name, df in output_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
