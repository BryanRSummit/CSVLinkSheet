from sheets_login import sheets_login
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from gspread.exceptions import APIError
import time
import re
import csv

def first_empty_row(sheet):
    str_list = list(filter(None, sheet.col_values(1)))  # Get all values in column A
    return len(str_list) + 1

def select_csv_file():
    # Hide the root window
    Tk().withdraw()
    # Open file dialog
    file_path = askopenfilename(
        filetypes=[("CSV files", "*.csv")], 
        title="Select a CSV file"
    )
    return file_path


def read_csv_file(file_path):
    if file_path:
        data = pd.read_csv(file_path)
        return data
    else:
        print("No file selected")

def exponential_backoff(attempt):
    if attempt == 0:
        return 1
    elif attempt == 1:
        return 2
    else:
        return (attempt ** 2)

def update_sheet_with_retry(sheet, values, start_row, end_row, max_attempts=5):
    attempt = 0
    while attempt < max_attempts:
        try:
            sheet.batch_update([{
                'range': f'A{start_row}:b{end_row}',
                'values': [[cell[2] for cell in values[i:i+2]] for i in range(0, len(values), 2)]
            }])
            return
        except APIError as e:
            if e.response.status_code == 429:  # Too Many Requests
                attempt += 1
                if attempt == max_attempts:
                    raise
                sleep_time = exponential_backoff(attempt)
                time.sleep(sleep_time)
            else:
                raise


def filter_salesforce_ids_dict(account_dict):
    # This pattern matches 15 or 18 character Salesforce IDs
    pattern = re.compile(r'^[a-zA-Z0-9]{15}(?:[a-zA-Z0-9]{3})?$')
    
    # Create a new dictionary with only the entries that have valid IDs
    #filtered_dict = {name: (id, size) for name, id, size in account_dict.items() if pattern.match(id)}
    filtered_dict = {name: (id, size) for name, (id, size) in account_dict.items() if pattern.match(id)}
    
    return filtered_dict



if __name__ == "__main__":
    sheet = sheets_login()

    #open CSV from somewhere
    csv_file_path = select_csv_file()
    data = read_csv_file(csv_file_path)
    #account_dict = dict(zip(data['Account Name'], data['Account ID'], data["Average Policy Size"]))
    account_dict = {name: (id, size) for name, id, size in zip(data['Account Name'], data['Account ID'], data["Average Policy Size"])}

    filtered_account_dict = filter_salesforce_ids_dict(account_dict)


    row = first_empty_row(sheet)
    start_row = row
    headers = sheet.row_values(1)
    # Prepare batch update
    batch_update = []

    for name, info in filtered_account_dict.items():
        row_data = [
            (row, headers.index("Name") + 1, name),
            (row, headers.index("AVG Policy Size") + 1, info[1]),
            (row, headers.index("Link") + 1, f"https://reddsummit.lightning.force.com/lightning/r/Account/{info[0]}/view")
        ]
        batch_update.extend(row_data)
        row += 1
    # Perform batch update with retry
    update_sheet_with_retry(sheet, batch_update, start_row, row)

    print("Sheet has been filled with links!")
