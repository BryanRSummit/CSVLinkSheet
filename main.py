from sheets_login import sheets_login
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import requests
import json


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

if __name__ == "__main__":
    sheet = sheets_login()

    #open CSV from somewhere
    csv_file_path = select_csv_file()
    data = read_csv_file(csv_file_path)
