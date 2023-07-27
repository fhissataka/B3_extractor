import os
import csv
import tempfile
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import webbrowser

# B3 web scrapping
# by Fernando Hissataka 2023  fhissataka@gmail.com
# input: an Excel file with a single column and no header, with dates to be researched
# output: CSV files in your Windows TEMP directory, under a B3 directory

def fetch_data_from_url(curve, dates_column):
    for date in dates_column:
        # Convert the date to a Python datetime object
        date_obj = pd.to_datetime(date, errors='coerce')

        # A list to store the data for saving to CSV file
        data_to_save = []  

        if not pd.isnull(date_obj):
            formatted_date = date_obj.strftime("%d/%m/%Y")
            print("Processing the curve "  + curve + " for day " + formatted_date)

            # Get the URL with properly replaced day, month, and year
            url = f"https://www2.bmf.com.br/pages/portal/bmfbovespa/boletim1/TxRef1.asp?Data={formatted_date}&Data1={date_obj.strftime('%Y%m%d')}&slcTaxa={curve}"

            # Open the URL in the default web browser
            # webbrowser.open_new_tab(url)

            # Fetch the content using requests
            response = requests.get(url)
            if response.status_code == 200:
                # Parse the content with Beautiful Soup
                soup = BeautifulSoup(response.content, 'html.parser')
                tabela = soup.find_all("table")[1].getText();
                
                # Remove \r characters from tabela
                tabela = tabela.replace('\r', '')
                
                # Combine every two consecutive lines with a semicolon (;)
                data_lines = [line.strip() for line in tabela.split('\n') if line.strip()]

                # Adding the first line, which is a header of the date 
                data_to_save.append([formatted_date, " "])

                for i in range(4, len(data_lines), 2):
                    if i + 1 < len(data_lines):
                        data_to_save.append([data_lines[i], data_lines[i + 1]])
                        # print(f"{data_lines[i]}; {data_lines[i + 1]}")

                # Save the data to a CSV file
                file_name = f"{date_obj.strftime('%Y%m%d')}_{curve}.csv"

                temp_directory = tempfile.gettempdir()

                # Create the directory path for B3\LIB under the temp directory
                b3_tmp_directory = os.path.join(temp_directory, "B3", curve)

                # Create the directory if it does not exist
                os.makedirs(b3_tmp_directory, exist_ok=True)

                file_path = os.path.join(b3_tmp_directory, file_name)

                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    csvwriter = csv.writer(csvfile, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                    # csvwriter.writerow([formatted_date, None])  # Write header row     
                    for row in data_to_save:
                        # Replace decimal indicator "." with ","
                        row = [item.replace('.', ',') for item in row]
                        csvwriter.writerow(row) 

            else:
                content = f"Failed to fetch the page. Status code: {response.status_code}"


if __name__ == "__main__":
    # Create a Tkinter root window (main window)
    root = tk.Tk()
    root.withdraw()

    # Ask user to select an Excel file
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    # Check if the user selected a file
    if file_path:
        # Read the Excel file using pandas
        df = pd.read_excel(file_path, header=None)

        # Extract the first column (index 0) from the DataFrame
        dates_column = df.iloc[0:, 0]

        list_of_curves = [ "DOC","DOL", "DIC", "PTX", "SLP", "LIB"]
        # list_of_curves = [ "DOC","DOL" ]
        for curve in list_of_curves: 
            fetch_data_from_url(curve, dates_column)
