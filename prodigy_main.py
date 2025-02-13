import os
import re
import pdfplumber
import pandas as pd
from openpyxl import Workbook
from datetime import datetime, timedelta



pdf_path = "get from streamlit"

folder_name = "tables"
with pdfplumber.open(pdf_path) as pdf:
    for page_num, page in enumerate(pdf.pages, start=1):
        tables = page.extract_tables()  # Extract all tables on the page

        if tables:
            for table_idx, table in enumerate(tables, start=1):
                df = pd.DataFrame(table)  # Convert table to DataFrame

                # Extract date only from the first table on page 1
                if page_num == 1 and table_idx == 1 and df.shape[1] > 1:
                    raw_date = df.iloc[0, 1]  # First row, second column

                    try:
                        # Convert to datetime object and format as "7-Oct-2024"
                        date_obj = datetime.strptime(raw_date, "%d %B %Y")
                        formatted_date = date_obj.strftime("%d-%b-%Y")
                    except ValueError:
                        formatted_date = "Unknown-Date"

                    print(f"Extracted Date: {formatted_date}")  # Debugging

                # Generate output filename
                output_filename = f"extracted_table_page_{page_num}_table_{table_idx}.xlsx"

                # Save each table without headers (all rows as data)
                output_filename = os.path.join(folder_name, output_filename)
                df.to_excel(output_filename, index=False, header=False)
                print(f"Table {table_idx} from page {page_num} saved as {output_filename}")



# Compute adjusted dates for column names
date_minus_1_month = (datetime.strptime(formatted_date, "%d-%b-%Y") - timedelta(days=29)).strftime("%d-%b-%Y")
date_minus_2_months = (datetime.strptime(formatted_date, "%d-%b-%Y") - timedelta(days=91)).strftime("%d-%b-%Y")
# date_minus_1_month

table_column_mappings = {
    "extracted_table_page_1_table_2.xlsx": [
        "Bond", "ISIN", "Currency", "Notes Held", "Clean Price",
        "Clean Price + Interest", "Remaining Principal", "Accrued Interest",
        f"Balance {formatted_date}", "Pool factor"
    ],
    "extracted_table_page_1_table_3.xlsx": [
        "Bond", "Currency", "Margin Above Base", f"Base Rate {date_minus_2_months}",
        "Target Interest Rate", f"Balance {date_minus_1_month}",
        "New Investment or Sale", "Interest Earned", "Interest Payment",
        "Principal Payment", f"Balance {formatted_date}"
    ],
    "extracted_table_page_2_table_2.xlsx": [
        "Bond", "ISIN", "Currency", "Notes Held", "Clean Price",
        "Clean Price + Interest", "Remaining Principal", "Accrued Interest",
        f"Balance {formatted_date}", "Pool factor"
    ],
    "extracted_table_page_2_table_3.xlsx": [
        "Bond", "Currency", "Margin Above Base", f"Base Rate {date_minus_2_months}",
        "Target Interest Rate", f"Balance {date_minus_1_month}",
        "New Investment or Sale", "Interest Earned", "Interest Payment",
        "Principal Payment", f"Balance {formatted_date}"
    ],
    "extracted_table_page_3_table_2.xlsx": [
        "Bond", "ISIN", "Currency", "Notes Held", "Clean Price",
        "Clean Price + Interest", "Remaining Principal", "Accrued Interest",
        f"Balance {formatted_date}", "Pool factor"
    ],
    "extracted_table_page_4_table_2.xlsx": [
        "Bond", "Currency", "Margin Above Base", f"Base Rate {date_minus_2_months}",
        "Target Interest Rate", f"Balance {date_minus_1_month}",
        "New Investment or Sale", "Interest Earned", "Interest Payment",
        "Principal Payment", f"Balance {formatted_date}"
    ]
}


for file_name, columns in table_column_mappings.items():
    try:
        file_path = os.path.join(folder_name, file_name)
        df = pd.read_excel(file_path, header=None)

        df.columns = columns

        df.to_excel(file_path, index=False)

        print(f"Updated column names for {file_name}")
    except Exception as e:
        print(f"Error updating {file_name}: {e}")

tables_folder = os.path.join(os.getcwd(), "tables")
excel_files = [f for f in os.listdir(tables_folder) if f.endswith('.xlsx') and 'table_1' not in f]

# Function to extract page and table numbers from the filename
def extract_page_table_numbers(file_name):
    # Use regular expressions to extract the page and table numbers based on your filename format
    page_match = re.search(r'page_(\d+)', file_name)  # Match 'page_X'
    table_match = re.search(r'table_(\d+)', file_name)  # Match 'table_Y'

    if page_match and table_match:
        page_num = int(page_match.group(1))  # Extract the page number
        table_num = int(table_match.group(1))  # Extract the table number
        return (page_num, table_num)
    else:
        return (float('inf'), float('inf'))  # In case something goes wrong, return large numbers

# Sort the files based on extracted page and table numbers
sorted_files = sorted(excel_files, key=extract_page_table_numbers)
sorted_files

output_file = 'prodigy_pdf_output.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    start_row = 0  # Track the starting row for each dataset
    
    for file in sorted_files:
        # Read the data from the current file
        file_path = os.path.join(folder_name, file)
        data = pd.read_excel(file_path)
        
        # Write data to the final Excel file at the correct row position
        data.to_excel(writer, sheet_name='Sheet1', startrow=start_row, index=False)

        # Update start_row to place the next dataset after 2 empty rows
        start_row += len(data) + 2  