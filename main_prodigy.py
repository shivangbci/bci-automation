import os
import re
import io
import pdfplumber
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

# Streamlit App Title
st.title("📄 Prodigy Bond Warehouse PDF Extractor")

# File Uploader (PDF)
uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

if uploaded_file:
    st.success("✅ PDF uploaded successfully!")

    # Create 'tables' folder if it doesn't exist
    folder_name = "tables"
    os.makedirs(folder_name, exist_ok=True)

    # Process the PDF file
    with pdfplumber.open(uploaded_file) as pdf:
        formatted_date = "Unknown-Date"  # Default value
        extracted_tables = []
        
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            
            if tables:
                for table_idx, table in enumerate(tables, start=1):
                    df = pd.DataFrame(table)

                    # Extract the date from the first table on page 1
                    if page_num == 1 and table_idx == 1 and df.shape[1] > 1:
                        raw_date = df.iloc[0, 1]
                        try:
                            date_obj = datetime.strptime(raw_date, "%d %B %Y")
                            formatted_date = date_obj.strftime("%d-%b-%Y")
                        except ValueError:
                            formatted_date = "Unknown-Date"

                    # Save each table as an Excel file
                    output_filename = f"extracted_table_page_{page_num}_table_{table_idx}.xlsx"
                    file_path = os.path.join(folder_name, output_filename)
                    df.to_excel(file_path, index=False, header=False)
                    extracted_tables.append((file_path, df))

    st.write(f"**Extracted Date:** {formatted_date}")

    # Compute adjusted dates for dynamic column names
    date_minus_1_month = (datetime.strptime(formatted_date, "%d-%b-%Y") - timedelta(days=29)).strftime("%d-%b-%Y")
    date_minus_2_months = (datetime.strptime(formatted_date, "%d-%b-%Y") - timedelta(days=91)).strftime("%d-%b-%Y")

    # Define column mappings for specific tables
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

    # Update column names based on mapping
    for file_name, columns in table_column_mappings.items():
        try:
            file_path = os.path.join(folder_name, file_name)
            df = pd.read_excel(file_path, header=None)
            df.columns = columns
            df.to_excel(file_path, index=False)
        except Exception as e:
            st.error(f"Error updating {file_name}: {e}")

    # Merge all extracted tables into a single Excel file
    output_file = "prodigy_pdf_output.xlsx"
    tables_folder = os.path.join(os.getcwd(), "tables")
    excel_files = sorted(
        [f for f in os.listdir(tables_folder) if f.endswith('.xlsx') and 'table_1' not in f],
        key=lambda x: (int(re.search(r'page_(\d+)', x).group(1)), int(re.search(r'table_(\d+)', x).group(1)))
    )

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        start_row = 0
        for file in excel_files:
            file_path = os.path.join(folder_name, file)
            data = pd.read_excel(file_path)
            data.to_excel(writer, sheet_name='Sheet1', startrow=start_row, index=False)
            start_row += len(data) + 2  

    # Provide download button for final Excel file
    with open(output_file, "rb") as f:
        st.download_button("📥 Download Processed Excel", f, file_name="extracted_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

