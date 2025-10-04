import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
import calendar

# Define the required columns to extract and output
required_columns = [
    'Date', 'Time', 'Debtor', 'Account No.', 'Card No.', 'Service No.', 
    'Call Status', 'Status', 'Remark', 'Remark By', 'Remark Type', 'Collector', 
    'Client', 'Product Description', 'PTP Amount', 'Next Call', 'PTP Date', 
    'Claim Paid Amount', 'Claim Paid Date', 'Dialed Number', 'Balance', 
    'Cycle', 'Old IC', 'Debtor ID'
]

# Define the status values to filter on (substring match)
status_filter = ['BP', 'PAYMENT', 'PTP', 'RPC', 'BANK ESCALATION', 'TPC', 'DROPPED', 'NEGATIVE', 'NEGATIVE CALLOUTS', 'VM']

# Streamlit app title
st.title("XLSX File Processor for Campaigns")

# Sidebar file uploader (allow multiple files)
uploaded_files = st.sidebar.file_uploader("Upload your XLSX files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    with st.spinner("Processing files..."):
        try:
            # Initialize an empty list to store filtered DataFrames from all files
            all_filtered_dfs = []
            
            # Process each uploaded file
            for uploaded_file in uploaded_files:
                # Read the XLSX file, loading only required columns
                df = pd.read_excel(uploaded_file, usecols=required_columns, engine='openpyxl')
                
                # Filter rows where 'Status' contains any of the status_filter values (case-insensitive)
                pattern = '|'.join([re.escape(status) for status in status_filter])
                filtered_df = df[df['Status'].str.contains(pattern, case=False, na=False, regex=True)]
                
                # Append filtered DataFrame to the list
                all_filtered_dfs.append(filtered_df)
            
            # Concatenate all filtered DataFrames
            combined_df = pd.concat(all_filtered_dfs, ignore_index=True)
            
            # Get unique clients (campaigns) from combined data
            unique_clients = combined_df['Client'].unique()
            
            # Determine date range from the 'Date' column
            if 'Date' in combined_df.columns:
                combined_df['Date'] = pd.to_datetime(combined_df['Date'], errors='coerce')
                min_date = combined_df['Date'].min()
                max_date = combined_df['Date'].max()
                
                # Default to current month if dates are invalid
                if pd.isna(min_date) or pd.isna(max_date):
                    today = datetime.today()
                    month_name = today.strftime("%B")
                    last_day = calendar.monthrange(today.year, today.month)[1]
                    date_range_str = f"{month_name} 1-{last_day}"
                else:
                    # Check if dates are in the same month
                    if min_date.month == max_date.month and min_date.year == max_date.year:
                        month_name = min_date.strftime("%B")
                        date_range_str = f"{month_name} {min_date.day}-{max_date.day}"
                    else:
                        # If spanning multiple months, use full month for min_date
                        month_name = min_date.strftime("%B")
                        last_day = calendar.monthrange(min_date.year, min_date.month)[1]
                        date_range_str = f"{month_name} 1-{last_day}"
            else:
                # Fallback if no Date column
                today = datetime.today()
                month_name = today.strftime("%B")
                last_day = calendar.monthrange(today.year, today.month)[1]
                date_range_str = f"{month_name} 1-{last_day}"
            
            # Get today's date in MMDDYY format
            today = datetime.today()
            today_str = today.strftime("%m%d%y")
            
            # Construct file name
            file_name = f"{date_range_str} DRR Summary {today_str}.xlsx"
            
            # Create an in-memory buffer for the output Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for client in unique_clients:
                    # Get subset for this client
                    client_df = combined_df[combined_df['Client'] == client]
                    # Write to sheet named after client (truncate to 31 chars for Excel limit)
                    sheet_name = str(client)[:31]
                    client_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Prepare the buffer for download
            output.seek(0)
            
            # Provide download button
            st.download_button(
                label="Download Filtered Excel",
                data=output,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Processing complete! Download the file above.")
        
        except ValueError as e:
            # If a ValueError occurs (e.g., column mismatch), read headers from the first file
            temp_df = pd.read_excel(uploaded_files[0], nrows=0, engine='openpyxl')  # Read only headers
            actual_headers = list(temp_df.columns)
            st.error(f"Error: {str(e)}\n\nActual headers in the first file: {actual_headers}")
            st.write("Please check if the required columns match the actual headers exactly (including spaces and capitalization).")
