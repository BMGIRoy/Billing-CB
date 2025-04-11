import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import calendar
import io

def process_excel_data(uploaded_file):
    """
    Process the uploaded Excel file to extract hierarchical data and contract information
    """
    # Initialize month_patterns here to avoid "possibly unbound" error
    month_patterns = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
    # Read Excel file
    xls = pd.ExcelFile(uploaded_file)
    
    # Get all sheet names
    available_sheets = xls.sheet_names
    
    # Check if required sheets exist (case insensitive)
    sheet_map = {}
    required_sheets = ['Contracts', 'Consultant Billing']
    
    for sheet in available_sheets:
        for required in required_sheets:
            if required.lower() == sheet.lower() or required.lower() in sheet.lower():
                sheet_map[required] = sheet
    
    # Verify all required sheets are found
    missing_sheets = [sheet for sheet in required_sheets if sheet not in sheet_map]
    if missing_sheets:
        available_sheets_str = ", ".join(available_sheets)
        raise ValueError(f"Required sheet(s) {missing_sheets} not found. Available sheets are: {available_sheets_str}")
    
    # Process Contracts sheet
    try:
        # Read with explicit column names as Excel might have unnamed columns
        contracts_data = pd.read_excel(xls, sheet_map['Contracts'])
        
        # Print the columns we found for debugging
        print(f"Detected Contracts sheet columns: {contracts_data.columns.tolist()}")
        
        # Convert unnamed columns to string format for easier handling
        contracts_data.columns = [str(col) for col in contracts_data.columns]
        
        # Map the actual column names to the expected ones
        contract_column_mapping = {
            'Client': 'Client Name',
            'Work': 'Type of Work',
            'PO No.': 'PO No',
            'BH': 'Business Head',
            'Total Value (F+V)': 'Total PO Value',
            'Fixed Balance': 'PO Balance',
            # Add mappings for potentially unnamed columns
            'Unnamed: 0': 'ID',
            'Unnamed: 1': 'Client Name',
            'Unnamed: 2': 'Type of Work',
            'Unnamed: 3': 'PO No',
            'Unnamed: 4': 'Business Head',
            'Unnamed: 5': 'Total PO Value',
            'Unnamed: 6': 'PO Balance'
        }
        
        # Rename columns if they exist in the Contracts sheet
        for actual_col, expected_col in contract_column_mapping.items():
            if actual_col in contracts_data.columns:
                contracts_data.rename(columns={actual_col: expected_col}, inplace=True)
                
        # Convert numeric columns to appropriate types
        numeric_columns = ['Total PO Value', 'PO Balance']
        for col in numeric_columns:
            if col in contracts_data.columns:
                contracts_data[col] = pd.to_numeric(contracts_data[col], errors='coerce').fillna(0)
    except Exception as e:
        # If there's an error processing the Contracts sheet, create an empty DataFrame
        print(f"Error processing Contracts sheet: {str(e)}")
        contracts_data = pd.DataFrame(columns=['Client Name', 'Type of Work', 'PO No', 'Business Head', 'Total PO Value', 'PO Balance'])
    
    # Process Consultant Billing sheet (which may contain pivot data)
    try:
        # First, try to detect the structure of the pivot table
        # For pivot tables, the first row often contains month labels and the second row contains column labels (T Amt, N Amt)
        pivot_preview = pd.read_excel(xls, sheet_map['Consultant Billing'], nrows=5)
        
        # Initialize month_patterns here to avoid "possibly unbound" error
        month_patterns = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
        
        # Check if it looks like a pivot table by examining the first few rows for month patterns
        has_month_headers = any(
            any(month in str(col).lower() for month in month_patterns) 
            for col in pivot_preview.columns
        )
        
        if has_month_headers:
            # This looks like a pivot table with month headers
            # Try reading with a multi-index header (usually 2 rows for month and T Amt/N Amt)
            billing_data_raw = pd.read_excel(xls, sheet_map['Consultant Billing'], header=[0, 1])
            
            # Process the pivot data into a structured format
            billing_data = process_pivot_table(billing_data_raw)
        else:
            # Try standard reading if not a pivot
            billing_data = pd.read_excel(xls, sheet_map['Consultant Billing'])
            
            # If successful but no expected columns, try with different header structure
            if not any(col.lower() in ['t amt', 'n amt', 'date'] for col in billing_data.columns):
                raise ValueError("Expected columns not found with standard header")
                
    except Exception as e:
        # Try different approaches if the above methods fail
        try:
            # Read without any headers and try to infer structure
            billing_data_raw = pd.read_excel(xls, sheet_map['Consultant Billing'], header=None)
            
            # Look for month patterns in the first few rows to determine if it's a pivot
            first_rows = billing_data_raw.iloc[:5].astype(str)
            
            month_found = False
            header_row = 0  # Default header row if we can't identify a specific one
            for i, row in first_rows.iterrows():
                if any(month in row.astype(str).str.lower().str.contains('|'.join(month_patterns)).any() 
                      for month in month_patterns):
                    month_found = True
                    header_row = i
                    break
            
            if month_found:
                # Re-read with the identified header row
                billing_data_raw = pd.read_excel(xls, sheet_map['Consultant Billing'], header=[header_row, header_row+1])
                billing_data = process_pivot_table(billing_data_raw)
            else:
                # If we still can't find a month pattern, use basic processing
                billing_data = pd.read_excel(xls, sheet_map['Consultant Billing'])
        except Exception as e2:
            # Last resort: just read the data and set column names
            billing_data = pd.read_excel(xls, sheet_map['Consultant Billing'])
            # Set temporary column names if needed
            if billing_data.columns.dtype == 'int64':
                billing_data.columns = [f"Column_{i}" for i in range(len(billing_data.columns))]
    
    # Ensure required columns exist in billing_data
    required_columns = ['Business Head', 'Consultant', 'Client', 'Date', 'T Amt', 'N Amt']
    
    # Check if date column exists in various forms
    date_columns = [col for col in billing_data.columns if 'date' in col.lower()]
    if date_columns:
        # Rename the first date column to 'Date' if it exists but with different case/name
        billing_data.rename(columns={date_columns[0]: 'Date'}, inplace=True)
    
    # Check for T Amt and N Amt columns (variations in naming)
    t_amt_cols = [col for col in billing_data.columns if 't amt' in col.lower() or 'total amt' in col.lower()]
    n_amt_cols = [col for col in billing_data.columns if 'n amt' in col.lower() or 'net amt' in col.lower()]
    
    if t_amt_cols:
        billing_data.rename(columns={t_amt_cols[0]: 'T Amt'}, inplace=True)
    if n_amt_cols:
        billing_data.rename(columns={n_amt_cols[0]: 'N Amt'}, inplace=True)
    
    # Check for missing required columns and try to identify/create them
    business_head_col = None
    consultant_col = None
    client_col = None
    
    # Try to identify hierarchical columns
    for col in billing_data.columns:
        # Check the first few rows to determine column types
        sample = billing_data[col].head(10).astype(str).str.strip()
        
        # Look for Business Head (all caps might indicate headers)
        if any(s.isupper() and len(s) > 3 for s in sample) and not business_head_col:
            business_head_col = col
        
        # Check column names that might indicate these fields
        if any(term in col.lower() for term in ['business', 'head', 'bh']):
            business_head_col = col
        elif any(term in col.lower() for term in ['consultant', 'cons', 'resource']):
            consultant_col = col
        elif any(term in col.lower() for term in ['client', 'customer', 'account']):
            client_col = col
    
    # If we've identified replacements for the missing columns, rename them
    column_renames = {}
    if business_head_col and 'Business Head' not in billing_data.columns:
        column_renames[business_head_col] = 'Business Head'
    if consultant_col and 'Consultant' not in billing_data.columns:
        column_renames[consultant_col] = 'Consultant'
    if client_col and 'Client' not in billing_data.columns:
        column_renames[client_col] = 'Client'
    
    if column_renames:
        billing_data.rename(columns=column_renames, inplace=True)
    
    # Make sure required columns exist even if we couldn't find them
    for col in required_columns:
        if col not in billing_data.columns:
            if col == 'Date':
                billing_data[col] = pd.Timestamp.now()
            elif col in ['T Amt', 'N Amt']:
                billing_data[col] = 0
            else:
                billing_data[col] = 'Unknown'
    
    # Clean and prepare billing data
    billing_data = clean_billing_data(billing_data)
    
    # Get unique values for filters
    business_heads = sorted(billing_data['Business Head'].unique().tolist())
    consultants = sorted(billing_data['Consultant'].unique().tolist())
    clients = sorted(billing_data['Client'].unique().tolist())
    
    # Generate fiscal periods (e.g., FY 2022-23, FY 2023-24)
    fiscal_periods = get_fiscal_periods(billing_data['Date'])
    
    # Return the processed data
    return billing_data, contracts_data, business_heads, consultants, clients, fiscal_periods

def process_pivot_table(pivot_data):
    """
    Process a pivot table with hierarchical structure into a flat table with proper columns
    """
    # Print the columns we found in the pivot table for debugging
    print(f"Pivot table columns: {pivot_data.columns}")
    
    # Convert pivot_data columns to strings to handle unnamed columns
    if isinstance(pivot_data.columns, pd.MultiIndex):
        # For multi-index columns, create more descriptive strings
        new_cols = []
        for col in pivot_data.columns:
            if isinstance(col, tuple):
                new_cols.append(' '.join(str(x) for x in col if pd.notna(x)))
            else:
                new_cols.append(str(col))
        pivot_data.columns = new_cols
    else:
        pivot_data.columns = [str(col) for col in pivot_data.columns]
    
    # Create an empty DataFrame to store the processed data
    processed_data = pd.DataFrame(columns=['Business Head', 'Consultant', 'Client', 'Date', 'T Amt', 'N Amt'])
    
    # Get the multi-index columns (typically month and metric like T Amt, N Amt)
    date_columns = []
    t_amt_indices = []
    n_amt_indices = []
    
    # Extract month columns from multi-index
    for i, col in enumerate(pivot_data.columns):
        col_str = str(col)
        if 't amt' in col_str.lower() or 'total' in col_str.lower():
            date_str = col_str.split('T Amt')[0].strip() if 'T Amt' in col_str else col_str
            date_columns.append(date_str)
            t_amt_indices.append(i)
        elif 'n amt' in col_str.lower() or 'net' in col_str.lower():
            n_amt_indices.append(i)
            
    # If we found date columns with T Amt, process them
    if date_columns and t_amt_indices:
        # Initialize variables to track the hierarchy
        current_business_head = None
        current_consultant = None
        rows_to_add = []
        
        # Process each row in the pivot table
        for idx, row in pivot_data.iterrows():
            row_values = row.values
            first_cell = str(row_values[0]).strip() if pd.notna(row_values[0]) else ""
            
            # Check if this is a business head row (typically in ALL CAPS)
            if first_cell and first_cell.isupper() and len(first_cell) > 2:
                current_business_head = first_cell
                current_consultant = None
            # Check if this is a consultant row (typically first indented level)
            elif first_cell and not first_cell.isupper() and pd.notna(row_values[0]) and current_business_head is not None:
                current_consultant = first_cell
            # Check if this is a client row (typically second indented level)
            elif first_cell and current_consultant is not None:
                client = first_cell
                
                # Extract data for each date column
                for date_idx, date_str in enumerate(date_columns):
                    # Try to parse the date
                    try:
                        # Common date formats in Excel: 'Apr-22', 'April 2022', etc.
                        date_parts = date_str.split('-') if '-' in date_str else date_str.split(' ')
                        month_str = date_parts[0].strip()
                        year_str = date_parts[1].strip() if len(date_parts) > 1 else "2023"  # Default year if not specified
                        
                        # Convert month abbreviation to number
                        month_map = {
                            'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                            'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
                        }
                        month_num = None
                        for abbr, num in month_map.items():
                            if abbr in month_str.lower():
                                month_num = num
                                break
                        
                        if month_num is None:
                            continue
                        
                        # Format the year (handle '22' to '2022')
                        if len(year_str) == 2:
                            year = int("20" + year_str)
                        else:
                            year = int(year_str)
                        
                        # Create a datetime object for the first of the month
                        date = pd.Timestamp(year=year, month=month_num, day=1)
                        
                        # Get the T Amt and N Amt values
                        t_amt = row_values[t_amt_indices[date_idx]] if t_amt_indices[date_idx] < len(row_values) else 0
                        n_amt = row_values[n_amt_indices[date_idx]] if date_idx < len(n_amt_indices) and n_amt_indices[date_idx] < len(row_values) else 0
                        
                        # Add row to our results if we have T Amt or N Amt
                        if pd.notna(t_amt) or pd.notna(n_amt):
                            rows_to_add.append({
                                'Business Head': current_business_head,
                                'Consultant': current_consultant,
                                'Client': client,
                                'Date': date,
                                'T Amt': t_amt if pd.notna(t_amt) else 0,
                                'N Amt': n_amt if pd.notna(n_amt) else 0
                            })
                    except Exception as e:
                        # Skip this date if there's a parsing error
                        continue
        
        # Create the processed DataFrame
        if rows_to_add:
            processed_data = pd.DataFrame(rows_to_add)
        else:
            # If we couldn't extract structured data, return the original with flattened column names
            processed_data = pivot_data.copy()
            
            # Try to identify key columns
            for col in processed_data.columns:
                if 't amt' in col.lower() or 'total' in col.lower():
                    processed_data.rename(columns={col: 'T Amt'}, inplace=True)
                    break
            
            for col in processed_data.columns:
                if 'n amt' in col.lower() or 'net' in col.lower():
                    processed_data.rename(columns={col: 'N Amt'}, inplace=True)
                    break
            
            # Add missing columns if needed
            if 'Business Head' not in processed_data.columns:
                processed_data['Business Head'] = 'Unknown'
            if 'Consultant' not in processed_data.columns:
                processed_data['Consultant'] = 'Unknown'
            if 'Client' not in processed_data.columns:
                processed_data['Client'] = 'Unknown'
            if 'Date' not in processed_data.columns:
                processed_data['Date'] = pd.Timestamp.now()
    else:
        # If we couldn't find T Amt columns with dates, return the original data
        processed_data = pivot_data.copy()
        
        # Add necessary columns if they don't exist
        if 'Business Head' not in processed_data.columns:
            processed_data['Business Head'] = 'Unknown'
        if 'Consultant' not in processed_data.columns:
            processed_data['Consultant'] = 'Unknown'
        if 'Client' not in processed_data.columns:
            processed_data['Client'] = 'Unknown'
        if 'Date' not in processed_data.columns:
            processed_data['Date'] = pd.Timestamp.now()
        if 'T Amt' not in processed_data.columns:
            # Try to find a column that might contain T Amt
            t_amt_col = None
            for col in processed_data.columns:
                if 'total' in col.lower() or 'amount' in col.lower():
                    t_amt_col = col
                    break
            if t_amt_col:
                processed_data['T Amt'] = processed_data[t_amt_col]
            else:
                processed_data['T Amt'] = 0
        if 'N Amt' not in processed_data.columns:
            # Try to find a column that might contain N Amt
            n_amt_col = None
            for col in processed_data.columns:
                if 'net' in col.lower():
                    n_amt_col = col
                    break
            if n_amt_col:
                processed_data['N Amt'] = processed_data[n_amt_col]
            else:
                processed_data['N Amt'] = 0
    
    return processed_data

def clean_billing_data(df):
    """
    Clean and prepare the billing data for analysis
    """
    # Make a copy of the DataFrame to avoid modifying the original
    df = df.copy()
    
    # Print the columns we have for debugging
    print(f"Columns before cleaning: {df.columns.tolist()}")
    print(f"Sample data types: {df.dtypes}")
    
    # Handle potential column name issues
    for col in df.columns:
        # Convert any problematic column with spaces or special chars to strings
        if 'PO billing' in str(col) or 'Milestone' in str(col) or 'Planned' in str(col):
            try:
                # Convert the column to string type
                df[col] = df[col].astype(str)
                print(f"Converting column {col} to string type")
            except Exception as e:
                print(f"Error converting column {col}: {str(e)}")
    
    # Convert date columns to datetime
    if 'Date' in df.columns and df['Date'].dtype != 'datetime64[ns]':
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    
    # Drop rows with missing dates
    if 'Date' in df.columns:
        df = df.dropna(subset=['Date'])
    
    # Convert amount columns to numeric, handling any non-numeric values
    if 'T Amt' in df.columns:
        df['T Amt'] = pd.to_numeric(df['T Amt'], errors='coerce').fillna(0)
    if 'N Amt' in df.columns:
        df['N Amt'] = pd.to_numeric(df['N Amt'], errors='coerce').fillna(0)
    
    # Fill any missing values in categorical columns
    for col in ['Business Head', 'Consultant', 'Client']:
        if col in df.columns:
            df[col] = df[col].fillna('Unknown')
    
    # Add fiscal year and quarter columns based on the date
    if 'Date' in df.columns:
        df['Fiscal Year'] = df['Date'].apply(get_fiscal_year_for_date)
        df['Fiscal Quarter'] = df['Date'].apply(get_fiscal_quarter_for_date)
        df['Year-Month'] = df['Date'].dt.strftime('%Y-%m')
    
    return df

def get_fiscal_year_for_date(date):
    """
    Determine the fiscal year (April to March) for a given date
    Returns a string like "FY 2022-23"
    """
    if date.month < 4:  # Jan-Mar are part of the previous fiscal year
        fiscal_year = f"FY {date.year-1}-{str(date.year)[2:]}"
    else:  # Apr-Dec are part of the current fiscal year
        fiscal_year = f"FY {date.year}-{str(date.year+1)[2:]}"
    return fiscal_year

def get_fiscal_quarter_for_date(date):
    """
    Determine the fiscal quarter (Q1: Apr-Jun, Q2: Jul-Sep, Q3: Oct-Dec, Q4: Jan-Mar)
    """
    if date.month >= 4 and date.month <= 6:
        return 'Q1'
    elif date.month >= 7 and date.month <= 9:
        return 'Q2'
    elif date.month >= 10 and date.month <= 12:
        return 'Q3'
    else:  # Jan-Mar
        return 'Q4'

def get_fiscal_periods(dates):
    """
    Generate a list of fiscal periods (e.g., "FY 2022-23") based on the date range in the data
    """
    # Convert to datetime if not already
    dates = pd.to_datetime(dates)
    
    # Get min and max dates
    min_date = dates.min()
    max_date = dates.max()
    
    # Get the earliest fiscal year
    if min_date.month < 4:
        earliest_fy_start = min_date.year - 1
    else:
        earliest_fy_start = min_date.year
    
    # Get the latest fiscal year
    if max_date.month >= 4:
        latest_fy_start = max_date.year
    else:
        latest_fy_start = max_date.year - 1
    
    # Generate list of fiscal periods
    fiscal_periods = []
    for year in range(earliest_fy_start, latest_fy_start + 1):
        fiscal_period = f"FY {year}-{str(year+1)[2:]}"
        fiscal_periods.append(fiscal_period)
    
    return fiscal_periods

def filter_data(df, business_heads, consultants, clients, fiscal_period):
    """
    Filter the billing data based on selected filters
    """
    # Make a copy of the DataFrame to avoid modifying the original
    filtered_df = df.copy()
    
    # Apply business head filter if selected
    if business_heads:
        filtered_df = filtered_df[filtered_df['Business Head'].isin(business_heads)]
    
    # Apply consultant filter if selected
    if consultants:
        filtered_df = filtered_df[filtered_df['Consultant'].isin(consultants)]
    
    # Apply client filter if selected
    if clients:
        filtered_df = filtered_df[filtered_df['Client'].isin(clients)]
    
    # Apply fiscal period filter if selected
    if fiscal_period:
        filtered_df = filtered_df[filtered_df['Fiscal Year'] == fiscal_period]
    
    return filtered_df
