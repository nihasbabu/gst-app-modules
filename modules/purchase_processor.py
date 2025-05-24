import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime
import os
import pyperclip
import logging
import math

# Set up logging
# Set level to INFO for less output, DEBUG for more detailed output.
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Sheet titles with desired order
SECTION_TITLES = [
    ("PUR-Total", "Purchase Register - Total"),
    ("PUR-Total_sws", "Purchase Register - Total - Supplier wise"),
    ("PUR-Summary-Total", "Purchase Register - Total - Summary"),
]

# Define headers to exclude from total calculation (case-sensitive, matches final_headers)
EXCLUDE_FROM_TOTAL_HEADERS = [
    'GSTIN/UIN of Supplier',
    'Supplier Name',
    'Branch',
    'Supplier Invoice number',
    'Invoice date',
    'Supplier Invoice date',
    'Invoice Type',
    'Voucher number', # Added Voucher number to exclude from total
    'Voucher Ref. No' # Assuming Voucher Ref. No might appear as an extra column and should be excluded
]

# Columns to exclude from Cr/Dr and format checks (these are generally text or date columns)
# This list should contain all fixed headers EXCEPT 'Round Off' (case-sensitive)
exclude_headers_from_crdr_check = [
    'GSTIN/UIN of Supplier',
    'Supplier Name',
    'Branch',
    'Supplier Invoice number',
    'Invoice date',
    'Supplier Invoice date',
    'Invoice Type',
    'Voucher number', # Added Voucher number to exclude from Cr/Dr check
    'Invoice value', # Invoice value is excluded from Cr/Dr check
    'Taxable Value', # Taxable Value is excluded from Cr/Dr check
    'Integrated Tax',
    'Central Tax',
    'State/UT Tax',
    'Cess',
    'Voucher Ref. No' # Assuming Voucher Ref. No might appear as an extra column
]

# Number format to handle negative numbers without red color - Defined globally
STANDARD_NUMBER_FORMAT = r"[>=10000000]##\,##\,##\,##0.00;[>=100000]##\,##\,##0.00;##,##0.00;-;"


def find_header_row(worksheet):
    logging.info("Searching for header row starting with 'Date'")
    for row in worksheet.iter_rows():
        # Case-insensitive check for 'Date' in the first cell
        if isinstance(row[0].value, str) and row[0].value.strip().lower() == "date":
            logging.info(f"Found header row at row {row[0].row}")
            return row[0].row
    logging.warning("Header row not found")
    return None

def get_financial_year(date):
    if date is None:
        return None
    if date.month >= 4:
        return date.year
    return date.year - 1

def create_or_replace_sheet(wb, sheet_name, title_text, columns):
    logging.info(f"Creating or replacing sheet: {sheet_name}")
    if sheet_name in wb.sheetnames:
        logging.info(f"Sheet {sheet_name} exists, removing it")
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
    cell = ws.cell(row=1, column=1)
    cell.value = title_text
    cell.font = Font(bold=True, size=12)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    for idx, col in enumerate(columns, start=1):
        header_cell = ws.cell(row=2, column=idx, value=col)
        header_cell.font = Font(bold=True)
        header_cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        header_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "B3"
    return ws

def apply_format_and_autofit(ws, columns, start_row=3, col_format_map=None):
    logging.info(f"Applying formats and autofitting columns for sheet: {ws.title}")
    for col_idx, col_name in enumerate(columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(col_name))
        if col_format_map and col_name in col_format_map:
            for row in range(start_row, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                # Apply format only if the value is numeric
                if isinstance(cell.value, (int, float)):
                     cell.number_format = col_format_map[col_name]
        for row in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row, column=col_idx).value
            if cell_value is not None:
                max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[col_letter].width = max(15, max_len + 1)
    logging.info("Finished applying formats and autofitting")

def add_total_row(ws, columns, start_row, end_row, STANDARD_NUMBER_FORMAT):
    logging.info(f"Adding total row for sheet: {ws.title}, from row {start_row} to {end_row}")
    total_row_data = ['Total'] + [''] * (len(columns) - 1)
    # Define a number format for totals that handles negative numbers (no red)
    # STANDARD_NUMBER_FORMAT is now passed as an argument

    # Iterate through ALL columns provided to the function (using case-sensitive final_headers)
    for col_idx in range(1, len(columns) + 1):
        # Skip the first column which is the 'Total' label
        if col_idx == 1:
            continue

        # Get the header name for the current column (case-sensitive)
        header_name = columns[col_idx - 1]

        # Check if the header should be excluded from totaling (case-sensitive check)
        if header_name in EXCLUDE_FROM_TOTAL_HEADERS:
            logging.debug(f"Excluding column '{header_name}' from total.")
            continue # Skip totaling for this column

        total = 0 # Initialize total before the try block
        has_numeric_data = False # Flag to check if the column contains any numeric data

        try:
            for row in range(start_row, end_row + 1):
                value = ws.cell(row=row, column=col_idx).value
                # Ensure we are summing numeric values, handle None/NaN/Inf
                if isinstance(value, (int, float)) and not (math.isnan(value) or math.isinf(value)):
                    total += float(value)
                    has_numeric_data = True # Set flag if at least one numeric value is found
                elif value is not None and str(value).strip() != '':
                    # If the value is not numeric and not empty, log a debug message
                    logging.debug(f"Non-numeric value encountered in column '{header_name}' at row {row}: {value}")


            # Only add the total if the column contained at least one numeric value
            if has_numeric_data:
                total_row_data[col_idx - 1] = total
            else:
                # If no numeric data was found in a column that wasn't excluded,
                # leave it blank in the total row.
                pass # Leave as empty string

        except Exception as e:
            # If any error occurs during summing (e.g., unexpected data type), mark as Error
            logging.error(f"Error totaling column '{header_name}': {e}")
            total_row_data[col_idx - 1] = "Error"


    row_num = end_row + 1
    logging.debug(f"Writing total row at row {row_num}")
    for col_idx, value in enumerate(total_row_data, start=1):
        cell = ws.cell(row=row_num, column=col_idx, value=value)
        cell.font = Font(bold=True, color="FF0000") # Make total row bold and red (only the "Total" text)
        # Apply number format only if the value in the total cell is numeric
        if isinstance(value, (int, float)):
            cell.number_format = STANDARD_NUMBER_FORMAT

    logging.info("Total row added successfully")

# Helper function to convert value to float, handling Cr/Dr and None, and checking cell format
# Added 'header' argument to control Cr/Dr and format checks (header is case-sensitive final header)
def safe_float_conversion(value, header, cell=None):
    if value is None:
        return 0.0

    # Columns to exclude from Cr/Dr and format checks (these are generally text or date columns)
    # This list should contain all fixed headers EXCEPT 'Round Off' (case-sensitive)
    # exclude_headers_from_crdr_check is defined globally

    # If the header is in the exclude list for Cr/Dr check (case-sensitive check), just attempt standard float conversion
    if header in exclude_headers_from_crdr_check:
        try:
            # Attempt to convert to float, handle empty string and None
            float_value = float(value) if value != '' and value is not None else 0.0
            # Check if the original value was not numeric but converted to 0.0 from empty/None
            # If it was a string that couldn't be converted, return original value
            if isinstance(value, str) and value.strip() != '' and float_value == 0.0 and value.strip() != '0' and value.strip() != '0.0':
                 return value # Keep original string if it wasn't empty/None and couldn't be converted
            return float_value
        except (ValueError, TypeError):
            logging.debug(f"Could not convert value to float for header '{header}' (standard conversion): {value}")
            return value # Return original value if conversion fails


    # For 'Round Off' and other potentially numeric columns (including extra headers), apply Cr/Dr and format checks
    numeric_value = 0.0

    # Attempt direct conversion to float first
    try:
        numeric_value = float(value)
    except (ValueError, TypeError):
        # If direct conversion fails, try converting string values, handling Cr/Dr (PURCHASE LOGIC)
        if isinstance(value, str):
            value_str = value.strip()
            if value_str.endswith(' Cr'):
                try:
                    # For Purchase, Cr means negative
                    numeric_value = -float(value_str[:-3])
                except ValueError:
                    logging.debug(f"Could not convert Cr string to float for header '{header}': {value}")
                    numeric_value = 0.0
            elif value_str.endswith(' Dr'):
                try:
                    # For Purchase, Dr means positive
                    numeric_value = float(value_str[:-3])
                except ValueError:
                    logging.debug(f"Could not convert Dr string to float for header '{header}': {value}")
                    numeric_value = 0.0
            else:
                try:
                    numeric_value = float(value_str) if value_str else 0.0
                except ValueError:
                    logging.debug(f"Could not convert string to float for header '{header}': {value}")
                    numeric_value = 0.0
        else:
            logging.debug(f"Unexpected value type for conversion for header '{header}': {type(value)}")
            numeric_value = 0.0

    # After attempting conversion, check the cell's number format if the value is numeric
    # This handles cases where the value is stored as positive/negative but formatted as Dr/Cr (PURCHASE LOGIC)
    if cell and cell.number_format and isinstance(numeric_value, (int, float)):
        format_str = str(cell.number_format)
        # Check if the number format contains 'Dr' and the numeric value is negative (should be positive for purchase)
        if 'Dr' in format_str and numeric_value < 0:
             numeric_value = abs(numeric_value) # Make it positive
        # Check if the number format contains 'Cr' and the numeric value is positive (should be negative for purchase)
        elif 'Cr' in format_str and numeric_value > 0:
             numeric_value = -abs(numeric_value) # Make it negative


    return numeric_value


def process_purchase_data(input_files, template_file=None, existing_wb=None):
    logging.info(f"Starting processing with {len(input_files)} input files")
    all_data = []

    # Define the mapping from lowercase source header keys to the desired standardized header names (case-sensitive)
    # This dictionary is used for case-insensitive lookup of known columns.
    source_key_to_standard_header = {
        'gstin/uin': 'GSTIN/UIN of Supplier',
        'particulars': 'Supplier Name',
        'supplier invoice no.': 'Supplier Invoice number',
        'date': 'Invoice date',
        'supplier invoice date': 'Supplier Invoice date',
        'voucher type': 'Invoice Type',
        'voucher no.': 'Voucher number', # Added mapping for 'Voucher No.' to 'Voucher number'
        'gross total': 'Invoice value', # This is the source for 'Invoice value' in detail sheets
        'igst': 'Integrated Tax',
        'cgst': 'Central Tax',
        'sgst': 'State/UT Tax',
        'cess': 'Cess',
        'round off': 'Round Off',
        'voucher ref. no': 'Voucher Ref. No',
        # 'Value', 'PURCHASE GST', 'PURCHASE IGST' are handled as a special case for 'Taxable Value'
        # and are not included in this general mapping.
    }

    # Updated fixed headers - 15 columns (case-sensitive, as required for output)
    fixed_headers = [
        'GSTIN/UIN of Supplier',
        'Supplier Name',
        'Branch',
        'Supplier Invoice number',
        'Invoice date',
        'Supplier Invoice date',
        'Invoice Type',
        'Voucher number', # Included in fixed headers
        'Invoice value',  # This is the 'Invoice value' used in detail sheets
        'Taxable Value', # Taxable Value is a fixed header
        'Integrated Tax',
        'Central Tax',
        'State/UT Tax',
        'Cess',
        'Round Off' # Added Round Off to fixed headers
    ]

    # Create a set of lowercase fixed headers for case-insensitive comparison
    fixed_headers_lower = {h.lower() for h in fixed_headers}

    extra_headers_set = set() # To collect unique extra headers from all files (stores original casing)

    logging.info("Starting file processing loop")
    for filepath, branch_key in input_files:
        logging.info(f"Processing file: {filepath} with branch_key: {branch_key}")
        try:
            wb = load_workbook(filepath)
            logging.info(f"Loaded workbook: {filepath}")

            if len(wb.sheetnames) > 1:
                logging.info(f"Multiple sheets found: {wb.sheetnames}")
                if "Purchase Register" in wb.sheetnames:
                    ws = wb["Purchase Register"]
                    logging.info("Selected 'Purchase Register' sheet")
                else:
                    logging.error(f"'Purchase Register' sheet not found in {filepath}")
                    raise ValueError(f"Multiple sheets found but 'Purchase Register' not present in {filepath}")
            else:
                ws = wb.active
                logging.info("Single sheet found, using active sheet")
                if not ws:
                    logging.error(f"No active sheet in {filepath}")
                    raise ValueError(f"No active sheet found in {filepath}")

            header_row = find_header_row(ws)
            if not header_row:
                logging.error(f"Header row not found in {filepath}")
                raise ValueError(f"Could not find header row starting with 'Date' in {filepath}")
            logging.info(f"Header row found at row {header_row}")

            # Read original headers and create case-insensitive lookup
            original_headers = [cell.value for cell in ws[header_row] if cell.value is not None]
            # Map lowercase header to its original casing and cell object
            original_headers_lower_map = {}
            original_cells_lower_map = {}
            for col_idx, header in enumerate(original_headers):
                 if header is not None:
                      header_lower = str(header).lower()
                      original_headers_lower_map[header_lower] = header
                      original_cells_lower_map[header_lower] = ws.cell(row=header_row, column=col_idx + 1) # Store header row cells too


            logging.debug(f"Original headers: {original_headers}")
            logging.debug(f"Original headers lowercase map: {original_headers_lower_map}")


            # Identify extra headers in the current file (case-insensitive comparison)
            # Iterate through original headers to find those not explicitly mapped or fixed
            for original_header in original_headers:
                if original_header is not None:
                    header_lower = str(original_header).lower()

                    # Check if the lowercase header is a key in the case-insensitive source_key_to_standard_header
                    is_mapped_key = header_lower in source_key_to_standard_header

                    # Check if it's one of the special Taxable Value source keys
                    is_taxable_value_source = header_lower in ['value', 'purchase gst', 'purchase igst']

                    # Check if the original header (or its mapped standard header) is in the fixed headers
                    is_fixed = False
                    if is_mapped_key:
                         standard_header = source_key_to_standard_header[header_lower]
                         if standard_header.lower() in fixed_headers_lower:
                              is_fixed = True
                    elif header_lower in fixed_headers_lower:
                         is_fixed = True


                    # If the header is not a mapped key, not a taxable value source, and not a fixed header, it's an extra header
                    if not is_mapped_key and not is_taxable_value_source and not is_fixed:
                        # Add the original casing to the set of all extra headers
                        extra_headers_set.add(original_header)


            logging.debug(f"Current extra_headers_set (original case): {extra_headers_set}")


            for row_idx, row in enumerate(ws.iter_rows(min_row=header_row + 1), start=header_row + 1):
                # Create a case-insensitive mapping of original headers to cell values and cell objects for the current row
                row_data_orig = {}
                current_row_cells = {}
                row_data_orig_lower_map = {}
                current_row_cells_lower_map = {}

                for i, cell in enumerate(row):
                    if i < len(original_headers) and original_headers[i] is not None:
                        original_header = original_headers[i]
                        row_data_orig[original_header] = cell.value
                        current_row_cells[original_header] = cell
                        header_lower = str(original_header).lower()
                        row_data_orig_lower_map[header_lower] = cell.value
                        current_row_cells_lower_map[header_lower] = cell


                # Skip rows without a valid date or 'Grand Total' rows
                # Use case-insensitive check for 'Date' and 'Particulars' (for 'Grand Total')
                date_value = None
                particulars_value = None
                # Find 'Date' and 'Particulars' case-insensitively
                if 'date' in row_data_orig_lower_map:
                     date_value = row_data_orig_lower_map['date']
                if 'particulars' in row_data_orig_lower_map:
                     particulars_value = row_data_orig_lower_map['particulars']


                if not date_value or (isinstance(particulars_value, str) and 'grand total' in particulars_value.lower()):
                    logging.debug(f"Skipping row {row_idx}: no Date or contains 'Grand Total'")
                    continue

                # Convert date string to datetime object if necessary
                # Use the value retrieved case-insensitively
                if isinstance(date_value, str):
                    try:
                        # Need the original case header to update row_data_orig
                        original_date_header = original_headers_lower_map.get('date')
                        if original_date_header: # Check if original_date_header is not None
                           row_data_orig[original_date_header] = datetime.datetime.strptime(date_value, '%Y-%m-%d %H:%M:%S')
                    except ValueError:
                        logging.debug(f"Skipping row {row_idx} due to date parsing error: {date_value}")
                        continue
                elif not isinstance(date_value, datetime.datetime): # Check if date_value is not a datetime object
                     logging.debug(f"Skipping row {row_idx} due to invalid date type: {type(date_value)}")
                     continue


                # --- Extract and Convert Values for Output ---
                processed_row_data = {}

                # Process Branch separately as it's not from original headers
                processed_row_data['Branch'] = branch_key

                # Iterate through original headers to extract data
                for original_header in original_headers:
                    if original_header is None:
                         continue

                    header_lower = str(original_header).lower()
                    original_value = row_data_orig.get(original_header)
                    cell = current_row_cells.get(original_header)

                    # --- Special handling for Taxable Value extraction priority (case-insensitive lookup) ---
                    if header_lower in ['value', 'purchase gst', 'purchase igst']:
                         # This original header is a potential source for Taxable Value.
                         # We will handle the priority logic below after processing all original headers.
                         continue # Skip processing here, handle in the dedicated Taxable Value logic

                    # --- Handle other columns ---
                    # Find the standard header name for this original header
                    standard_header = source_key_to_standard_header.get(header_lower) # Use get to return None if not found

                    # If the original header maps to a standard header (and it's not Taxable Value)
                    if standard_header and standard_header != 'Taxable Value':
                         # Process the value using safe_float_conversion
                         processed_row_data[standard_header] = safe_float_conversion(original_value, standard_header, cell)
                    # If the original header is not mapped and not a Taxable Value source, it's an extra header
                    elif header_lower not in ['value', 'purchase gst', 'purchase igst']:
                         # Add the original header (case-sensitive) to processed_row_data
                         processed_row_data[original_header] = safe_float_conversion(original_value, original_header, cell)


                # --- Dedicated Taxable Value Extraction Logic (after processing all original headers) ---
                taxable_value_source = None
                taxable_value_cell = None

                # Priority 1: 'Value' (case-insensitive lookup, check for non-zero)
                if 'value' in row_data_orig_lower_map:
                     value_from_source = row_data_orig_lower_map['value']
                     cell_from_source = current_row_cells_lower_map.get('value')
                     converted_value = safe_float_conversion(value_from_source, 'Taxable Value', cell_from_source)
                     # Check if converted value is numeric and non-zero, or a non-empty/non-zero string
                     if (isinstance(converted_value, (int, float)) and converted_value != 0) or \
                        (isinstance(converted_value, str) and converted_value.strip() != '' and converted_value.strip() != '0' and converted_value.strip() != '0.0'):
                         taxable_value_source = value_from_source
                         taxable_value_cell = cell_from_source
                         logging.debug(f"Using 'Value' for Taxable Value (non-zero): {taxable_value_source}")


                # Priority 2: 'PURCHASE GST' if 'Value' not found or zero (case-insensitive lookup, check for non-zero)
                if taxable_value_source is None or (isinstance(safe_float_conversion(taxable_value_source, 'Taxable Value', taxable_value_cell), (int, float)) and safe_float_conversion(taxable_value_source, 'Taxable Value', taxable_value_cell) == 0):
                     if 'purchase gst' in row_data_orig_lower_map:
                          value_from_source = row_data_orig_lower_map['purchase gst']
                          cell_from_source = current_row_cells_lower_map.get('purchase gst')
                          converted_value = safe_float_conversion(value_from_source, 'Taxable Value', cell_from_source)
                          if (isinstance(converted_value, (int, float)) and converted_value != 0) or \
                             (isinstance(converted_value, str) and converted_value.strip() != '' and converted_value.strip() != '0' and converted_value.strip() != '0.0'):
                              taxable_value_source = value_from_source
                              taxable_value_cell = cell_from_source
                              logging.debug(f"Using 'PURCHASE GST' for Taxable Value (non-zero): {taxable_value_source}")


                # Priority 3: 'PURCHASE IGST' if previous not found or zero (case-insensitive lookup, check for non-zero)
                if taxable_value_source is None or (isinstance(safe_float_conversion(taxable_value_source, 'Taxable Value', taxable_value_cell), (int, float)) and safe_float_conversion(taxable_value_source, 'Taxable Value', taxable_value_cell) == 0):
                      if 'purchase igst' in row_data_orig_lower_map:
                           value_from_source = row_data_orig_lower_map['purchase igst']
                           cell_from_source = current_row_cells_lower_map.get('purchase igst')
                           converted_value = safe_float_conversion(value_from_source, 'Taxable Value', cell_from_source)
                           if (isinstance(converted_value, (int, float)) and converted_value != 0) or \
                              (isinstance(converted_value, str) and converted_value.strip() != '' and converted_value.strip() != '0' and converted_value.strip() != '0.0'):
                               taxable_value_source = value_from_source
                               taxable_value_cell = cell_from_source
                               logging.debug(f"Using 'PURCHASE IGST' for Taxable Value (non-zero): {taxable_value_source}")


                # If after all checks, no non-zero source was found, default to None (which safe_float_conversion makes 0.0)
                if taxable_value_source is None:
                     logging.debug("Taxable Value source not found or is zero, defaulting to None.")
                     taxable_value_source = None
                     taxable_value_cell = None # Ensure cell is None too

                # Process the selected source value for Taxable Value using safe_float_conversion
                processed_row_data['Taxable Value'] = safe_float_conversion(taxable_value_source, 'Taxable Value', taxable_value_cell)


                # Ensure all fixed headers are present in processed_row_data, add empty string if missing
                # This handles fixed headers that might not have been in the source file
                for header in fixed_headers:
                    if header not in processed_row_data:
                        processed_row_data[header] = ''


                all_data.append(processed_row_data) # Add the processed row data
                logging.debug(f"Processed row {row_idx} and added to all_data")


        except Exception as e:
            logging.error(f"Error processing file {filepath}: {e}")
            # Depending on desired behavior, you might want to continue or raise the exception
            # raise e
        logging.info(f"Finished processing file: {filepath}")

    logging.info("Finished file processing loop")

    # Construct final headers: fixed headers + sorted unique extra headers from ALL files
    # extra_headers_set already contains original casing
    extra_headers_list = sorted(list(extra_headers_set))
    final_headers = fixed_headers + extra_headers_list

    logging.info(f"Final headers for detail sheets: {final_headers}")
    logging.info(f"Identified unique extra headers (original case): {extra_headers_list}")


    logging.info(f"Total records processed across all files: {len(all_data)}")

    # --- Data Structuring and Sorting ---
    logging.info("Structuring and sorting data for sheets")
    data_by_type = {"PUR-Total": all_data} # PUR-Total uses all data initially

    # Custom sort key for PUR-Total based on Supplier Invoice date and financial year
    def sort_key_pur_total(row):
        date = row.get('Supplier Invoice date')
        if isinstance(date, datetime.datetime):
            # Get financial year (April to March)
            fy = get_financial_year(date)
            # Create a sortable tuple: (financial_year, month, day)
            # This ensures sorting by FY first, then by month and day within the FY
            return (fy, date.month if date.month >= 4 else date.month + 12, date.day)
        # Handle cases where date is not a datetime object (e.g., None, string)
        # Place these at the beginning or end depending on desired behavior.
        # Sorting by datetime.min places them at the beginning.
        return (0, 0, 0) # Or a large tuple to place at the end: (9999, 12, 31)


    logging.info("Sorting data for PUR-Total by Supplier Invoice date (Financial Year)")
    data_by_type["PUR-Total"].sort(key=sort_key_pur_total)


    def sort_key_sws(row):
        supplier = row.get('Supplier Name', '')
        date = row.get('Invoice date', datetime.datetime.min) # Ensure date is a datetime object for comparison
        if not isinstance(date, datetime.datetime): # Handle if date is not datetime
            date = datetime.datetime.min
        if supplier.lower() in ['cash', '(cancelled )'] or not supplier:
            return (1, supplier or '', date)
        return (0, supplier, date)

    logging.info("Sorting data for PUR-Total_sws by Supplier Name and Invoice date")
    data_by_type["PUR-Total_sws"] = sorted(all_data, key=sort_key_sws)
    logging.info("Data structuring and sorting completed")


    # Use existing workbook or load template or create new
    logging.info("Determining workbook to use")
    if existing_wb is not None:
        output_wb = existing_wb
        logging.info("Using provided existing workbook")
    elif template_file:
        logging.info(f"Loading template: {template_file}")
        output_wb = load_workbook(template_file)
    else:
        logging.info("Creating new workbook")
        output_wb = Workbook()
        if 'Sheet' in output_wb.sheetnames:
            logging.info("Removing default 'Sheet'")
            del output_wb['Sheet']

    # Number format to handle negative numbers without red color
    # standard_number_format is now defined globally as STANDARD_NUMBER_FORMAT

    # Apply standard number format to all potentially numeric fixed headers
    col_format_map = {
        'Invoice date': 'DD-MM-YYYY',
        'Supplier Invoice date': 'DD-MM-YYYY',
        'Invoice value': STANDARD_NUMBER_FORMAT,
        'Taxable Value': STANDARD_NUMBER_FORMAT,
        'Integrated Tax': STANDARD_NUMBER_FORMAT,
        'Central Tax': STANDARD_NUMBER_FORMAT,
        'State/UT Tax': STANDARD_NUMBER_FORMAT,
        'Cess': STANDARD_NUMBER_FORMAT,
        'Round Off': STANDARD_NUMBER_FORMAT
        # Formatting for extra headers will be applied based on their data type by Excel
    }


    sheets_to_create = [
        ("PUR-Total", data_by_type.get("PUR-Total", [])),
        ("PUR-Total_sws", data_by_type.get("PUR-Total_sws", [])),
    ]
    logging.info(f"Sheets to create: {[s[0] for s in sheets_to_create]}")

    for sheet_name, data in sheets_to_create:
        logging.info(f"Processing sheet: {sheet_name} with {len(data)} records")
        title = next(t for n, t in SECTION_TITLES if n == sheet_name)
        ws = create_or_replace_sheet(output_wb, sheet_name, title, final_headers)
        start_row = 3
        logging.info(f"Starting data population at row {start_row}")
        for row_idx, row_data in enumerate(data, start=start_row):
            row_values = []
            # Populate row values based on final_headers (case-sensitive)
            for header in final_headers:
                value = row_data.get(header, '')
                # Date formatting is handled here for display
                if header in ['Invoice date', 'Supplier Invoice date'] and isinstance(value, datetime.datetime):
                     value = value.strftime('%d-%m-%Y')
                row_values.append(value)
            ws.append(row_values)
            logging.debug(f"Appended row {row_idx} to {sheet_name}")

        if data:
            logging.info(f"Adding total row for {sheet_name}")
            # Pass STANDARD_NUMBER_FORMAT to add_total_row
            add_total_row(ws, final_headers, start_row, start_row + len(data) - 1, STANDARD_NUMBER_FORMAT)

        # Pass col_format_map to apply_format_and_autofit
        apply_format_and_autofit(ws, final_headers, col_format_map=col_format_map)
        logging.info(f"Created sheet {sheet_name} with {len(data)} records")


    # --- Summary Sheet Processing ---
    summary_sheets = [
        ("PUR-Summary-Total", all_data), # Summary sheets still process the full data for totals
    ]
    logging.info(f"Summary sheets to create: {[s[0] for s in summary_sheets]}")

    # MODIFIED: Added "Invoice Value" to summary_headers
    summary_headers = ['Month', 'No. of Records', 'Invoice Value', 'Taxable Value', 'Integrated Tax',
                       'Central Tax', 'State/UT Tax', 'Cess']
    logging.info(f"Summary headers: {summary_headers}")

    # MODIFIED: Added "Invoice Value" to summary_col_format_map
    summary_col_format_map = {
        'Invoice Value': STANDARD_NUMBER_FORMAT, # Added formatting for Invoice Value
        'Taxable Value': STANDARD_NUMBER_FORMAT,
        'Integrated Tax': STANDARD_NUMBER_FORMAT,
        'Central Tax': STANDARD_NUMBER_FORMAT,
        'State/UT Tax': STANDARD_NUMBER_FORMAT,
        'Cess': STANDARD_NUMBER_FORMAT
    }
    logging.debug(f"Summary column format map: {summary_col_format_map}")

    months = ['April', 'May', 'June', 'July', 'August', 'September',
              'October', 'November', 'December', 'January', 'February', 'March']
    logging.debug(f"Month order for summaries: {months}")

    for sheet_name, data in summary_sheets:
        logging.info(f"Processing summary sheet: {sheet_name}")
        title = next(t for n, t in SECTION_TITLES if n == sheet_name)
        ws = create_or_replace_sheet(output_wb, sheet_name, title, summary_headers)
        summary_data = {}
        invoice_numbers = set()
        logging.debug("Initialized summary_data and invoice_numbers")

        for row in data:
            date = row.get('Invoice date')
            inv_num = row.get('Supplier Invoice number')
            if inv_num is None or date is None:
                logging.debug("Skipping summary row: no supplier invoice number or date")
                continue
            if isinstance(date, datetime.datetime):
                fy = get_financial_year(date)
                month_idx = (date.month - 4) % 12
                month = months[month_idx]
                if month not in summary_data:
                    # MODIFIED: Initialize 'Invoice Value' in summary_data
                    summary_data[month] = {
                        'count': 0, 'Invoice Value': 0, 'Taxable Value': 0, 'Integrated Tax': 0,
                        'Central Tax': 0, 'State/UT Tax': 0, 'Cess': 0
                    }
                # Count unique invoice numbers per month
                if (month, inv_num) not in invoice_numbers:
                     summary_data[month]['count'] += 1
                     invoice_numbers.add((month, inv_num))

                # MODIFIED: Added 'Invoice value' to the list of fields to sum
                # Note: The field from `all_data` (which `row` comes from) is 'Invoice value' (lowercase v)
                # This matches the `fixed_headers` and `source_key_to_standard_header` mapping for 'gross total'.
                for field in ['Invoice value', 'Taxable Value', 'Integrated Tax', 'Central Tax', 'State/UT Tax', 'Cess']:
                    value = row.get(field, 0) # Get the value for the current field
                    # The key in summary_data for 'Invoice value' from row should be 'Invoice Value' (capital V)
                    summary_field_key = 'Invoice Value' if field == 'Invoice value' else field

                    if isinstance(value, (int, float)) and not (math.isnan(value) or math.isinf(value)):
                       summary_data[month][summary_field_key] += float(value)


        start_row = 3
        row_count = 0
        logging.info("Populating summary sheet rows")
        for month in months:
            if month in summary_data:
                # MODIFIED: Append 'Invoice Value' to the row
                ws.append([
                    month,
                    summary_data[month]['count'],
                    summary_data[month]['Invoice Value'], # Added Invoice Value
                    summary_data[month]['Taxable Value'],
                    summary_data[month]['Integrated Tax'],
                    summary_data[month]['Central Tax'],
                    summary_data[month]['State/UT Tax'],
                    summary_data[month]['Cess']
                ])
                row_count += 1
                logging.debug(f"Appended summary row for {month}")

        if summary_data:
            logging.info("Adding total row for summary sheet")
            total_row = ['Total']
            # MODIFIED: The loop for total_row calculation should correctly sum all numeric columns
            # based on the updated summary_headers.
            for col_idx in range(1, len(summary_headers)): # Iterate based on new summary_headers length
                total = 0 # Initialize total for summing column values in summary sheet
                for r_idx in range(start_row, start_row + row_count): # Iterate through data rows
                    cell_to_sum = ws.cell(row=r_idx, column=col_idx + 1) # Get cell for current column
                    value = cell_to_sum.value
                    if isinstance(value, (int, float)) and not (math.isnan(value) or math.isinf(value)):
                       total += float(value)
                total_row.append(total)
            ws.append(total_row)
            logging.debug(f"Writing total row at row {start_row + row_count}")
            for col_idx in range(1, len(summary_headers) + 1):
                cell = ws.cell(row=start_row + row_count, column=col_idx)
                cell.font = Font(bold=True, color="FF0000")
                # Apply number format to all numeric total cells (except the 'Total' label)
                if col_idx > 1 and isinstance(cell.value, (int, float)):
                    cell.number_format = STANDARD_NUMBER_FORMAT # Use global variable

        apply_format_and_autofit(ws, summary_headers, col_format_map=summary_col_format_map)
        logging.info(f"Created summary sheet {sheet_name}")

    logging.info("Purchase data processing completed")
    return output_wb


def select_files_and_process():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    input_files_data = []
    while True:
        filepath = filedialog.askopenfilename(
            title="Select Excel File (or Cancel to finish)",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if not filepath:  # User cancelled
            break

        # Simple dialog to get branch key
        branch_key = tk.simpledialog.askstring("Input", f"Enter Branch Key for {os.path.basename(filepath)}:", parent=root)
        if branch_key is None: # User cancelled branch input
            messagebox.showinfo("Cancelled", f"File selection cancelled for {os.path.basename(filepath)}.")
            continue # Skip this file and ask for next

        input_files_data.append((filepath, branch_key))
        messagebox.showinfo("File Added", f"{os.path.basename(filepath)} added with branch key '{branch_key}'. Select another file or cancel.")


    if not input_files_data:
        messagebox.showinfo("No Files Selected", "No files were selected for processing.")
        logging.info("No input files selected.")
        return

    logging.info(f"Selected {len(input_files_data)} files for processing.")

    output_file = filedialog.asksaveasfilename(
        title="Save Consolidated Report As",
        defaultextension=".xlsx",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )

    if not output_file:
        messagebox.showinfo("Cancelled", "Process cancelled by user.")
        logging.info("Output file saving cancelled.")
        return

    try:
        logging.info(f"Processing data and saving to: {output_file}")
        # For now, we are not using a template or existing workbook for simplicity in this example.
        # You can re-introduce template_file or existing_wb logic if needed.
        final_wb = process_purchase_data(input_files_data)
        final_wb.save(output_file)
        messagebox.showinfo("Success", f"Consolidated report saved to {output_file}")
        logging.info("Processing successful.")
        # Copy output path to clipboard
        try:
            pyperclip.copy(output_file)
            logging.info("Output file path copied to clipboard.")
        except pyperclip.PyperclipException as e:
            logging.warning(f"Could not copy to clipboard: {e}. Pyperclip might not be configured for your system or you might be running in a headless environment.")


    except Exception as e:
        logging.error(f"An error occurred: {e}", exc_info=True) # Log full traceback
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    # Create a root window for dialogs if not already created by another part of a larger app
    # This helps prevent issues with Tkinter dialogs on some systems.
    root = tk.Tk()
    root.withdraw() # Hide the main Tkinter window

    # Setup console logging for debugging, if not already configured
    if not logging.getLogger().handlers: # Check if handlers are already set
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO) # Or DEBUG for more verbose output
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(formatter)
        logging.getLogger().addHandler(console_handler)
        logging.getLogger().setLevel(logging.INFO) # Set root logger level

    select_files_and_process()

    # Ensure Tkinter main loop is properly handled if it was used.
    # For simple dialogs, explicit mainloop isn't always needed if withdraw() is used.
    # However, if you had a visible GUI, you'd call root.mainloop() here.
    # Since we only use dialogs and withdraw the root, this is generally fine.
