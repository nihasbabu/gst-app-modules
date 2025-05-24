import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime
import os
import pyperclip  # Not used in the provided snippet, but kept as it was in the original
import logging
import math

# Set up logging
# Set level to DEBUG to see all messages. Change to INFO or WARNING for less output.
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Sheet titles with desired order
SECTION_TITLES = [
    ("SALE-Total", "Sales Register - Total"),
    ("SALE-Total_sws", "Sales Register - Total - Receiver wise"),
    ("SALE-B2B", "Sales Register - B2B only"),
    ("SALE-B2C", "Sales Register - B2C only"),
    ("SALE-Others", "Sales Register - Others"),
    ("SALE-Summary-Total", "Sales Register - Total - Summary"),
    ("SALE-Summary-B2B", "Sales Register - B2B only - Summary"),
    ("SALE-Summary-B2C", "Sales Register - B2C only - Summary"),
    ("SALE-Summary-Others", "Sales Register - Others - Summary"),
]

# Define headers to exclude from total calculation
EXCLUDE_FROM_TOTAL_HEADERS = [
    'GSTIN/UIN of Recipient',
    'Receiver Name',
    'Branch',
    'Invoice number',
    'Invoice date',
    'Invoice Type',
    'Voucher Ref. No'
]


def find_header_row(worksheet):
    logging.debug(f"Searching for header row in sheet: {worksheet.title}")
    for row in worksheet.iter_rows():
        if row[0].value == "Date":
            logging.debug(f"Header row found at row: {row[0].row}")
            return row[0].row
    logging.debug("Header row not found.")
    return None


def get_financial_year(date):
    if date is None:
        return None
    if date.month >= 4:
        return date.year
    return date.year - 1


def create_or_replace_sheet(wb, sheet_name, title_text, columns):
    logging.info(f"Creating sheet: {sheet_name}")
    if sheet_name in wb.sheetnames:
        logging.debug(f"Sheet '{sheet_name}' already exists, deleting.")
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
    ws.freeze_panes = "B3"  # Assuming first column might be row numbers or similar, freezing from B
    logging.info(f"Finished creating sheet: {sheet_name}")
    return ws


def apply_format_and_autofit(ws, columns, start_row=3, col_format_map=None):
    logging.debug(f"Applying format and autofit for sheet: {ws.title}")
    for col_idx, col_name in enumerate(columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(col_name))  # Start with header length
        if col_format_map and col_name in col_format_map:
            for row_num in range(start_row, ws.max_row + 1):  # Iterate data rows
                cell = ws.cell(row=row_num, column=col_idx)
                # Apply format only if the value is numeric
                if isinstance(cell.value, (int, float)):
                    cell.number_format = col_format_map[col_name]
        # Check header row (row 2) and data rows for max_len
        for row_num in range(2, ws.max_row + 1):  # Iterate from header row
            cell_value = ws.cell(row=row_num, column=col_idx).value
            if cell_value is not None:
                max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[col_letter].width = max(15, max_len + 2)  # Add a little padding
    logging.debug(f"Finished applying format and autofit for sheet: {ws.title}")


def add_total_row(ws, columns, start_row, end_row):
    logging.debug(f"Adding total row for sheet: {ws.title}")
    total_row_data = ['Total'] + [''] * (len(columns) - 1)
    # This format is specific for detail sheet totals and remains unchanged as per request.
    total_number_format = r"[>=10000000]##\,##\,##\,##0.00;[>=100000]##\,##\,##0.00;##,##0.00;-;"

    for col_idx in range(1, len(columns) + 1):
        if col_idx == 1:  # Skip the 'Total' label column
            continue

        header_name = columns[col_idx - 1]
        if header_name in EXCLUDE_FROM_TOTAL_HEADERS:
            logging.debug(f"Excluding column '{header_name}' from total.")
            continue

        total = 0
        has_numeric_data = False
        try:
            for row in range(start_row, end_row + 1):
                value = ws.cell(row=row, column=col_idx).value
                if isinstance(value, (int, float)) and not (math.isnan(value) or math.isinf(value)):
                    total += float(value)
                    has_numeric_data = True
                elif value is not None and str(value).strip() != '':
                    logging.debug(f"Non-numeric value encountered in column '{header_name}' at row {row}: {value}")

            if has_numeric_data:
                total_row_data[col_idx - 1] = total
        except Exception as e:
            logging.error(f"Error totaling column '{header_name}': {e}")
            total_row_data[col_idx - 1] = "Error"

    row_num = end_row + 1
    ws.append(total_row_data)  # Append the list directly
    for col_idx, value in enumerate(total_row_data, start=1):
        cell = ws.cell(row=row_num, column=col_idx)
        cell.font = Font(bold=True, color="FF0000")
        if isinstance(value, (int, float)):
            cell.number_format = total_number_format  # Use the specific total format for detail sheets
    logging.debug(f"Finished adding total row for sheet: {ws.title}")


def safe_float_conversion(value, header, cell=None):
    if value is None:
        return 0.0

    exclude_headers_from_crdr_check = [
        'GSTIN/UIN of Recipient',
        'Receiver Name',
        'Branch',
        'Invoice number',
        'Invoice date',
        'Invoice Type',
        'Voucher Ref. No',
        'Invoice value'
    ]

    if header in exclude_headers_from_crdr_check:
        try:
            # Attempt to convert to float, handle empty string and None
            float_value = float(value) if str(value).strip() != '' and value is not None else 0.0
            # Check if the original value was not numeric but converted to 0.0 from empty/None
            # If it was a string that couldn't be converted, return original value
            if isinstance(value,
                          str) and value.strip() != '' and float_value == 0.0 and value.strip() != '0' and value.strip() != '0.0':
                return value
            return float_value
        except (ValueError, TypeError):
            logging.debug(f"Could not convert value to float for header '{header}' (standard conversion): {value}")
            return value

            # For 'Round Off' and other potentially numeric columns (including extra headers), apply Cr/Dr and format checks
    numeric_value = 0.0

    # Attempt direct conversion to float first
    try:
        numeric_value = float(value)
    except (ValueError, TypeError):
        # If direct conversion fails, try converting string values, handling Cr/Dr
        if isinstance(value, str):
            value_str = value.strip()
            if value_str.endswith(' Cr'):
                try:
                    numeric_value = float(value_str[:-3])
                except ValueError:
                    logging.debug(f"Could not convert Cr string to float for header '{header}': {value}")
                    numeric_value = 0.0  # Default to 0.0 if conversion fails
            elif value_str.endswith(' Dr'):
                try:
                    numeric_value = -float(value_str[:-3])
                except ValueError:
                    logging.debug(f"Could not convert Dr string to float for header '{header}': {value}")
                    numeric_value = 0.0  # Default to 0.0 if conversion fails
            else:
                try:
                    numeric_value = float(value_str) if value_str else 0.0
                except ValueError:
                    logging.debug(f"Could not convert string to float for header '{header}': {value}")
                    numeric_value = 0.0  # Default to 0.0 if conversion fails
        else:
            logging.debug(f"Unexpected value type for conversion for header '{header}': {type(value)}")
            numeric_value = 0.0  # Default to 0.0 for unexpected types

    # After attempting conversion, check the cell's number format if the value is numeric
    if cell and cell.number_format and isinstance(numeric_value, (int, float)):
        if 'Dr' in str(cell.number_format) and numeric_value > 0:  # Convert format to string for check
            numeric_value *= -1
    return numeric_value


def process_excel_data(input_files, template_file=None, existing_wb=None):
    logging.info("Starting sales data processing")
    all_data = []
    # data_by_type will be populated after all_data is collected and processed

    column_mapping = {
        'GSTIN/UIN': 'GSTIN/UIN of Recipient',
        'Particulars': 'Receiver Name',
        'Voucher No': 'Invoice number',
        'Voucher No.': 'Invoice number',
        'Date': 'Invoice date',
        'Gross Total': 'Invoice value',  # Maps source 'Gross Total' to 'Invoice value'
        'Voucher Type': 'Invoice Type',
        'Value': 'Taxable Value',
        'IGST': 'Integrated Tax',
        'CGST': 'Central Tax',
        'SGST': 'State/UT Tax',
        'Cess': 'Cess',
        'ROUND OFF': 'Round Off',
        'Round Off': 'Round Off',
        'Voucher Ref. No': 'Voucher Ref. No'
    }

    fixed_headers = [
        'GSTIN/UIN of Recipient',
        'Receiver Name',
        'Branch',
        'Invoice number',
        'Invoice date',
        'Invoice Type',
        'Invoice value',  # This is the key used internally for invoice's total value
        'Taxable Value',
        'Integrated Tax',
        'Central Tax',
        'State/UT Tax',
        'Cess',
        'Round Off'
    ]

    extra_headers_set = set()

    logging.debug("Starting file processing loop")
    for filepath, branch_key in input_files:
        logging.info(f"Processing sales file: {filepath}")
        try:
            wb = load_workbook(filepath, data_only=True)  # data_only=True to get values not formulas
            if len(wb.sheetnames) > 1:
                if "Sales Register" in wb.sheetnames:
                    ws = wb["Sales Register"]
                else:
                    logging.error(f"'Sales Register' sheet not found in {filepath}")
                    # Consider raising a more specific error or using messagebox here if GUI context is available
                    raise ValueError(f"Multiple sheets found but 'Sales Register' not present in {filepath}")
            else:
                ws = wb.active
                if not ws:
                    logging.error(f"No active sheet in {filepath}")
                    raise ValueError(f"No active sheet found in {filepath}")

            header_row_num = find_header_row(ws)
            if not header_row_num:
                logging.error(f"Header row not found in {filepath}")
                raise ValueError(f"Could not find header row starting with 'Date' in {filepath}")

            headers_from_sheet = [cell.value for cell in ws[header_row_num] if cell.value is not None]
            logging.debug(f"Headers from sheet '{filepath}': {headers_from_sheet}")

            # Identify extra headers in the current file
            current_file_extra_headers = [
                h for h in headers_from_sheet
                if h not in column_mapping.keys() and h not in column_mapping.values()  # Not a key or value in mapping
                   and h not in fixed_headers  # Not one of the fixed headers
                   and h not in ['Others-Cr', 'Others-Dr']  # Not the known Others columns
            ]
            extra_headers_set.update(current_file_extra_headers)  # Add to the set of all extra headers

            for row_idx, row_cells_tuple in enumerate(ws.iter_rows(min_row=header_row_num + 1, values_only=False),
                                                      start=header_row_num + 1):
                # Build row_data_orig_values and current_row_cells_map
                row_data_orig_values = {}
                current_row_cells_map = {}

                for i, cell_obj in enumerate(row_cells_tuple):
                    if i < len(headers_from_sheet):  # Ensure we don't go out of bounds for headers
                        header = headers_from_sheet[i]
                        row_data_orig_values[header] = cell_obj.value
                        current_row_cells_map[header] = cell_obj

                        # Skip rows without a valid date or 'Grand Total' rows
                if not row_data_orig_values.get('Date') or 'Grand Total' in str(
                        row_data_orig_values.get('Particulars', '')):
                    continue

                # Convert date string to datetime object if necessary
                date_val = row_data_orig_values.get('Date')
                if isinstance(date_val, str):
                    try:
                        # Try parsing common datetime format first
                        row_data_orig_values['Date'] = datetime.datetime.strptime(date_val, '%Y-%m-%d %H:%M:%S')
                    except ValueError:
                        try:
                            # Fallback to try parsing date-only format (DD-MM-YYYY)
                            row_data_orig_values['Date'] = datetime.datetime.strptime(date_val.split()[0], '%d-%m-%Y')
                        except ValueError:
                            logging.debug(f"Skipping row {row_idx} due to date parsing error: {date_val}")
                            continue
                elif not isinstance(date_val, datetime.datetime):
                    logging.debug(f"Skipping row {row_idx} due to invalid date type: {type(date_val)}")
                    continue

                # --- Extract and Convert Values for Output ---
                processed_row_data = {}
                for original_header, original_value in row_data_orig_values.items():
                    mapped_header = column_mapping.get(original_header, original_header)
                    cell_for_conversion = current_row_cells_map.get(original_header)  # Get cell for format check

                    # Consolidate numeric conversion
                    if mapped_header in ['Invoice value', 'Taxable Value', 'Integrated Tax', 'Central Tax',
                                         'State/UT Tax', 'Cess',
                                         'Round Off'] or original_header in current_file_extra_headers:
                        processed_row_data[mapped_header] = safe_float_conversion(original_value, mapped_header,
                                                                                  cell_for_conversion)

                    elif original_header == 'Others-Dr':  # Handle Others-Dr specifically if needed
                        processed_row_data[original_header] = -abs(
                            safe_float_conversion(original_value, original_header, cell_for_conversion))
                    elif original_header == 'Others-Cr':  # Handle Others-Cr specifically if needed
                        processed_row_data[original_header] = abs(
                            safe_float_conversion(original_value, original_header, cell_for_conversion))

                    # For non-numeric fixed headers (or those already handled by mapping to numeric ones)
                    elif mapped_header in fixed_headers:
                        if mapped_header == 'Invoice date':  # Date is already converted
                            processed_row_data[mapped_header] = row_data_orig_values['Date']
                        else:  # Other non-numeric fixed headers
                            processed_row_data[mapped_header] = original_value
                    elif mapped_header == 'Voucher Ref. No':  # Explicitly handle Voucher Ref. No
                        processed_row_data[mapped_header] = original_value

                # Ensure all fixed headers are present in processed_row_data, add empty string if missing
                for header in fixed_headers:
                    if header not in processed_row_data:
                        processed_row_data[header] = ''
                        # Ensure all identified extra headers (from this file) are present
                for header in current_file_extra_headers:
                    if header not in processed_row_data:
                        processed_row_data[header] = ''

                # Store original Others-Cr/Dr if they were in the source (for potential later use, though not outputted in final_headers)
                if 'Others-Cr' in row_data_orig_values and 'Others-Cr' not in processed_row_data:  # Check if not already processed as an extra_header
                    processed_row_data['Others-Cr'] = safe_float_conversion(row_data_orig_values['Others-Cr'],
                                                                            'Others-Cr',
                                                                            current_row_cells_map.get('Others-Cr'))
                if 'Others-Dr' in row_data_orig_values and 'Others-Dr' not in processed_row_data:
                    processed_row_data['Others-Dr'] = safe_float_conversion(row_data_orig_values['Others-Dr'],
                                                                            'Others-Dr',
                                                                            current_row_cells_map.get('Others-Dr'))

                processed_row_data['Branch'] = branch_key  # Add branch key
                all_data.append(processed_row_data)  # Add the processed row data

        except Exception as e:
            logging.error(f"Error processing file {filepath}: {e}", exc_info=True)
            messagebox.showerror("File Processing Error",
                                 f"Error processing file {os.path.basename(filepath)}: {e}\n\nPlease check the logs for more details.")
            # Optionally re-raise or return an indicator of failure if this function is part of a larger flow
            # return None # Or some error object
        logging.info(f"Finished processing sales file: {filepath}")

    logging.debug("Finished file processing loop")

    # Construct final headers: fixed headers + sorted unique extra headers (collected from all files)
    extra_headers_list = sorted(list(extra_headers_set))
    final_headers = fixed_headers + extra_headers_list
    logging.info(f"Final headers for detail sheets: {final_headers}")

    logging.info("Invoice Value validation removed as requested.")

    # --- Data Sorting and Categorization ---
    data_by_type = {"B2B": [], "B2C": [], "Others": []}
    for row_data in all_data:
        inv_type = str(row_data.get('Invoice Type', '')).strip()  # Ensure string and strip spaces
        if inv_type == "B2B":
            data_by_type["B2B"].append(row_data)
        elif inv_type == "B2C":
            data_by_type["B2C"].append(row_data)
        else:  # Includes empty, None, or any other type
            data_by_type["Others"].append(row_data)
            if inv_type not in ["B2B", "B2C", ""]:  # Log if it's an unexpected non-empty type
                logging.debug(
                    f"Row categorized as 'Others' due to Invoice Type: '{inv_type}' from Invoice: {row_data.get('Invoice number')}")

    # Sort each category by date then invoice number
    for key, data_list in data_by_type.items():
        logging.debug(f"Sorting data for category: {key}")
        # Ensure sorting key handles potential None or non-datetime values gracefully
        data_list.sort(key=lambda x: (x.get('Invoice date', datetime.datetime.min) if isinstance(x.get('Invoice date'),
                                                                                                 datetime.datetime) else datetime.datetime.min,
                                      str(x.get('Invoice number', ''))))

    # Sort all_data (for SALE-Total) by date then invoice number
    all_data.sort(key=lambda x: (x.get('Invoice date', datetime.datetime.min) if isinstance(x.get('Invoice date'),
                                                                                            datetime.datetime) else datetime.datetime.min,
                                 str(x.get('Invoice number', ''))))

    logging.debug("Completed data sorting")

    # Sorting for SALE-Total_sws (Receiver Wise)
    def sort_key_sws(row):
        receiver = str(row.get('Receiver Name', '')).strip().lower()  # Normalize receiver name
        date = row.get('Invoice date', datetime.datetime.min)
        inv_num = str(row.get('Invoice number', ''))  # Ensure invoice number is part of sort for consistency
        # Group specific receiver names like 'cash' or '(cancelled )' together
        if receiver in ['cash', '(cancelled )'] or not receiver:  # also group empty receiver names
            return (1, receiver, date, inv_num)  # Secondary sort by receiver, then date, then invoice
        return (0, receiver, date, inv_num)  # Primary sort by receiver, then date, then invoice

    logging.debug("Starting sorting for SALE-Total_sws")
    all_data_sws = sorted(all_data, key=sort_key_sws)  # Sort a copy of all_data

    logging.debug("Completed sorting for SALE-Total_sws")

    # --- Workbook Creation/Loading ---
    if existing_wb is not None:
        output_wb = existing_wb
        logging.debug("Using existing workbook.")
    elif template_file:  # This path might not be hit if GUI loads template into existing_wb
        logging.debug(f"Loading template file: {template_file}")
        output_wb = load_workbook(template_file)
    else:
        logging.debug("Creating new workbook.")
        output_wb = Workbook()
        if 'Sheet' in output_wb.sheetnames and len(
                output_wb.sheetnames) == 1:  # Only delete if it's the default 'Sheet'
            del output_wb['Sheet']

    # MODIFIED: Define the standard number format using the Indian numbering system
    standard_number_format = r"[>=10000000]##\,##\,##\,##0.00;[>=100000]##\,##\,##0.00;##,##0.00;-;"

    col_format_map = {
        'Invoice date': 'DD-MM-YYYY',
        'Invoice value': standard_number_format,
        'Taxable Value': standard_number_format,
        'Integrated Tax': standard_number_format,
        'Central Tax': standard_number_format,
        'State/UT Tax': standard_number_format,
        'Cess': standard_number_format,
        'Round Off': standard_number_format
    }
    # Apply standard_number_format to any extra headers that might be numeric
    for eh in extra_headers_list:
        # This is a heuristic. If an extra column is known to be non-numeric, it shouldn't get this format.
        # However, since extra columns are often numeric ledger values, this is a common case.
        if eh not in col_format_map:  # Avoid overwriting specific formats if any were defined for extra headers
            col_format_map[eh] = standard_number_format

    # --- Detail Sheet Creation ---
    sheets_to_create = [
        ("SALE-Total", all_data),
        ("SALE-Total_sws", all_data_sws),
        ("SALE-B2B", data_by_type.get("B2B", [])),
        ("SALE-B2C", data_by_type.get("B2C", [])),
        ("SALE-Others", data_by_type.get("Others", [])),
    ]

    logging.debug("Starting main sheet creation loop")
    for sheet_name_key, data_list_for_sheet in sheets_to_create:
        if data_list_for_sheet:
            logging.info(f"Starting population of sheet: {sheet_name_key}")
            display_title = next((t for n, t in SECTION_TITLES if n == sheet_name_key), sheet_name_key)

            ws = create_or_replace_sheet(output_wb, sheet_name_key, display_title, final_headers)
            current_start_row = 3  # Data starts at row 3, headers at row 2

            for row_data_item in data_list_for_sheet:
                row_values_to_append = []
                for header_name in final_headers:
                    value = row_data_item.get(header_name, '')  # Default to empty string if key is missing
                    if header_name == 'Invoice date' and isinstance(value, datetime.datetime):
                        value = value.strftime('%d-%m-%Y')  # Format date for display
                    row_values_to_append.append(value)
                ws.append(row_values_to_append)

            if data_list_for_sheet:
                # ws.max_row should now correctly point to the last data row before adding total
                add_total_row(ws, final_headers, current_start_row, ws.max_row)
            apply_format_and_autofit(ws, final_headers, col_format_map=col_format_map, start_row=current_start_row)
            logging.info(f"Completed population of sheet: {sheet_name_key}")
        else:
            logging.info(f"No data for sheet: {sheet_name_key}, skipping creation.")
    logging.debug("Finished main sheet creation loop")

    # --- Summary Sheet Creation ---
    summary_sheets_data_map = [
        ("SALE-Summary-Total", all_data),
        ("SALE-Summary-B2B", data_by_type.get("B2B", [])),
        ("SALE-Summary-B2C", data_by_type.get("B2C", [])),  # Typo fixed: B2C
        ("SALE-Summary-Others", data_by_type.get("Others", [])),
    ]

    summary_headers = ['Month', 'No. of Records', 'Invoice Value', 'Taxable Value', 'Integrated Tax',
                       'Central Tax', 'State/UT Tax', 'Cess']

    summary_col_format_map = {
        'Invoice Value': standard_number_format,  # Use the new Indian format
        'Taxable Value': standard_number_format,  # Use the new Indian format
        'Integrated Tax': standard_number_format,  # Use the new Indian format
        'Central Tax': standard_number_format,  # Use the new Indian format
        'State/UT Tax': standard_number_format,  # Use the new Indian format
        'Cess': standard_number_format  # Use the new Indian format
    }
    months_order = ['April', 'May', 'June', 'July', 'August', 'September',
                    'October', 'November', 'December', 'January', 'February', 'March']

    logging.debug("Starting summary sheet creation loop")
    for sheet_name_key, data_for_summary in summary_sheets_data_map:
        if data_for_summary:
            logging.info(f"Starting population of summary sheet: {sheet_name_key}")
            display_title = next((t for n, t in SECTION_TITLES if n == sheet_name_key), sheet_name_key)

            ws_summary = create_or_replace_sheet(output_wb, sheet_name_key, display_title, summary_headers)

            monthly_summary_data = {}
            unique_invoices_by_month = {}

            for row_item in data_for_summary:
                date_obj = row_item.get('Invoice date')
                # Ensure invoice number is treated as a string for uniqueness check
                invoice_num_val = str(row_item.get('Invoice number', '')).strip()

                if not isinstance(date_obj,
                                  datetime.datetime) or not invoice_num_val:  # Skip if no date or invoice number
                    logging.debug(
                        f"Skipping summary calculation for row due to missing date/invoice: {row_item.get('Invoice number')}")
                    continue

                month_index = (date_obj.month - 4 + 12) % 12
                month_name = months_order[month_index]

                if month_name not in monthly_summary_data:
                    monthly_summary_data[month_name] = {
                        'count': 0, 'Invoice Value': 0.0, 'Taxable Value': 0.0,
                        'Integrated Tax': 0.0, 'Central Tax': 0.0,
                        'State/UT Tax': 0.0, 'Cess': 0.0
                    }
                    unique_invoices_by_month[month_name] = set()

                if invoice_num_val not in unique_invoices_by_month[month_name]:
                    monthly_summary_data[month_name]['count'] += 1
                    unique_invoices_by_month[month_name].add(invoice_num_val)

                fields_to_sum_in_summary = {
                    'Invoice value': 'Invoice Value',  # Detail key 'Invoice value' maps to summary 'Invoice Value'
                    'Taxable Value': 'Taxable Value',
                    'Integrated Tax': 'Integrated Tax',
                    'Central Tax': 'Central Tax',
                    'State/UT Tax': 'State/UT Tax',
                    'Cess': 'Cess'
                }
                for detail_key, summary_key in fields_to_sum_in_summary.items():
                    value_to_add = row_item.get(detail_key, 0.0)  # Default to 0.0 if key missing
                    if isinstance(value_to_add, (int, float)) and not (
                            math.isnan(value_to_add) or math.isinf(value_to_add)):
                        monthly_summary_data[month_name][summary_key] += float(value_to_add)
                    elif isinstance(value_to_add, str) and value_to_add.strip() == '':
                        pass  # Empty string treated as zero for sum
                    elif value_to_add is not None and str(value_to_add).strip() != '':  # Log non-empty, non-numeric
                        logging.warning(
                            f"Non-numeric or invalid value '{value_to_add}' for '{detail_key}' in summary for invoice '{invoice_num_val}', treating as 0.")

            summary_start_row = 3
            rows_added_to_summary = 0
            for month_n in months_order:
                if month_n in monthly_summary_data:
                    month_data = monthly_summary_data[month_n]
                    ws_summary.append([
                        month_n,
                        month_data['count'],
                        month_data['Invoice Value'],
                        month_data['Taxable Value'],
                        month_data['Integrated Tax'],
                        month_data['Central Tax'],
                        month_data['State/UT Tax'],
                        month_data['Cess']
                    ])
                    rows_added_to_summary += 1

            if rows_added_to_summary > 0:
                summary_total_row_values = ['Total'] + [0.0] * (len(summary_headers) - 1)
                for col_idx_summary in range(1, len(summary_headers)):
                    # header_for_sum = summary_headers[col_idx_summary] # Not strictly needed for sum logic here
                    current_col_total = 0.0
                    for r_idx in range(summary_start_row, summary_start_row + rows_added_to_summary):
                        cell_val = ws_summary.cell(row=r_idx, column=col_idx_summary + 1).value
                        if isinstance(cell_val, (int, float)) and not (math.isnan(cell_val) or math.isinf(cell_val)):
                            current_col_total += float(cell_val)
                    summary_total_row_values[col_idx_summary] = current_col_total

                ws_summary.append(summary_total_row_values)
                total_row_number_on_sheet = summary_start_row + rows_added_to_summary  # This is the row num of the total line
                for c_idx, val in enumerate(summary_total_row_values, start=1):
                    total_cell = ws_summary.cell(row=total_row_number_on_sheet, column=c_idx)
                    total_cell.font = Font(bold=True, color="FF0000")
                    # Apply standard_number_format to summed columns in the total row
                    if isinstance(val, (int, float)) and summary_headers[c_idx - 1] != 'Month' and summary_headers[
                        c_idx - 1] != 'No. of Records':
                        total_cell.number_format = standard_number_format  # Use the new Indian format

            apply_format_and_autofit(ws_summary, summary_headers, col_format_map=summary_col_format_map,
                                     start_row=summary_start_row)
            logging.info(f"Completed population of summary sheet: {sheet_name_key}")
        else:
            logging.info(f"No data for summary sheet: {sheet_name_key}, skipping creation.")
    logging.debug("Finished summary sheet creation loop")

    logging.info("Completed sales data processing")
    return output_wb


# --- GUI Code ---
class SalesProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Sales Data Processor")
        master.geometry("600x450")

        self.input_files = []
        self.template_file = None
        self.output_dir = os.getcwd()

        # Styling
        label_font = ("Arial", 10)
        button_font = ("Arial", 10, "bold")
        listbox_font = ("Arial", 9)

        input_frame = tk.Frame(master, pady=10)
        input_frame.pack(fill=tk.X, padx=10)

        self.add_file_button = tk.Button(input_frame, text="Add Sales File(s)", command=self.add_files,
                                         font=button_font, bg="#4CAF50", fg="white", relief=tk.RAISED, borderwidth=2)
        self.add_file_button.pack(side=tk.LEFT, padx=5)

        self.files_listbox = tk.Listbox(master, selectmode=tk.EXTENDED, width=80, height=10, font=listbox_font)
        self.files_listbox.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        self.remove_file_button = tk.Button(master, text="Remove Selected", command=self.remove_selected_files,
                                            font=button_font, bg="#f44336", fg="white", relief=tk.RAISED, borderwidth=2)
        self.remove_file_button.pack(pady=5)

        option_frame = tk.Frame(master, pady=10)
        option_frame.pack(fill=tk.X, padx=10)

        self.select_template_button = tk.Button(option_frame, text="Select Template (Optional)",
                                                command=self.select_template, font=button_font, bg="#2196F3",
                                                fg="white", relief=tk.RAISED, borderwidth=2)
        self.select_template_button.pack(side=tk.LEFT, padx=5)
        self.template_label = tk.Label(option_frame, text="No template selected", font=label_font)
        self.template_label.pack(side=tk.LEFT, padx=5)

        self.select_output_button = tk.Button(option_frame, text="Select Output Directory",
                                              command=self.select_output_dir, font=button_font, bg="#FF9800",
                                              fg="white", relief=tk.RAISED, borderwidth=2)
        self.select_output_button.pack(side=tk.LEFT, padx=5, pady=(0, 5))
        self.output_dir_label = tk.Label(option_frame, text=f"Output: {os.path.basename(self.output_dir)}",
                                         font=label_font, wraplength=200)
        self.output_dir_label.pack(side=tk.LEFT, padx=5, pady=(0, 5))

        self.process_button = tk.Button(master, text="Process Data", command=self.process_data, font=button_font,
                                        bg="#008CBA", fg="white", height=2, relief=tk.RAISED, borderwidth=2)
        self.process_button.pack(pady=20, fill=tk.X, padx=10)

        self.status_label = tk.Label(master, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 8))
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def add_files(self):
        filepaths = filedialog.askopenfilenames(
            title="Select Sales Files",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filepaths:
            for fp in filepaths:
                base = os.path.basename(fp)
                # Improved branch key extraction: look for common separators or use filename before first dot
                branch_key_parts = base.split('_')
                if len(branch_key_parts) > 1:
                    branch_key = branch_key_parts[0]
                else:
                    branch_key_parts = base.split('-')
                    if len(branch_key_parts) > 1:
                        branch_key = branch_key_parts[0]
                    else:
                        branch_key = base.split('.')[0]

                if not any(f[0] == fp for f in self.input_files):
                    self.input_files.append((fp, branch_key))
                    self.files_listbox.insert(tk.END, f"{os.path.basename(fp)} (Branch: {branch_key})")
                    self.status_label.config(text=f"Added: {os.path.basename(fp)}")
                else:
                    messagebox.showinfo("File Exists", f"{os.path.basename(fp)} is already in the list.")

    def remove_selected_files(self):
        selected_indices = self.files_listbox.curselection()
        for i in sorted(selected_indices, reverse=True):
            self.files_listbox.delete(i)
            del self.input_files[i]
        self.status_label.config(text="Selected files removed.")

    def select_template(self):
        filepath = filedialog.askopenfilename(
            title="Select Template Excel File",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if filepath:
            self.template_file = filepath
            self.template_label.config(text=os.path.basename(filepath))
            self.status_label.config(text=f"Template selected: {os.path.basename(filepath)}")

    def select_output_dir(self):
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir = directory
            self.output_dir_label.config(text=f"Output: {os.path.basename(self.output_dir)}")
            self.status_label.config(text=f"Output directory set to: {self.output_dir}")

    def process_data(self):
        if not self.input_files:
            messagebox.showwarning("No Input Files", "Please add sales files to process.")
            return

        self.status_label.config(text="Processing... Please wait.")
        self.master.update_idletasks()

        try:
            existing_wb = None
            if self.template_file:
                try:
                    existing_wb = load_workbook(self.template_file)
                    logging.info(f"Loaded template workbook: {self.template_file}")
                except Exception as e:
                    logging.error(f"Failed to load template workbook: {e}")
                    messagebox.showerror("Template Error", f"Could not load template: {e}")
                    self.status_label.config(text="Error loading template.")
                    return

            result_wb = process_excel_data(self.input_files, template_file=None, existing_wb=existing_wb)

            if result_wb:
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                output_filename = os.path.join(self.output_dir, f"Processed_Sales_Data_{timestamp}.xlsx")

                result_wb.save(output_filename)
                logging.info(f"Successfully processed data and saved to {output_filename}")
                messagebox.showinfo("Processing Complete",
                                    f"Data processed successfully!\nOutput saved to: {output_filename}")
                self.status_label.config(text="Processing complete!")
            else:
                logging.error("Processing returned no workbook. Check logs.")
                messagebox.showerror("Processing Error",
                                     "An error occurred during processing. Workbook was not generated. Please check logs.")
                self.status_label.config(text="Processing error.")

        except Exception as e:
            logging.error(f"An error occurred during data processing: {e}", exc_info=True)
            messagebox.showerror("Processing Error", f"An unexpected error occurred: {e}\nCheck logs for details.")
            self.status_label.config(text="An error occurred.")
        finally:
            self.master.update_idletasks()


if __name__ == '__main__':
    root = tk.Tk()
    app = SalesProcessorApp(root)
    root.mainloop()
