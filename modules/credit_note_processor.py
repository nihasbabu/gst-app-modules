import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime
import os
import logging
import math

# Set up logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Sheet titles with desired order for CREDIT processing
SECTION_TITLES = [
    ("CREDIT-Total", "Credit Register - Total"),
    ("CREDIT-Total_sws", "Credit Register - Total - Receiver wise"),
    ("CREDIT-R", "Credit Register - Registered (GSTIN Available)"),
    ("CREDIT-UR", "Credit Register - Unregistered (GSTIN Not Available)"),
    ("CREDIT-Summary-Total", "Credit Register - Total - Summary"),
    ("CREDIT-Summary-R", "Credit Register - Registered - Summary"),
    ("CREDIT-Summary-UR", "Credit Register - Unregistered - Summary"),
]

# Consistent naming with Debit Note Processor for these global lists
CREDIT_FIXED_HEADERS_LIST = [
    'GSTIN/UIN of Recipient',
    'Receiver Name',
    'Branch',
    'Note number',
    'Note date',
    'Note ref no',
    'Note ref date',
    'Note value',  # Changed from Invoice value
    'Taxable Value',
    'Integrated Tax',
    'Central Tax',
    'State/UT Tax',
    'Cess',
    'Round Off'
]

CREDIT_EXCLUDE_FROM_TOTAL_HEADERS = [
    'GSTIN/UIN of Recipient',
    'Receiver Name',
    'Branch',
    'Note number',
    'Note date',
    'Note ref no',
    'Note ref date',
]


def find_header_row(worksheet):
    logging.debug(f"Searching for header row in sheet: {worksheet.title}")
    for row_idx, row in enumerate(worksheet.iter_rows()):
        if row[0].value and "Date" in str(row[0].value):
            logging.debug(f"Header row potentially found at Excel row: {row[0].row}")
            return row[0].row
    logging.debug("Header row not found by 'Date' in first column.")
    return None


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
    ws.freeze_panes = "B3"
    logging.info(f"Finished creating sheet: {sheet_name}")
    return ws


def apply_format_and_autofit(ws, columns, start_row=3, col_format_map=None):
    logging.debug(f"Applying format and autofit for sheet: {ws.title}")
    for col_idx, col_name in enumerate(columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(col_name))
        if col_format_map and col_name in col_format_map:
            for row_num in range(start_row, ws.max_row + 1):
                cell = ws.cell(row=row_num, column=col_idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = col_format_map[col_name]
        for row_num in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_num, column=col_idx).value
            if cell_value is not None:
                max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[col_letter].width = max(15, max_len + 2)
    logging.debug(f"Finished applying format and autofit for sheet: {ws.title}")


def add_total_row(ws, columns, start_row, end_row):
    logging.debug(f"Adding total row for sheet: {ws.title}")
    total_row_data = ['Total'] + [''] * (len(columns) - 1)
    total_number_format = r"[>=10000000]##\,##\,##\,##0.00;[>=100000]##\,##\,##0.00;##,##0.00;-;"

    for col_idx in range(1, len(columns) + 1):
        if col_idx == 1:
            continue
        header_name = columns[col_idx - 1]
        if header_name in CREDIT_EXCLUDE_FROM_TOTAL_HEADERS:
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
    ws.append(total_row_data)
    for col_idx, value in enumerate(total_row_data, start=1):
        cell = ws.cell(row=row_num, column=col_idx)
        cell.font = Font(bold=True, color="FF0000")
        if isinstance(value, (int, float)):
            cell.number_format = total_number_format
    logging.debug(f"Finished adding total row for sheet: {ws.title}")


def safe_float_conversion(value, header, cell=None, fixed_headers_ref=None, note_type="credit"):
    if fixed_headers_ref is None:
        fixed_headers_ref = CREDIT_FIXED_HEADERS_LIST

    if value is None:
        return 0.0

    always_sign_specific_headers = [
        'Note value',  # Changed from Invoice value
        'Taxable Value', 'Integrated Tax',
        'Central Tax', 'State/UT Tax', 'Cess'
    ]
    non_numeric_special_treatment_headers = [
        'GSTIN/UIN of Recipient', 'Receiver Name', 'Branch',
        'Note number', 'Note date', 'Note ref no', 'Note ref date'
    ]

    try:
        numeric_value = float(value)
    except (ValueError, TypeError):
        if isinstance(value, str):
            value_str = value.strip()
            if header in non_numeric_special_treatment_headers:
                return value_str
            is_cr_dr_sensitive = header == 'Round Off' or (
                        header not in fixed_headers_ref and header not in always_sign_specific_headers)
            if is_cr_dr_sensitive:
                if value_str.endswith(' Cr'):
                    try:
                        numeric_value = float(value_str[:-3])
                    except ValueError:
                        return value_str
                elif value_str.endswith(' Dr'):
                    try:
                        numeric_value = -float(value_str[:-3])
                    except ValueError:
                        return value_str
                else:
                    return value_str
            else:
                return value_str
        else:
            if header in non_numeric_special_treatment_headers:
                return str(value)
            return 0.0

    if header in always_sign_specific_headers:
        if note_type == "credit":
            return -abs(numeric_value)
        else:
            return abs(numeric_value)

    is_cr_dr_sensitive_header = header == 'Round Off' or (
                header not in fixed_headers_ref and header not in always_sign_specific_headers)
    if is_cr_dr_sensitive_header:
        if cell and cell.number_format and isinstance(numeric_value, (int, float)):
            original_value_str = str(value).strip()
            if 'Dr' in str(cell.number_format) and numeric_value > 0 and not original_value_str.endswith(' Dr'):
                numeric_value = -abs(numeric_value)
    return numeric_value


def process_credit_data(input_files, template_file=None, existing_wb=None):
    logging.info("Starting credit data processing")
    all_credit_data = []

    column_mapping = {
        'GSTIN/UIN': 'GSTIN/UIN of Recipient',
        'Particulars': 'Receiver Name',
        'Voucher No': 'Note number',
        'Voucher No.': 'Note number',
        'Date': 'Note date',
        'Gross Total': 'Note value',  # <<<< UPDATED HERE
        'Value': 'Taxable Value',
        'IGST': 'Integrated Tax',
        'CGST': 'Central Tax',
        'SGST': 'State/UT Tax',
        'Cess': 'Cess',
        'ROUND OFF': 'Round Off',
        'Round Off': 'Round Off',
        'Voucher Ref. No.': 'Note ref no',
        'Voucher Ref. Date': 'Note ref date'
    }

    fixed_headers = CREDIT_FIXED_HEADERS_LIST
    extra_headers_set = set()

    logging.debug("Starting file processing loop for credit data")
    for filepath, branch_key in input_files:
        logging.info(f"Processing credit file: {filepath}")
        try:
            wb = load_workbook(filepath, data_only=True)
            sheet_to_process_name = "Sales Register"
            ws = None
            if len(wb.sheetnames) > 1:
                if sheet_to_process_name in wb.sheetnames:
                    ws = wb[sheet_to_process_name]
                else:
                    logging.warning(
                        f"'{sheet_to_process_name}' sheet not found in {filepath}. Trying 'register' sheets.")
                    possible_sheets = [s_name for s_name in wb.sheetnames if "register" in s_name.lower()]
                    if possible_sheets:
                        ws = wb[possible_sheets[0]]
                        logging.info(f"Using first sheet with 'register' in name: {ws.title}")
                    else:
                        ws = wb.active
                        logging.warning(f"No 'register' sheet found. Falling back to active sheet: {ws.title}")
            else:
                ws = wb.active

            if not ws:
                logging.error(f"No active/suitable sheet in {filepath}")
                raise ValueError(f"No active or suitable sheet found in {filepath}")
            logging.info(f"Processing sheet: {ws.title} from file {filepath}")

            header_row_num = find_header_row(ws)
            if not header_row_num:
                logging.error(f"Header row (source 'Date') not found in {filepath} (sheet: {ws.title})")
                raise ValueError(f"Could not find header row with 'Date' in {filepath} (sheet: {ws.title})")

            headers_from_sheet = [cell.value for cell in ws[header_row_num] if cell.value is not None]
            logging.debug(f"Headers from sheet '{filepath}': {headers_from_sheet}")

            current_file_extra_headers = [
                h for h in headers_from_sheet
                if h not in column_mapping.keys() and h not in column_mapping.values()
                   and h not in fixed_headers
                   and h not in ['Others-Cr', 'Others-Dr']
            ]
            extra_headers_set.update(current_file_extra_headers)

            for row_idx, row_cells_tuple in enumerate(ws.iter_rows(min_row=header_row_num + 1, values_only=False),
                                                      start=header_row_num + 1):
                row_data_orig_values = {}
                current_row_cells_map = {}
                for i, cell_obj in enumerate(row_cells_tuple):
                    if i < len(headers_from_sheet):
                        header_val_from_sheet = headers_from_sheet[i]
                        row_data_orig_values[header_val_from_sheet] = cell_obj.value
                        current_row_cells_map[header_val_from_sheet] = cell_obj

                raw_note_date_value = row_data_orig_values.get('Date')
                if not raw_note_date_value or 'Grand Total' in str(row_data_orig_values.get('Particulars', '')):
                    continue

                processed_row_data = {}
                try:
                    if isinstance(raw_note_date_value, str):
                        try:
                            processed_row_data['Note date'] = datetime.datetime.strptime(raw_note_date_value,
                                                                                         '%Y-%m-%d %H:%M:%S')
                        except ValueError:
                            processed_row_data['Note date'] = datetime.datetime.strptime(raw_note_date_value.split()[0],
                                                                                         '%d-%m-%Y')
                    elif isinstance(raw_note_date_value, datetime.datetime):
                        processed_row_data['Note date'] = raw_note_date_value
                    elif isinstance(raw_note_date_value, (int, float)):
                        processed_row_data['Note date'] = datetime.datetime.fromordinal(
                            datetime.datetime(1900, 1, 1).toordinal() + int(raw_note_date_value) - 2)
                    else:
                        raise ValueError("Invalid date type for Note date")
                except Exception as e:
                    logging.debug(
                        f"Skipping row {row_idx} due to 'Note date' parsing error: {raw_note_date_value}, Error: {e}")
                    continue

                raw_note_ref_date_value = row_data_orig_values.get('Voucher Ref. Date')
                if raw_note_ref_date_value:
                    try:
                        if isinstance(raw_note_ref_date_value, str):
                            try:
                                processed_row_data['Note ref date'] = datetime.datetime.strptime(
                                    raw_note_ref_date_value, '%Y-%m-%d %H:%M:%S')
                            except ValueError:
                                processed_row_data['Note ref date'] = datetime.datetime.strptime(
                                    raw_note_ref_date_value.split()[0], '%d-%m-%Y')
                        elif isinstance(raw_note_ref_date_value, datetime.datetime):
                            processed_row_data['Note ref date'] = raw_note_ref_date_value
                        elif isinstance(raw_note_ref_date_value, (int, float)):
                            processed_row_data['Note ref date'] = datetime.datetime.fromordinal(
                                datetime.datetime(1900, 1, 1).toordinal() + int(raw_note_ref_date_value) - 2)
                        else:
                            processed_row_data['Note ref date'] = str(raw_note_ref_date_value)
                    except Exception as e:
                        logging.debug(
                            f"Warning: Could not parse 'Note ref date' for row {row_idx}: {raw_note_ref_date_value}, Error: {e}. Storing as string.")
                        processed_row_data['Note ref date'] = str(raw_note_ref_date_value)
                else:
                    processed_row_data['Note ref date'] = ''

                for original_header, original_value in row_data_orig_values.items():
                    mapped_header = column_mapping.get(original_header, original_header)
                    cell_for_conversion = current_row_cells_map.get(original_header)

                    if original_header == 'Date' or original_header == 'Voucher Ref. Date':
                        continue

                    if mapped_header in fixed_headers or mapped_header in current_file_extra_headers:
                        processed_row_data[mapped_header] = safe_float_conversion(original_value, mapped_header,
                                                                                  cell_for_conversion,
                                                                                  fixed_headers_ref=fixed_headers,
                                                                                  note_type="credit")
                    elif original_header == 'Others-Dr':
                        processed_row_data[original_header] = safe_float_conversion(original_value, original_header,
                                                                                    cell_for_conversion,
                                                                                    fixed_headers_ref=fixed_headers,
                                                                                    note_type="credit")
                    elif original_header == 'Others-Cr':
                        processed_row_data[original_header] = safe_float_conversion(original_value, original_header,
                                                                                    cell_for_conversion,
                                                                                    fixed_headers_ref=fixed_headers,
                                                                                    note_type="credit")

                for header in fixed_headers:
                    if header not in processed_row_data:
                        processed_row_data[header] = ''
                for header in current_file_extra_headers:
                    if header not in processed_row_data:
                        processed_row_data[header] = ''

                others_cr_val = row_data_orig_values.get('Others-Cr')
                others_dr_val = row_data_orig_values.get('Others-Dr')
                if others_cr_val is not None and 'Others-Cr' not in current_file_extra_headers and 'Others-Cr' not in fixed_headers:
                    processed_row_data['Others-Cr'] = safe_float_conversion(others_cr_val, 'Others-Cr',
                                                                            current_row_cells_map.get('Others-Cr'),
                                                                            fixed_headers_ref=fixed_headers,
                                                                            note_type="credit")
                if others_dr_val is not None and 'Others-Dr' not in current_file_extra_headers and 'Others-Dr' not in fixed_headers:
                    processed_row_data['Others-Dr'] = safe_float_conversion(others_dr_val, 'Others-Dr',
                                                                            current_row_cells_map.get('Others-Dr'),
                                                                            fixed_headers_ref=fixed_headers,
                                                                            note_type="credit")

                processed_row_data['Branch'] = branch_key
                all_credit_data.append(processed_row_data)

        except Exception as e:
            logging.error(f"Error processing file {filepath}: {e}", exc_info=True)
            messagebox.showerror("File Processing Error", f"Error processing file {os.path.basename(filepath)}: {e}")
        logging.info(f"Finished processing credit file: {filepath}")

    logging.debug("Finished file processing loop")

    extra_headers_list = sorted(list(extra_headers_set))
    temp_extra_check = set()
    for row in all_credit_data:
        for key in row.keys():
            if key not in fixed_headers and key not in extra_headers_list:
                temp_extra_check.add(key)
    newly_discovered_extras = [eh for eh in temp_extra_check if eh not in extra_headers_list]
    if newly_discovered_extras:
        extra_headers_list.extend(newly_discovered_extras)
        extra_headers_list.sort()

    final_headers = fixed_headers + extra_headers_list
    logging.info(f"Final headers for detail sheets: {final_headers}")

    data_registered = []
    data_unregistered = []
    for row_data in all_credit_data:
        gstin_recipient = str(row_data.get('GSTIN/UIN of Recipient', '')).strip()
        if gstin_recipient:
            data_registered.append(row_data)
        else:
            data_unregistered.append(row_data)
    logging.debug(
        f"Categorization: R={len(data_registered)}, UR={len(data_unregistered)}, Total={len(all_credit_data)}")

    sort_lambda = lambda x: (x.get('Note date', datetime.datetime.min) if isinstance(x.get('Note date'),
                                                                                     datetime.datetime) else datetime.datetime.min,
                             str(x.get('Note number', '')))
    all_credit_data.sort(key=sort_lambda)
    data_registered.sort(key=sort_lambda)
    data_unregistered.sort(key=sort_lambda)

    def sort_key_sws(row):
        receiver = str(row.get('Receiver Name', '')).strip().lower()
        date_val = row.get('Note date', datetime.datetime.min)
        note_num_val = str(row.get('Note number', ''))
        if receiver in ['cash', '(cancelled )'] or not receiver: return (1, receiver, date_val, note_num_val)
        return (0, receiver, date_val, note_num_val)

    all_credit_data_sws = sorted(all_credit_data, key=sort_key_sws)

    if existing_wb is not None:
        output_wb = existing_wb
    elif template_file:
        output_wb = load_workbook(template_file)
    else:
        output_wb = Workbook()
        if 'Sheet' in output_wb.sheetnames and len(output_wb.sheetnames) == 1: del output_wb['Sheet']

    standard_number_format = r"[>=10000000]##\,##\,##\,##0.00;[>=100000]##\,##\,##0.00;##,##0.00;-;"
    col_format_map = {
        'Note date': 'DD-MM-YYYY',
        'Note ref date': 'DD-MM-YYYY',
        'Note value': standard_number_format,  # Changed from Invoice value
        'Taxable Value': standard_number_format,
        'Integrated Tax': standard_number_format,
        'Central Tax': standard_number_format,
        'State/UT Tax': standard_number_format,
        'Cess': standard_number_format,
        'Round Off': standard_number_format
    }
    for eh in extra_headers_list:
        if eh not in col_format_map: col_format_map[eh] = standard_number_format

    sheets_to_create = [
        ("CREDIT-Total", all_credit_data),
        ("CREDIT-Total_sws", all_credit_data_sws),
        ("CREDIT-R", data_registered),
        ("CREDIT-UR", data_unregistered),
    ]
    for sheet_name_key, data_list_for_sheet in sheets_to_create:
        if data_list_for_sheet:
            display_title = next((t for n, t in SECTION_TITLES if n == sheet_name_key), sheet_name_key)
            ws = create_or_replace_sheet(output_wb, sheet_name_key, display_title, final_headers)
            current_start_row = 3
            for row_data_item in data_list_for_sheet:
                row_values_to_append = []
                for header_name in final_headers:
                    value = row_data_item.get(header_name, '')
                    if header_name in ['Note date', 'Note ref date'] and isinstance(value, datetime.datetime):
                        value = value.strftime('%d-%m-%Y')
                    row_values_to_append.append(value)
                ws.append(row_values_to_append)
            if data_list_for_sheet: add_total_row(ws, final_headers, current_start_row, ws.max_row)
            apply_format_and_autofit(ws, final_headers, col_format_map=col_format_map, start_row=current_start_row)
        else:
            logging.info(f"No data for sheet: {sheet_name_key}, skipping.")

    summary_sheets_data_map = [
        ("CREDIT-Summary-Total", all_credit_data),
        ("CREDIT-Summary-R", data_registered),
        ("CREDIT-Summary-UR", data_unregistered),
    ]
    summary_headers = ['Month', 'No. of Records', 'Note Value', 'Taxable Value', 'Integrated Tax',
                       # Changed from Invoice Value
                       'Central Tax', 'State/UT Tax', 'Cess']
    # Update summary_col_format_map to use the new 'Note Value' header
    summary_col_format_map = {
        'Note Value': standard_number_format,  # Changed from Invoice Value
        'Taxable Value': standard_number_format,
        'Integrated Tax': standard_number_format,
        'Central Tax': standard_number_format,
        'State/UT Tax': standard_number_format,
        'Cess': standard_number_format
    }
    # Ensure all relevant headers in summary_col_format_map get the standard_number_format
    # This is a more robust way if summary_headers changes
    for header in summary_headers:
        if header not in ['Month', 'No. of Records'] and header not in summary_col_format_map:
            summary_col_format_map[header] = standard_number_format

    months_order = ['April', 'May', 'June', 'July', 'August', 'September',
                    'October', 'November', 'December', 'January', 'February', 'March']

    for sheet_name_key, data_for_summary in summary_sheets_data_map:
        if data_for_summary:
            display_title = next((t for n, t in SECTION_TITLES if n == sheet_name_key), sheet_name_key)
            ws_summary = create_or_replace_sheet(output_wb, sheet_name_key, display_title, summary_headers)
            monthly_summary_data = {}
            unique_notes_by_month = {}
            for row_item in data_for_summary:
                date_obj = row_item.get('Note date')
                note_num_val = str(row_item.get('Note number', '')).strip()
                if not isinstance(date_obj, datetime.datetime) or not note_num_val: continue
                month_name = months_order[(date_obj.month - 4 + 12) % 12]
                if month_name not in monthly_summary_data:
                    # Initialize all numeric summary headers to 0.0
                    monthly_summary_data[month_name] = {h: 0.0 for h in summary_headers if
                                                        h not in ['Month', 'No. of Records']}
                    monthly_summary_data[month_name]['count'] = 0
                    unique_notes_by_month[month_name] = set()

                if note_num_val not in unique_notes_by_month[month_name]:
                    monthly_summary_data[month_name]['count'] += 1
                    unique_notes_by_month[month_name].add(note_num_val)

                # Map detail sheet headers to summary sheet headers for summation
                fields_to_sum = {
                    'Note value': 'Note Value',  # Changed from 'Invoice value': 'Invoice Value'
                    'Taxable Value': 'Taxable Value',
                    'Integrated Tax': 'Integrated Tax',
                    'Central Tax': 'Central Tax',
                    'State/UT Tax': 'State/UT Tax',
                    'Cess': 'Cess'
                }
                for detail_key, summary_key in fields_to_sum.items():
                    value_to_add = row_item.get(detail_key, 0.0)
                    if isinstance(value_to_add, (int, float)) and not (
                            math.isnan(value_to_add) or math.isinf(value_to_add)):
                        monthly_summary_data[month_name][summary_key] += float(value_to_add)

            summary_start_row = 3;
            rows_added = 0
            for month_n in months_order:
                if month_n in monthly_summary_data:
                    data = monthly_summary_data[month_n]
                    row_to_append = [month_n, data.get('count', 0)] + [data.get(h, 0.0) for h in summary_headers[2:]]
                    ws_summary.append(row_to_append)
                    rows_added += 1

            if rows_added > 0:
                total_vals = ['Total', sum(monthly_summary_data[m].get('count', 0) for m in monthly_summary_data)]
                for h_idx in range(2, len(summary_headers)):
                    col_header = summary_headers[h_idx]
                    total_vals.append(sum(monthly_summary_data[m].get(col_header, 0.0) for m in monthly_summary_data))
                ws_summary.append(total_vals)
                total_row_num = summary_start_row + rows_added
                for c_idx, val in enumerate(total_vals, 1):
                    cell = ws_summary.cell(row=total_row_num, column=c_idx)
                    cell.font = Font(bold=True, color="FF0000")
                    if isinstance(val, (int, float)) and summary_headers[c_idx - 1] not in ['Month', 'No. of Records']:
                        cell.number_format = standard_number_format
            apply_format_and_autofit(ws_summary, summary_headers, col_format_map=summary_col_format_map,
                                     start_row=summary_start_row)
        else:
            logging.info(f"No data for summary: {sheet_name_key}, skipping.")
    logging.info("Completed credit data processing")
    return output_wb


class CreditProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Credit Note Processor")
        master.geometry("600x450")

        self.input_files = []
        self.template_file = None
        self.output_dir = os.getcwd()

        label_font = ("Arial", 10)
        button_font = ("Arial", 10, "bold")
        listbox_font = ("Arial", 9)

        input_frame = tk.Frame(master, pady=10)
        input_frame.pack(fill=tk.X, padx=10)

        self.add_file_button = tk.Button(input_frame, text="Add Credit File(s)", command=self.add_files,
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

        self.process_button = tk.Button(master, text="Process Credit Data", command=self.process_data, font=button_font,
                                        bg="#008CBA", fg="white", height=2, relief=tk.RAISED, borderwidth=2)
        self.process_button.pack(pady=20, fill=tk.X, padx=10)

        self.status_label = tk.Label(master, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 8))
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def add_files(self):
        filepaths = filedialog.askopenfilenames(
            title="Select Credit Files",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filepaths:
            for fp in filepaths:
                base = os.path.basename(fp)
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
            messagebox.showwarning("No Input Files", "Please add credit files to process.")
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

            result_wb = process_credit_data(self.input_files, template_file=None, existing_wb=existing_wb)

            if result_wb:
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                output_filename = os.path.join(self.output_dir, f"Processed_Credit_Note_Data_{timestamp}.xlsx")

                result_wb.save(output_filename)
                logging.info(f"Successfully processed credit data and saved to {output_filename}")
                messagebox.showinfo("Processing Complete",
                                    f"Credit data processed successfully!\nOutput saved to: {output_filename}")
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
    app = CreditProcessorApp(root)
    root.mainloop()
