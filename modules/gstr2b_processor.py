# gstr2b_processor.py

import json
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import os, datetime

# Assuming telemetry is a module you have for sending events.
# If not, you might need to comment out or adjust these lines.
try:
    from telemetry import send_event
except ImportError:
    print("[WARNING] telemetry module not found. send_event calls will be skipped.")


    def send_event(event_type, payload):
        print(f"[TELEMETRY_STUB] Event: {event_type}, Payload: {payload}")

# ----------------------- Global Variables for Totals ----------------------- #
INDIAN_FORMAT_GSTR2B = r"[>=10000000]##\,##\,##\,##0.00;[>=100000]##\,##\,##0.00;##,##0.00" + \
                r";-##\,##\,##\,##0.00" + \
                r";-"
RED_BOLD_FONT = Font(bold=True, color="FF0000")
BOLD_FONT = Font(bold=True)  # Standard bold font

# Configuration for summing main document values uniquely in detail sheets
GSTR2B_MAIN_VALUE_CONFIG = {
    "2B-B2B": {"value_col": "Invoice Value", "id_col": "Invoice number"},
    "2B-B2BA": {"value_col": "Invoice Value", "id_col": "Invoice number"},
    "2B-B2B(ITC_Rej)": {"value_col": "Invoice Value", "id_col": "Invoice number"},
    "2B-CDNR": {"value_col": "Note Value", "id_col": "Note number"},
    "2B-IMPG": {"value_col": "Calculated Invoice Value", "id_col": "Bill of Entry Number"},
    "2B-B2BA(cum)": {"value_col": "Calculated Invoice Value", "id_col": None}  # Direct sum for this cumulative sheet
}

# Columns to sum in the total row of detail sheets
GSTR2B_DETAIL_SHEET_TOTAL_COLUMNS = {
    "2B-B2B": ["Invoice Value", "Total Taxable Value", "Integrated Tax", "Central Tax", "State/UT Tax", "Cess"],
    "2B-B2BA": ["Invoice Value", "Total Taxable Value", "Integrated Tax", "Central Tax", "State/UT Tax", "Cess"],
    "2B-B2BA(cum)": ["Total Documents", "Calculated Invoice Value", "Total Taxable Value", "Integrated Tax",
                     "Central Tax", "State/UT Tax", "Cess"],
    "2B-CDNR": ["Note Value", "Total Taxable Value", "Integrated Tax", "Central Tax", "State/UT Tax", "Cess"],
    "2B-IMPG": ["Calculated Invoice Value", "Taxable Value", "Integrated Tax", "Cess"],
    "2B-B2B(ITC_Rej)": ["Invoice Value", "Total Taxable Value", "Integrated Tax", "Central Tax", "State/UT Tax", "Cess"]
}
# Add _sws variations to detail sheet total columns, they share the same columns for totals
for key in list(GSTR2B_DETAIL_SHEET_TOTAL_COLUMNS.keys()):  # Iterate over a copy of keys
    if not key.endswith("_sws"):
        sws_key = key + "_sws"
        GSTR2B_DETAIL_SHEET_TOTAL_COLUMNS[sws_key] = GSTR2B_DETAIL_SHEET_TOTAL_COLUMNS[key]
        if key in GSTR2B_MAIN_VALUE_CONFIG:
            GSTR2B_MAIN_VALUE_CONFIG[sws_key] = GSTR2B_MAIN_VALUE_CONFIG[key]

# Columns to sum in the total row of summary sheets
GSTR2B_SUMMARY_SHEET_TOTAL_COLUMNS = ["Invoice Value", "Taxable Value", "Integrated Tax", "Central Tax", "State/UT Tax",
                                      "Cess"]

# ----------------------- Allowed Sections ----------------------- #
ALLOWED_DOC_SECTIONS = {
    "docdata": {"b2b", "b2ba", "cdnr", "impg"},
    "docRejdata": {"cdnr", "b2b"},
    "cpsumm": {"b2ba"}
}


# ----------------------- Utility Functions ----------------------- #
def get_tax_period(ret_period):
    month_map = {
        "01": "January", "02": "February", "03": "March", "04": "April",
        "05": "May", "06": "June", "07": "July", "08": "August",
        "09": "September", "10": "October", "11": "November", "12": "December"
    }
    if ret_period and len(ret_period) >= 2:
        return month_map.get(ret_period[:2], "Unknown")
    return "Unknown"


def parse_number(val, float_2dec=False, int_no_dec=False):
    try:
        num = float(val)
        if int_no_dec:
            return int(num)
        if float_2dec:
            return round(num, 2)
        return num
    except (ValueError, TypeError):
        return 0


def get_numeric_value(item, key):
    val = item.get(key, 0)
    if isinstance(val, str):
        val = val.strip()
    return parse_number(val, float_2dec=True)


# ----------------------- Extraction Functions ----------------------- #
def extract_b2b(data, filing_period):
    """Extract B2B section data from GSTR-2B JSON (ITC Accepted)."""
    b2b_data = data.get("data", {}).get("docdata", {}).get("b2b", [])
    if not isinstance(b2b_data, list): b2b_data = []
    extracted_data = {"2B-B2B": []}
    for supplier in b2b_data:
        if not isinstance(supplier, dict): continue
        ctin = supplier.get("ctin", "")
        trdnm = supplier.get("trdnm", "")
        invoices = supplier.get("inv", [])
        if not isinstance(invoices, list): invoices = [invoices] if isinstance(invoices, dict) else []
        for inv in invoices:
            if not isinstance(inv, dict): continue
            base_row = {
                "GSTIN of supplier": ctin, "Trade/legal name": trdnm, "Invoice number": inv.get("inum", ""),
                "Invoice type": inv.get("typ", ""), "Invoice Date": inv.get("dt", ""),
                "Invoice Value": get_numeric_value(inv, "val"),
                "Place of supply": parse_number(inv.get("pos", "0"), int_no_dec=True),
                "Supply Attract Reverse Charge": inv.get("rev", ""),
                "GSTR Period": get_tax_period(supplier.get("supprd", "")),
                "GSTR Filing Date": supplier.get("supfildt", ""),
                "GSTR Filing Period": filing_period,
                "ITC Availability": inv.get("itcavl", ""), "Reason": inv.get("rsn", ""),
                "Source": inv.get("srctyp", "")
            }
            items = inv.get("items", [])
            if not isinstance(items, list): items = []
            if items:
                for item in items:
                    if not isinstance(item, dict): continue
                    row = base_row.copy()
                    row.update({
                        "Invoice Part": item.get("num", ""), "Tax Rate": get_numeric_value(item, "rt"),
                        "Total Taxable Value": get_numeric_value(item, "txval"),
                        "Integrated Tax": get_numeric_value(item, "igst"),
                        "Central Tax": get_numeric_value(item, "cgst"),
                        "State/UT Tax": get_numeric_value(item, "sgst"),
                        "Cess": get_numeric_value(item, "cess")
                    })
                    extracted_data["2B-B2B"].append(row)
            else:
                row = base_row.copy()
                row.update({
                    "Invoice Part": "", "Tax Rate": "",
                    "Total Taxable Value": get_numeric_value(inv, "txval"),
                    "Integrated Tax": get_numeric_value(inv, "igst"),
                    "Central Tax": get_numeric_value(inv, "cgst"),
                    "State/UT Tax": get_numeric_value(inv, "sgst"),
                    "Cess": get_numeric_value(inv, "cess")
                })
                extracted_data["2B-B2B"].append(row)
    return extracted_data


def extract_b2ba(data, filing_period):
    """Extract B2BA section data from GSTR-2B JSON."""
    b2ba_data = data.get("data", {}).get("docdata", {}).get("b2ba", [])
    if not isinstance(b2ba_data, list): b2ba_data = []
    extracted_data = {"2B-B2BA": []}
    for supplier in b2ba_data:
        if not isinstance(supplier, dict): continue
        ctin = supplier.get("ctin", "")
        invoices = supplier.get("inv", [])
        if not isinstance(invoices, list): invoices = [invoices] if isinstance(invoices, dict) else []
        for inv in invoices:
            if not isinstance(inv, dict): continue
            base_row = {
                "GSTIN of supplier": ctin, "Trade/legal name": supplier.get("trdnm", ""),
                "Original Invoice number": inv.get("oinum", ""), "Original Invoice Date": inv.get("oidt", ""),
                "Invoice number": inv.get("inum", ""), "Invoice Date": inv.get("dt", ""),
                "Invoice type": inv.get("typ", ""), "Invoice Value": get_numeric_value(inv, "val"),
                "Place of supply": parse_number(inv.get("pos", "0"), int_no_dec=True),
                "Supply Attract Reverse Charge": inv.get("rev", ""),
                "GSTR Period": get_tax_period(supplier.get("supprd", "")),
                "GSTR Filing Date": supplier.get("supfildt", ""),
                "GSTR Filing Period": filing_period,
                "ITC Availability": inv.get("itcavl", ""), "Reason": inv.get("rsn", "")
            }
            items = inv.get("items", [])
            if not isinstance(items, list): items = []
            if items:
                for item in items:
                    if not isinstance(item, dict): continue
                    row = base_row.copy()
                    row.update({
                        "Invoice Part": item.get("num", ""), "Tax Rate": get_numeric_value(item, "rt"),
                        "Total Taxable Value": get_numeric_value(item, "txval"),
                        "Integrated Tax": get_numeric_value(item, "igst"),
                        "Central Tax": get_numeric_value(item, "cgst"),
                        "State/UT Tax": get_numeric_value(item, "sgst"),
                        "Cess": get_numeric_value(item, "cess")
                    })
                    extracted_data["2B-B2BA"].append(row)
            else:
                row = base_row.copy()
                row.update({
                    "Invoice Part": "", "Tax Rate": "",
                    "Total Taxable Value": get_numeric_value(inv, "txval"),
                    "Integrated Tax": get_numeric_value(inv, "igst"),
                    "Central Tax": get_numeric_value(inv, "cgst"),
                    "State/UT Tax": get_numeric_value(inv, "sgst"),
                    "Cess": get_numeric_value(inv, "cess")
                })
                extracted_data["2B-B2BA"].append(row)
    return extracted_data


def extract_b2ba_cum(data, filing_period):
    """Extract B2BA cumulative section data from GSTR-2B JSON."""
    b2ba_cum_data = data.get("data", {}).get("cpsumm", {}).get("b2ba", [])
    if not isinstance(b2ba_cum_data, list): b2ba_cum_data = []
    extracted_data = {"2B-B2BA(cum)": []}
    for supplier in b2ba_cum_data:
        if not isinstance(supplier, dict): continue
        row = {
            "GSTIN of supplier": supplier.get("ctin", ""),
            "Trade/legal name": supplier.get("trdnm", ""),
            "Total Documents": parse_number(supplier.get("ttldocs", 0), int_no_dec=True),
            "Total Taxable Value": get_numeric_value(supplier, "txval"),
            "Integrated Tax": get_numeric_value(supplier, "igst"),
            "Central Tax": get_numeric_value(supplier, "cgst"),
            "State/UT Tax": get_numeric_value(supplier, "sgst"),
            "Cess": get_numeric_value(supplier, "cess"),
            "GSTR Period": get_tax_period(supplier.get("supprd", "")),
            "GSTR Filing Date": supplier.get("supfildt", ""),
            "GSTR Filing Period": filing_period
        }
        row["Calculated Invoice Value"] = sum(get_numeric_value(row, k) for k in
                                              ["Total Taxable Value", "Integrated Tax", "Central Tax", "State/UT Tax",
                                               "Cess"])
        extracted_data["2B-B2BA(cum)"].append(row)
    return extracted_data


def extract_cdnr(data, filing_period):
    """Extract CDNR section data from GSTR-2B JSON."""
    cdnr_data = data.get("data", {}).get("docdata", {}).get("cdnr", [])
    if not isinstance(cdnr_data, list): cdnr_data = []
    extracted_data = {"2B-CDNR": []}
    for supplier in cdnr_data:
        if not isinstance(supplier, dict): continue
        ctin = supplier.get("ctin", "")
        notes = supplier.get("nt", [])
        if not isinstance(notes, list): notes = [notes] if isinstance(notes, dict) else []
        for note in notes:
            if not isinstance(note, dict): continue
            base_row = {
                "GSTIN of supplier": ctin, "Trade/legal name": supplier.get("trdnm", ""),
                "Note number": note.get("ntnum", ""), "Note type": note.get("typ", ""),
                "Note supply type": note.get("suptyp", ""), "Note Date": note.get("dt", ""),
                "Note Value": get_numeric_value(note, "val"),
                "Place of supply": parse_number(note.get("pos", "0"), int_no_dec=True),
                "Supply Attract Reverse Charge": note.get("rev", ""),
                "GSTR Period": get_tax_period(supplier.get("supprd", "")),
                "GSTR Filing Date": supplier.get("supfildt", ""),
                "GSTR Filing Period": filing_period,
                "ITC Availability": note.get("itcavl", ""), "Reason": note.get("rsn", "")
            }
            items = note.get("items", [])
            if not isinstance(items, list): items = []
            if items:
                for item in items:
                    if not isinstance(item, dict): continue
                    row = base_row.copy()
                    row.update({
                        "Note Part": item.get("num", ""), "Tax Rate": get_numeric_value(item, "rt"),
                        "Total Taxable Value": get_numeric_value(item, "txval"),
                        "Integrated Tax": get_numeric_value(item, "igst"),
                        "Central Tax": get_numeric_value(item, "cgst"),
                        "State/UT Tax": get_numeric_value(item, "sgst"),
                        "Cess": get_numeric_value(item, "cess")
                    })
                    extracted_data["2B-CDNR"].append(row)
            else:
                row = base_row.copy()
                row.update({
                    "Note Part": "", "Tax Rate": "",
                    "Total Taxable Value": get_numeric_value(note, "txval"),
                    "Integrated Tax": get_numeric_value(note, "igst"),
                    "Central Tax": get_numeric_value(note, "cgst"),
                    "State/UT Tax": get_numeric_value(note, "sgst"),
                    "Cess": get_numeric_value(note, "cess")
                })
                extracted_data["2B-CDNR"].append(row)
    return extracted_data


def extract_impg(data, filing_period):
    """Extract IMPG section data from GSTR-2B JSON."""
    impg_data = data.get("data", {}).get("docdata", {}).get("impg", [])
    if not isinstance(impg_data, list): impg_data = []
    extracted_data = {"2B-IMPG": []}
    for entry in impg_data:
        if not isinstance(entry, dict): continue
        row = {
            "ICEGATE Reference Date": entry.get("refdt", ""), "Port Code": entry.get("portcode", ""),
            "Bill of Entry Number": entry.get("boenum", ""), "Bill of Entry Date": entry.get("boedt", ""),
            "Taxable Value": get_numeric_value(entry, "txval"),
            "Integrated Tax": get_numeric_value(entry, "igst"), "Cess": get_numeric_value(entry, "cess"),
            "Record Date": entry.get("recdt", ""), "GSTR Filing Period": filing_period,
            "Amended (Yes)": entry.get("isamd", "")
        }
        row["Calculated Invoice Value"] = sum(
            get_numeric_value(row, k) for k in ["Taxable Value", "Integrated Tax", "Cess"])
        extracted_data["2B-IMPG"].append(row)
    return extracted_data


def extract_b2b_itc_rej(data, filing_period):
    """Extract B2B (ITC Rejected) section data from GSTR-2B JSON."""
    b2b_rej_data = data.get("data", {}).get("docRejdata", {}).get("b2b", [])
    if not isinstance(b2b_rej_data, list): b2b_rej_data = []
    extracted_data = {"2B-B2B(ITC_Rej)": []}
    for supplier in b2b_rej_data:
        if not isinstance(supplier, dict): continue
        ctin = supplier.get("ctin", "")
        invoices = supplier.get("inv", [])
        if not isinstance(invoices, list): invoices = [invoices] if isinstance(invoices, dict) else []
        for inv in invoices:
            if not isinstance(inv, dict): continue
            base_row = {
                "GSTIN of supplier": ctin, "Trade/legal name": supplier.get("trdnm", ""),
                "Invoice number": inv.get("inum", ""), "Invoice type": inv.get("typ", ""),
                "Invoice Date": inv.get("dt", ""), "Invoice Value": get_numeric_value(inv, "val"),
                "Place of supply": parse_number(inv.get("pos", "0"), int_no_dec=True),
                "GSTR Period": get_tax_period(supplier.get("supprd", "")),
                "GSTR Filing Date": supplier.get("supfildt", ""),
                "GSTR Filing Period": filing_period, "Source": inv.get("srctyp", "")
            }
            items = inv.get("items", [])
            if not isinstance(items, list): items = []
            if items:
                for item in items:
                    if not isinstance(item, dict): continue
                    row = base_row.copy()
                    row.update({
                        "Invoice Part": item.get("num", ""), "Tax Rate": get_numeric_value(item, "rt"),
                        "Total Taxable Value": get_numeric_value(item, "txval"),
                        "Integrated Tax": get_numeric_value(item, "igst"),
                        "Central Tax": get_numeric_value(item, "cgst"),
                        "State/UT Tax": get_numeric_value(item, "sgst"),
                        "Cess": get_numeric_value(item, "cess")
                    })
                    extracted_data["2B-B2B(ITC_Rej)"].append(row)
            else:
                row = base_row.copy()
                row.update({
                    "Invoice Part": "", "Tax Rate": "",
                    "Total Taxable Value": get_numeric_value(inv, "txval"),
                    "Integrated Tax": get_numeric_value(inv, "igst"),
                    "Central Tax": get_numeric_value(inv, "cgst"),
                    "State/UT Tax": get_numeric_value(inv, "sgst"),
                    "Cess": get_numeric_value(inv, "cess")
                })
                extracted_data["2B-B2B(ITC_Rej)"].append(row)
    return extracted_data


# ----------------------- New Helper Function for Detail Sheet Totals ----------------------- #
def _add_total_row_to_gstr2b_detail_sheet(ws, sheet_key, rows_data, column_headers_for_sheet, column_formats_for_sheet):
    """Adds a formatted total row to a detail worksheet."""
    if not rows_data:
        return

    total_row_idx = ws.max_row + 1
    totals = {}
    columns_to_sum_for_current_sheet = GSTR2B_DETAIL_SHEET_TOTAL_COLUMNS.get(sheet_key, [])
    main_value_conf = GSTR2B_MAIN_VALUE_CONFIG.get(sheet_key, {})
    main_value_col = main_value_conf.get("value_col")
    id_col = main_value_conf.get("id_col")

    if main_value_col and main_value_col in columns_to_sum_for_current_sheet:
        if id_col:
            summed_ids_for_main_val = set()
            current_sum_main_val = 0
            for row_item in rows_data:
                doc_id = row_item.get(id_col)
                if doc_id and (doc_id, main_value_col) not in summed_ids_for_main_val:
                    current_sum_main_val += parse_number(row_item.get(main_value_col, 0), float_2dec=True)
                    summed_ids_for_main_val.add((doc_id, main_value_col))
                elif not doc_id:
                    current_sum_main_val += parse_number(row_item.get(main_value_col, 0), float_2dec=True)
            totals[main_value_col] = current_sum_main_val
        else:
            totals[main_value_col] = sum(parse_number(r.get(main_value_col, 0), float_2dec=True) for r in rows_data)

    for col_header in columns_to_sum_for_current_sheet:
        if col_header == main_value_col:
            continue
        totals[col_header] = sum(parse_number(r.get(col_header, 0), float_2dec=True) for r in rows_data)

    label_cell = ws.cell(row=total_row_idx, column=1)
    label_cell.value = "Total"
    label_cell.font = RED_BOLD_FONT
    label_cell.alignment = Alignment(horizontal="left")

    for col_idx, header_name in enumerate(column_headers_for_sheet, start=1):
        if col_idx == 1:  # First column is now strictly for "Total" label in the total row
            continue

        cell = ws.cell(row=total_row_idx, column=col_idx)
        if header_name in totals:
            cell.value = totals[header_name]
            cell.font = RED_BOLD_FONT
            fmt = column_formats_for_sheet.get(header_name, INDIAN_FORMAT_GSTR2B)
            cell.number_format = fmt
        else:
            cell.value = ""  # Blank for non-totaled columns


# ----------------------- New Helper Function for Summary Sheet Totals ----------------------- #
def _add_total_row_to_gstr2b_summary_sheet(ws, summary_sheet_data, display_column_headers, summary_column_formats_map):
    """Adds a formatted total row to a summary worksheet."""
    if not summary_sheet_data:
        return

    total_row_idx = ws.max_row + 1
    grand_totals = {}

    for col_header in GSTR2B_SUMMARY_SHEET_TOTAL_COLUMNS:
        grand_totals[col_header] = sum(
            parse_number(month_data.get(col_header, 0), float_2dec=True) for month_data in summary_sheet_data)

    for col_idx, header_name in enumerate(display_column_headers, start=1):
        cell = ws.cell(row=total_row_idx, column=col_idx)
        if header_name == "Month":
            cell.value = "Total"  # Changed from "Grand Total"
            cell.font = RED_BOLD_FONT
            cell.alignment = Alignment(horizontal="left")
        elif header_name in grand_totals:
            total_val = grand_totals.get(header_name, 0)
            cell.value = total_val
            cell.font = RED_BOLD_FONT
            fmt = summary_column_formats_map.get(header_name, INDIAN_FORMAT_GSTR2B)
            cell.number_format = fmt
        else:
            cell.value = ""


# ----------------------- Summary Generation (Modified) ----------------------- #
def create_summary_sheets(wb, combined_data):
    print("[DEBUG] Creating summary sheets...")
    summary_data_output = {
        "2B-Summary-B2B_not RC(ITC_Avl)": [], "2B-Summary-B2B_RC(ITC_Avl)": [],
        "2B-Summary-B2BA_not RC(ITC_Avl)": [], "2B-Summary-B2BA_RC(ITC_Avl)": [],
        "2B-Summary-B2BA_cum(ITC_Avl)": [],
        "2B-Summary-CDNR_DN(ITC_Avl)": [], "2B-Summary-CDNR_CN(ITC_Avl)": [], "2B-Summary-CDNR_RC(ITC_Avl)": [],
        "2B-Summary-IMPG(ITC_Avl)": [], "2B-Summary-IMPGA(ITC_Avl)": [],
        "2B-Summary-B2B(ITC_Rej)": []
    }
    financial_order = ["April", "May", "June", "July", "August", "September",
                       "October", "November", "December", "January", "February", "March"]

    all_filing_periods = set()
    for _, rows in combined_data.items():
        for row in rows:
            filing_period = row.get("GSTR Filing Period")
            if filing_period and isinstance(filing_period, str):
                all_filing_periods.add(filing_period)

    all_filing_periods_sorted = sorted(list(all_filing_periods),
                                       key=lambda x: financial_order.index(x) if x in financial_order else 999)

    monthly_template = {col: 0 for col in GSTR2B_SUMMARY_SHEET_TOTAL_COLUMNS}
    agg_summaries = {key: {month: {"Month": month, **monthly_template.copy()} for month in all_filing_periods_sorted}
                     for key in summary_data_output.keys()}

    b2b_rows = combined_data.get("2B-B2B", [])
    unique_inv_b2b_not_rc = {}
    unique_inv_b2b_rc = {}
    for r in b2b_rows:
        fp = r.get("GSTR Filing Period")
        if not fp or fp not in all_filing_periods_sorted: continue
        inv_val = get_numeric_value(r, "Invoice Value")
        inv_num = r.get("Invoice number")
        target_summary_key = None
        unique_set_for_inv = None

        if r.get("Supply Attract Reverse Charge") == "N":
            target_summary_key = "2B-Summary-B2B_not RC(ITC_Avl)"
            unique_set_for_inv = unique_inv_b2b_not_rc
        elif r.get("Supply Attract Reverse Charge") == "Y":
            target_summary_key = "2B-Summary-B2B_RC(ITC_Avl)"
            unique_set_for_inv = unique_inv_b2b_rc

        if not target_summary_key: continue

        target = agg_summaries[target_summary_key][fp]
        if inv_num not in unique_set_for_inv.setdefault(fp, {}):
            target["Invoice Value"] += inv_val
            unique_set_for_inv[fp][inv_num] = inv_val

        target["Taxable Value"] += get_numeric_value(r, "Total Taxable Value")
        target["Integrated Tax"] += get_numeric_value(r, "Integrated Tax")
        target["Central Tax"] += get_numeric_value(r, "Central Tax")
        target["State/UT Tax"] += get_numeric_value(r, "State/UT Tax")
        target["Cess"] += get_numeric_value(r, "Cess")

    b2ba_rows = combined_data.get("2B-B2BA", [])
    unique_inv_b2ba_not_rc = {}
    unique_inv_b2ba_rc = {}
    for r in b2ba_rows:
        fp = r.get("GSTR Filing Period")
        if not fp or fp not in all_filing_periods_sorted: continue
        inv_val = get_numeric_value(r, "Invoice Value")
        inv_num = r.get("Invoice number")
        target_summary_key = None
        unique_set_for_inv = None

        if r.get("Supply Attract Reverse Charge") == "N":
            target_summary_key = "2B-Summary-B2BA_not RC(ITC_Avl)"
            unique_set_for_inv = unique_inv_b2ba_not_rc
        elif r.get("Supply Attract Reverse Charge") == "Y":
            target_summary_key = "2B-Summary-B2BA_RC(ITC_Avl)"
            unique_set_for_inv = unique_inv_b2ba_rc

        if not target_summary_key: continue

        target = agg_summaries[target_summary_key][fp]
        if inv_num not in unique_set_for_inv.setdefault(fp, {}):
            target["Invoice Value"] += inv_val
            unique_set_for_inv[fp][inv_num] = inv_val

        target["Taxable Value"] += get_numeric_value(r, "Total Taxable Value")
        target["Integrated Tax"] += get_numeric_value(r, "Integrated Tax")
        target["Central Tax"] += get_numeric_value(r, "Central Tax")
        target["State/UT Tax"] += get_numeric_value(r, "State/UT Tax")
        target["Cess"] += get_numeric_value(r, "Cess")

    b2bacum_rows = combined_data.get("2B-B2BA(cum)", [])
    for r in b2bacum_rows:
        fp = r.get("GSTR Filing Period")
        if not fp or fp not in all_filing_periods_sorted: continue
        target = agg_summaries["2B-Summary-B2BA_cum(ITC_Avl)"][fp]
        target["Invoice Value"] += get_numeric_value(r, "Calculated Invoice Value")
        target["Taxable Value"] += get_numeric_value(r, "Total Taxable Value")
        target["Integrated Tax"] += get_numeric_value(r, "Integrated Tax")
        target["Central Tax"] += get_numeric_value(r, "Central Tax")
        target["State/UT Tax"] += get_numeric_value(r, "State/UT Tax")
        target["Cess"] += get_numeric_value(r, "Cess")

    cdnr_rows = combined_data.get("2B-CDNR", [])
    unique_notes_cn = {}
    unique_notes_dn = {}
    unique_notes_rc = {}
    for r in cdnr_rows:
        fp = r.get("GSTR Filing Period")
        if not fp or fp not in all_filing_periods_sorted: continue
        note_val = get_numeric_value(r, "Note Value")
        note_num = r.get("Note number")
        note_type = r.get("Note type", "").upper()
        target_key = None
        unique_set_for_note = None

        if r.get("Supply Attract Reverse Charge") == "Y":
            target_key = "2B-Summary-CDNR_RC(ITC_Avl)"
            unique_set_for_note = unique_notes_rc
        elif note_type == "C":
            target_key = "2B-Summary-CDNR_CN(ITC_Avl)"
            unique_set_for_note = unique_notes_cn
        elif note_type == "D":
            target_key = "2B-Summary-CDNR_DN(ITC_Avl)"
            unique_set_for_note = unique_notes_dn

        if not target_key: continue

        target = agg_summaries[target_key][fp]
        if note_num not in unique_set_for_note.setdefault(fp, {}):
            target["Invoice Value"] += note_val
            unique_set_for_note[fp][note_num] = note_val
        target["Taxable Value"] += get_numeric_value(r, "Total Taxable Value")
        target["Integrated Tax"] += get_numeric_value(r, "Integrated Tax")
        target["Central Tax"] += get_numeric_value(r, "Central Tax")
        target["State/UT Tax"] += get_numeric_value(r, "State/UT Tax")
        target["Cess"] += get_numeric_value(r, "Cess")

    impg_rows = combined_data.get("2B-IMPG", [])
    unique_boe_impg = {}
    unique_boe_impga = {}
    for r in impg_rows:
        fp = r.get("GSTR Filing Period")
        if not fp or fp not in all_filing_periods_sorted: continue
        boe_val = get_numeric_value(r, "Calculated Invoice Value")
        boe_num = r.get("Bill of Entry Number")
        is_amended = r.get("Amended (Yes)", "N").upper() == "Y"
        target_key = None
        unique_set_for_boe = None

        if is_amended:
            target_key = "2B-Summary-IMPGA(ITC_Avl)"
            unique_set_for_boe = unique_boe_impga
        else:
            target_key = "2B-Summary-IMPG(ITC_Avl)"
            unique_set_for_boe = unique_boe_impg

        target = agg_summaries[target_key][fp]
        if boe_num not in unique_set_for_boe.setdefault(fp, {}):
            target["Invoice Value"] += boe_val
            unique_set_for_boe[fp][boe_num] = boe_val

        target["Taxable Value"] += get_numeric_value(r, "Taxable Value")
        target["Integrated Tax"] += get_numeric_value(r, "Integrated Tax")
        target["Cess"] += get_numeric_value(r, "Cess")

    b2b_rej_rows = combined_data.get("2B-B2B(ITC_Rej)", [])
    unique_inv_b2b_rej = {}
    for r in b2b_rej_rows:
        fp = r.get("GSTR Filing Period")
        if not fp or fp not in all_filing_periods_sorted: continue
        target = agg_summaries["2B-Summary-B2B(ITC_Rej)"][fp]
        inv_val = get_numeric_value(r, "Invoice Value")
        inv_num = r.get("Invoice number")
        if inv_num not in unique_inv_b2b_rej.setdefault(fp, {}):
            target["Invoice Value"] += inv_val
            unique_inv_b2b_rej[fp][inv_num] = inv_val

        target["Taxable Value"] += get_numeric_value(r, "Total Taxable Value")
        target["Integrated Tax"] += get_numeric_value(r, "Integrated Tax")
        target["Central Tax"] += get_numeric_value(r, "Central Tax")
        target["State/UT Tax"] += get_numeric_value(r, "State/UT Tax")
        target["Cess"] += get_numeric_value(r, "Cess")

    for key, monthly_data_dict in agg_summaries.items():
        summary_data_output[key] = list(monthly_data_dict.values())

    section_titles_summary = {
        "2B-Summary-B2B_not RC(ITC_Avl)": "3. ITC Available - PART A1 - Supplies other than Reverse charge - B2B Invoices (IMS) - Summary",
        "2B-Summary-B2B_RC(ITC_Avl)": "3. ITC Available - PART A3 - Supplies liable for Reverse charge - B2B Invoices - Summary",
        "2B-Summary-B2BA_not RC(ITC_Avl)": "3. ITC Available - PART A1 - Supplies other than Reverse charge - B2BA Invoices (IMS) - Summary",
        "2B-Summary-B2BA_RC(ITC_Avl)": "3. ITC Available - PART A3 - Supplies liable for Reverse charge - B2BA Invoices (IMS) - Summary",
        "2B-Summary-B2BA_cum(ITC_Avl)": "3. ITC Available - PART A1&A3 - All Supplies including those liable for Reverse charge - B2BA(cum) Invoices (IMS) - Summary",
        "2B-Summary-CDNR_DN(ITC_Avl)": "3. ITC Available - PART A1 - B2B Debit Notes - Summary",
        "2B-Summary-CDNR_CN(ITC_Avl)": "3. ITC Available - PART B1 - B2B Credit Notes (IMS) - Summary",
        "2B-Summary-CDNR_RC(ITC_Avl)": "3. ITC Available - PART B1 - B2B Credit/Debit Notes (Reverse charge) - Summary",
        "2B-Summary-IMPG(ITC_Avl)": "3. ITC Available - PART A4 - Import of goods from overseas - Summary",
        "2B-Summary-IMPGA(ITC_Avl)": "3. ITC Available - PART A4 - Import of goods from overseas (Amendment) - Summary",
        "2B-Summary-B2B(ITC_Rej)": "6. ITC Rejected - PART A1 - Supplies other than Reverse charge - B2B Invoices (IMS) - Summary",
    }
    column_headers_map_summary = {key: ["Month"] + GSTR2B_SUMMARY_SHEET_TOTAL_COLUMNS for key in
                                  summary_data_output.keys()}

    column_formats_map_summary = {}
    for key in summary_data_output.keys():
        formats = {"Month": "General"}
        for col in GSTR2B_SUMMARY_SHEET_TOTAL_COLUMNS:
            formats[col] = INDIAN_FORMAT_GSTR2B
        column_formats_map_summary[key] = formats

    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    title_font_style = Font(bold=True, size=12)
    center_alignment = Alignment(horizontal="center", vertical="center")

    def sheet_has_valid_data(rows_list, numeric_cols_check):
        if not rows_list: return False
        for r_item in rows_list:
            for header_item in numeric_cols_check:
                if header_item in r_item and parse_number(r_item.get(header_item, 0)) != 0:
                    return True
        return False

    for sheet_key_name, monthly_data_list in summary_data_output.items():
        current_numeric_cols = [h for h in GSTR2B_SUMMARY_SHEET_TOTAL_COLUMNS if not (
                    sheet_key_name.startswith("2B-Summary-IMPG") and h in ["Central Tax", "State/UT Tax"])]

        if not sheet_has_valid_data(monthly_data_list, current_numeric_cols):
            print(f"[DEBUG] Skipping summary sheet {sheet_key_name} due to no valid data.")
            continue

        if sheet_key_name in wb.sheetnames: wb.remove(wb[sheet_key_name])
        ws = wb.create_sheet(sheet_key_name)
        ws.freeze_panes = "B3"

        current_sheet_headers = column_headers_map_summary.get(sheet_key_name,
                                                               ["Month"] + GSTR2B_SUMMARY_SHEET_TOTAL_COLUMNS)
        current_sheet_formats = column_formats_map_summary.get(sheet_key_name, {})

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(current_sheet_headers))
        title_cell = ws.cell(row=1, column=1, value=section_titles_summary.get(sheet_key_name, sheet_key_name))
        title_cell.font = title_font_style
        title_cell.alignment = center_alignment

        for col_idx, col_name_val in enumerate(current_sheet_headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=col_name_val)
            cell.font = BOLD_FONT
            cell.fill = header_fill
            cell.alignment = center_alignment

        for row_idx, data_item in enumerate(monthly_data_list, start=3):
            for col_idx, col_name_val in enumerate(current_sheet_headers, start=1):
                cell_value = data_item.get(col_name_val, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
                if col_name_val in current_sheet_formats:
                    cell.number_format = current_sheet_formats[col_name_val]

        _add_total_row_to_gstr2b_summary_sheet(ws, monthly_data_list, current_sheet_headers, current_sheet_formats)

        for col_idx, header_val in enumerate(current_sheet_headers, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            # MODIFICATION: Start from row 2 for width calculation
            for r_num in range(2, ws.max_row + 1):
                cell_val_str = str(ws.cell(row=r_num, column=col_idx).value or "")
                max_len = max(max_len, len(cell_val_str))
            ws.column_dimensions[col_letter].width = max(15, max_len + 2)

    print("[DEBUG] Finished creating summary sheets")


# ----------------------- Excel Report Generation (Modified) ----------------------- #
def create_excel_report(data_dict, save_path, template_path=None):
    if template_path and os.path.exists(template_path):
        wb = load_workbook(template_path)
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames: wb.remove(wb["Sheet"])

    section_titles_detail = {
        "2B-B2B": "Taxable inward supplies received from registered person : 2B-B2B",
        "2B-B2BA": "Amendments to previously filed invoices by supplier : 2B-B2BA",
        "2B-B2BA(cum)": "Amendments to previously filed invoices by supplier : 2B-B2BA (Cumulative)",
        "2B-CDNR": "Debit/Credit notes(Original) : 2B-CDNR",
        "2B-IMPG": "Import of Goods : 2B-IMPG",
        "2B-B2B(ITC_Rej)": "Taxable inward supplies received from registered person : 2B-B2B (ITC Rejected)",
        "2B-B2B_sws": "Taxable inward supplies received from registered person : 2B-B2B - Sorted Supplier_wise",
        "2B-CDNR_sws": "Debit/Credit notes(Original) : 2B-CDNR - Sorted Supplier_wise"
    }

    column_headers_detail = {
        "2B-B2B": ["GSTIN of supplier", "Trade/legal name", "Invoice number", "Invoice Part", "Invoice type",
                   "Tax Rate", "Invoice Date", "Invoice Value", "Place of supply", "Supply Attract Reverse Charge",
                   "Total Taxable Value", "Integrated Tax", "Central Tax", "State/UT Tax", "Cess", "GSTR Period",
                   "GSTR Filing Date", "GSTR Filing Period", "ITC Availability", "Reason", "Source"],
        "2B-B2BA": ["GSTIN of supplier", "Trade/legal name", "Original Invoice number", "Original Invoice Date",
                    "Invoice number", "Invoice Part", "Invoice Date", "Invoice type", "Tax Rate", "Invoice Value",
                    "Place of supply", "Supply Attract Reverse Charge", "Total Taxable Value", "Integrated Tax",
                    "Central Tax", "State/UT Tax", "Cess", "GSTR Period", "GSTR Filing Date", "GSTR Filing Period",
                    "ITC Availability", "Reason"],
        "2B-B2BA(cum)": ["GSTIN of supplier", "Trade/legal name", "Total Documents", "Calculated Invoice Value",
                         "Total Taxable Value", "Integrated Tax", "Central Tax", "State/UT Tax", "Cess", "GSTR Period",
                         "GSTR Filing Date", "GSTR Filing Period"],
        "2B-CDNR": ["GSTIN of supplier", "Trade/legal name", "Note number", "Note Part", "Note type", "Tax Rate",
                    "Note supply type", "Note Date", "Note Value", "Place of supply", "Supply Attract Reverse Charge",
                    "Total Taxable Value", "Integrated Tax", "Central Tax", "State/UT Tax", "Cess", "GSTR Period",
                    "GSTR Filing Date", "GSTR Filing Period", "ITC Availability", "Reason"],
        "2B-IMPG": ["ICEGATE Reference Date", "Port Code", "Bill of Entry Number", "Bill of Entry Date",
                    "Calculated Invoice Value", "Taxable Value", "Integrated Tax", "Cess", "Record Date",
                    "GSTR Filing Period", "Amended (Yes)"],
        "2B-B2B(ITC_Rej)": ["GSTIN of supplier", "Trade/legal name", "Invoice number", "Invoice Part", "Invoice type",
                            "Tax Rate", "Invoice Date", "Invoice Value", "Place of supply",
                            "Supply Attract Reverse Charge", "Total Taxable Value", "Integrated Tax", "Central Tax",
                            "State/UT Tax", "Cess", "GSTR Period", "GSTR Filing Date", "GSTR Filing Period", "Source"]
    }
    for key in list(column_headers_detail.keys()):
        if key in ["2B-B2B", "2B-CDNR"]:
            column_headers_detail[key + "_sws"] = column_headers_detail[key]

    column_formats_map_detail = {
        "2B-B2B": {"Invoice Date": "dd-mm-yyyy", "Invoice Value": INDIAN_FORMAT_GSTR2B,
                   "Total Taxable Value": INDIAN_FORMAT_GSTR2B, "Integrated Tax": INDIAN_FORMAT_GSTR2B,
                   "Central Tax": INDIAN_FORMAT_GSTR2B, "State/UT Tax": INDIAN_FORMAT_GSTR2B,
                   "Cess": INDIAN_FORMAT_GSTR2B},
        "2B-B2BA": {"Original Invoice Date": "dd-mm-yyyy", "Invoice Date": "dd-mm-yyyy",
                    "Invoice Value": INDIAN_FORMAT_GSTR2B, "Total Taxable Value": INDIAN_FORMAT_GSTR2B,
                    "Integrated Tax": INDIAN_FORMAT_GSTR2B, "Central Tax": INDIAN_FORMAT_GSTR2B,
                    "State/UT Tax": INDIAN_FORMAT_GSTR2B, "Cess": INDIAN_FORMAT_GSTR2B},
        "2B-B2BA(cum)": {"Total Documents": "#,##0", "Calculated Invoice Value": INDIAN_FORMAT_GSTR2B,
                         "Total Taxable Value": INDIAN_FORMAT_GSTR2B, "Integrated Tax": INDIAN_FORMAT_GSTR2B,
                         "Central Tax": INDIAN_FORMAT_GSTR2B, "State/UT Tax": INDIAN_FORMAT_GSTR2B,
                         "Cess": INDIAN_FORMAT_GSTR2B},
        "2B-CDNR": {"Note Date": "dd-mm-yyyy", "Note Value": INDIAN_FORMAT_GSTR2B,
                    "Total Taxable Value": INDIAN_FORMAT_GSTR2B, "Integrated Tax": INDIAN_FORMAT_GSTR2B,
                    "Central Tax": INDIAN_FORMAT_GSTR2B, "State/UT Tax": INDIAN_FORMAT_GSTR2B,
                    "Cess": INDIAN_FORMAT_GSTR2B},
        "2B-IMPG": {"Bill of Entry Date": "dd-mm-yyyy", "Calculated Invoice Value": INDIAN_FORMAT_GSTR2B,
                    "Taxable Value": INDIAN_FORMAT_GSTR2B, "Integrated Tax": INDIAN_FORMAT_GSTR2B,
                    "Cess": INDIAN_FORMAT_GSTR2B},
        "2B-B2B(ITC_Rej)": {"Invoice Date": "dd-mm-yyyy", "Invoice Value": INDIAN_FORMAT_GSTR2B,
                            "Total Taxable Value": INDIAN_FORMAT_GSTR2B, "Integrated Tax": INDIAN_FORMAT_GSTR2B,
                            "Central Tax": INDIAN_FORMAT_GSTR2B, "State/UT Tax": INDIAN_FORMAT_GSTR2B,
                            "Cess": INDIAN_FORMAT_GSTR2B}
    }
    for key in list(column_formats_map_detail.keys()):
        if key in ["2B-B2B", "2B-CDNR"]:
            column_formats_map_detail[key + "_sws"] = column_formats_map_detail[key]

    header_fill_style = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    title_font_main = Font(bold=True, size=12)
    center_align_style = Alignment(horizontal="center", vertical="center")

    def write_detail_sheet(sheet_name, data_rows, headers_list):
        if not data_rows:
            print(f"[DEBUG] No data for sheet {sheet_name}. Skipping sheet generation.")
            return

        if sheet_name in wb.sheetnames: wb.remove(wb[sheet_name])
        ws = wb.create_sheet(sheet_name)
        ws.freeze_panes = "B3"

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers_list))
        title_cell = ws.cell(row=1, column=1, value=section_titles_detail.get(sheet_name, sheet_name))
        title_cell.font = title_font_main
        title_cell.alignment = center_align_style

        for idx, h_val in enumerate(headers_list, start=1):
            c = ws.cell(row=2, column=idx, value=h_val)
            c.font = BOLD_FONT
            c.fill = header_fill_style
            c.alignment = center_align_style

        current_formats_for_sheet = column_formats_map_detail.get(sheet_name, {})
        for r_idx, r_item_data in enumerate(data_rows, start=3):
            for c_idx, h_val in enumerate(headers_list, start=1):
                val_to_write = r_item_data.get(h_val, "")
                cell = ws.cell(row=r_idx, column=c_idx, value=val_to_write)
                fmt_str = current_formats_for_sheet.get(h_val)
                if fmt_str: cell.number_format = fmt_str

        if sheet_name in GSTR2B_DETAIL_SHEET_TOTAL_COLUMNS:
            _add_total_row_to_gstr2b_detail_sheet(ws, sheet_name, data_rows, headers_list, current_formats_for_sheet)

        for c_idx, _ in enumerate(headers_list, start=1):
            col_letter = get_column_letter(c_idx)
            max_len = 0
            # MODIFICATION: Start from row 2 for width calculation
            for r_num in range(2, ws.max_row + 1):
                cell_val_str = str(ws.cell(row=r_num, column=c_idx).value or "")
                max_len = max(max_len, len(cell_val_str))
            ws.column_dimensions[col_letter].width = max(15, max_len + 2)

    for key_name, list_of_rows in data_dict.items():
        if key_name in column_headers_detail:
            write_detail_sheet(key_name, list_of_rows, column_headers_detail[key_name])

    if "2B-B2B" in data_dict and data_dict["2B-B2B"]:
        sws_b2b = sorted(
            data_dict["2B-B2B"],
            key=lambda r: (r.get("Trade/legal name", "").lower(),
                           datetime.datetime.strptime(r.get("Invoice Date", "01-01-1900"), "%d-%m-%Y") if r.get(
                               "Invoice Date") else datetime.datetime.min)
        )
        if sws_b2b: write_detail_sheet("2B-B2B_sws", sws_b2b, column_headers_detail["2B-B2B_sws"])

    if "2B-CDNR" in data_dict and data_dict["2B-CDNR"]:
        sws_cdnr = sorted(
            data_dict["2B-CDNR"],
            key=lambda r: (r.get("Trade/legal name", "").lower(),
                           datetime.datetime.strptime(r.get("Note Date", "01-01-1900"), "%d-%m-%Y") if r.get(
                               "Note Date") else datetime.datetime.min)
        )
        if sws_cdnr: write_detail_sheet("2B-CDNR_sws", sws_cdnr, column_headers_detail["2B-CDNR_sws"])

    create_summary_sheets(wb, data_dict)

    wb.save(save_path)
    return f"✅ Successfully saved Excel file with totals: {save_path}"


# ----------------------- Main Processing Function ----------------------- #
def process_gstr2b(json_files, template_path, save_path):
    combined_data = {
        "2B-B2B": [], "2B-B2BA": [], "2B-B2BA(cum)": [],
        "2B-CDNR": [], "2B-IMPG": [], "2B-B2B(ITC_Rej)": []
    }
    json_payloads = {}
    all_new_sections_found = set()
    all_extra_subsections_found = {}

    for file_path in json_files:
        basename = os.path.basename(file_path)
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except Exception as e:
            send_event("error", {"module": "gstr2b_processor", "file": basename, "error": f"Failed to load JSON: {e}"})
            continue

        json_payloads[basename] = payload
        data_node_main = payload.get("data", {})

        for section_type_key, allowed_sub_keys_set in ALLOWED_DOC_SECTIONS.items():
            actual_section_content = data_node_main.get(section_type_key)
            if actual_section_content is None:
                continue

            if isinstance(actual_section_content, dict):
                observed_sub_keys_set = set(actual_section_content.keys())
                extra_keys_found = observed_sub_keys_set - allowed_sub_keys_set
                if extra_keys_found:
                    all_extra_subsections_found.setdefault(section_type_key, set()).update(extra_keys_found)
            else:
                all_new_sections_found.add(
                    f"Unexpected type for {section_type_key} (expected dict): {type(actual_section_content)}")

        filing_period_str = get_tax_period(data_node_main.get("rtnprd", ""))

        for extractor_function in (
        extract_b2b, extract_b2ba, extract_b2ba_cum, extract_cdnr, extract_impg, extract_b2b_itc_rej):
            try:
                section_data_dict = extractor_function(payload, filing_period_str)
                for key_from_extractor, rows_from_extractor in section_data_dict.items():
                    combined_data.setdefault(key_from_extractor, []).extend(rows_from_extractor)
            except Exception as e:
                print(f"[ERROR] Failed during extraction with {extractor_function.__name__} for file {basename}: {e}")
                send_event("error",
                           {"module": "gstr2b_processor", "function": extractor_function.__name__, "file": basename,
                            "error": str(e)})

    def parse_date_robustly(date_val_str, date_fmt="%d-%m-%Y"):
        if not date_val_str or not isinstance(date_val_str, str): return datetime.datetime.min
        try:
            return datetime.datetime.strptime(date_val_str, date_fmt)
        except ValueError:
            if len(date_val_str) == 6 and date_val_str.isdigit():
                try:
                    return datetime.datetime.strptime(date_val_str, "%m%Y")
                except ValueError:
                    return datetime.datetime.min
            return datetime.datetime.min

    date_field_mapping = {
        "2B-B2B": "Invoice Date", "2B-B2BA": "Invoice Date", "2B-B2B(ITC_Rej)": "Invoice Date",
        "2B-CDNR": "Note Date", "2B-IMPG": "Bill of Entry Date",
        "2B-B2BA(cum)": "GSTR Filing Date"
    }

    for data_key, sort_date_field in date_field_mapping.items():
        if data_key in combined_data:  # Check if key exists before trying to sort
            combined_data[data_key].sort(key=lambda r: parse_date_robustly(r.get(sort_date_field, "")))

    report_generation_msg = create_excel_report(combined_data, save_path, template_path)

    final_output_message = report_generation_msg
    warning_message_parts = []
    if all_new_sections_found:
        warning_message_parts.append(
            f"new/unexpected sections/types: {', '.join(sorted(list(all_new_sections_found)))}")
    if all_extra_subsections_found:
        extras_str_list = []
        for sec, vals in all_extra_subsections_found.items():
            extras_str_list.append(f"{sec} -> [{','.join(sorted(list(vals)))}]")
        warning_message_parts.append(f"unexpected subsections: {'; '.join(extras_str_list)}")

    if warning_message_parts:
        final_output_message = f"✅ Completed with warnings: {' | '.join(warning_message_parts)}"
        send_event("gstr2b_unexpected_structure", {
            "file_count": len(json_files),
            "new_sections_found": list(all_new_sections_found),
            "extra_subsections_found": {k: list(v) for k, v in all_extra_subsections_found.items()},
        })

    send_event("gstr2b_complete", {
        "output_file": save_path, "file_count": len(json_files),
        "message": final_output_message
    })
    return final_output_message
