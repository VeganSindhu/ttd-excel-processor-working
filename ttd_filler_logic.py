import pandas as pd
from openpyxl import load_workbook
import re
import os

# ---------------- HELPERS ----------------

def clean_mobile(mobile):
    if pd.isna(mobile):
        return ""
    digits = re.sub(r"\D", "", str(mobile))
    if digits.startswith("91") and len(digits) > 10:
        digits = digits[2:]
    return digits if len(digits) == 10 else ""

def split_address(addr):
    """
    Remove last 3 parts (city, state, pincode)
    Split remaining into 3 non-empty lines
    """
    if pd.isna(addr) or not addr:
        return "", "", ""

    parts = [p.strip() for p in str(addr).split(",") if p.strip()]

    # remove last 3 parts
    if len(parts) > 3:
        parts = parts[:-3]

    if len(parts) == 0:
        return "", "", ""

    if len(parts) == 1:
        return parts[0], parts[0], parts[0]

    if len(parts) == 2:
        return parts[0], parts[1], parts[1]

    return parts[0], parts[1], ", ".join(parts[2:])

def get_dimensions(qty):
    qty = int(qty)
    if qty >= 5:
        return 57, 44, 2
    return 25, 18, 4

# ---------------- MAIN LOGIC ----------------

def generate_output(orders_path, postal_path):
    output_path = "Matching_Output.xlsx"

    # Load Postal
    postal = pd.read_excel(postal_path, header=3)

    tr_col = postal.columns[1]
    name_col = postal.columns[2]
    addr_col = postal.columns[3]
    city_col = postal.columns[4]
    pin_col = postal.columns[5]
    mobile_col = postal.columns[6]
    qty_col = postal.columns[7]
    weight_col = postal.columns[8]
    barcode_col = postal.columns[9]

    postal = postal[[tr_col, name_col, addr_col, city_col, pin_col,
                      mobile_col, qty_col, weight_col, barcode_col]].copy()

    postal["__TR"] = postal[tr_col].astype(str).str.strip()
    postal["Receiver name"] = postal[name_col]
    postal["Full Address"] = postal[addr_col]
    postal["Receiver city"] = postal[city_col]
    postal["Receiver pincode"] = pd.to_numeric(postal[pin_col], errors="coerce").fillna(0).astype(int)
    postal["Receiver mobile"] = postal[mobile_col].apply(clean_mobile)
    postal["Quantity"] = pd.to_numeric(postal[qty_col], errors="coerce").fillna(1).astype(int)
    postal["Physical weight in grams"] = (
        pd.to_numeric(postal[weight_col], errors="coerce").fillna(0).astype(int)
    )
    postal["Barcode"] = postal[barcode_col].astype(str).str.strip()

    postal = postal[postal["Receiver pincode"].between(100000, 999999)]

    # Load Orders (STATE ONLY)
    orders = pd.read_excel(orders_path, sheet_name="Publications_Report")
    orders["__TR"] = orders["Booking No"].astype(str).str.strip()
    orders = orders[["__TR", "State"]].drop_duplicates("__TR")

    merged = postal.merge(orders, on="__TR", how="left")
    merged["State"] = merged["State"].fillna("Tamil Nadu")

    # Load Template
    wb = load_workbook("TTD Template.xlsx")
    ws = wb.active

    headers = [c.value for c in ws[1]]
    defaults = [c.value for c in ws[2]]

    # Clear rows
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    serial_no = 1

    for _, row in merged.iterrows():
        r = 1 + serial_no

        L, B, H = get_dimensions(row["Quantity"])
        a1, a2, a3 = split_address(row["Full Address"])

        for i, h in enumerate(headers):
            cell = ws.cell(row=r, column=i + 1)
            h_norm = "" if h is None else h.lower()

            if "serial" in h_norm:
                cell.value = serial_no
            elif "barcode" in h_norm:
                cell.value = row["Barcode"]
            elif "physical weight" in h_norm:
                cell.value = row["Physical weight in grams"]
            elif "receiver name" in h_norm:
                cell.value = row["Receiver name"]
            elif "receiver mobile" in h_norm:
                cell.value = row["Receiver mobile"]
            elif "receiver city" in h_norm:
                cell.value = row["Receiver city"]
            elif "receiver pincode" in h_norm:
                cell.value = row["Receiver pincode"]
            elif "receiver state" in h_norm:
                cell.value = row["State"]
            elif "add line 1" in h_norm:
                cell.value = a1
            elif "add line 2" in h_norm:
                cell.value = a2
            elif "add line 3" in h_norm:
                cell.value = a3
            elif "length" in h_norm:
                cell.value = L
            elif "breadth" in h_norm or "diameter" in h_norm:
                cell.value = B
            elif "height" in h_norm:
                cell.value = H
            else:
                cell.value = defaults[i]

            # Sender (fixed)
            if "sender mobile" in h_norm:
                cell.value = 1234567890
            if "sender state" in h_norm:
                cell.value = "Andhra Pradesh"
            if "sender add line 1" in h_norm:
                cell.value = "SALES WING OF PUBLICATIONS"
            if "sender add line 2" in h_norm:
                cell.value = "TTD PRESS COMPOUND"
            if "sender add line 3" in h_norm:
                cell.value = "Tirupati-517507"

        serial_no += 1

    wb.save(output_path)
    return output_path, serial_no - 1
