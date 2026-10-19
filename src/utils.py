from pathlib import Path
from openpyxl import load_workbook
import csv
import re
from datetime import datetime

def add_review( 
        review_items,
        severity, 
        reason,
        workbook, 
        worksheet, 
        row, 
        outlet_name = "", 
        product_name = "", 
        details = "", 
):
    review_items.append({
        "severity": severity,
        "reason": reason,
        "workbook": workbook,
        "worksheet": worksheet,
        "row": row,
        "outlet_name": outlet_name,
        "product_name": product_name,
        "details": details,
    })

def export_csv(filename,data,columns):
    if not data:
        print(f"No data to export for {filename}.")
        return
    
    output_folder = Path("Output")
    output_folder.mkdir(exist_ok=True)

    with open(output_folder / filename, mode="w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=columns)
        writer.writeheader()
        writer.writerows(data)

    print(f"{filename} exported successfully.")

def load_lookup(workbook_path, sheet_name, key_field, value_field):
    '''
    Load a lookup table from an Excel sheet
    Returns:
    - key_to_value: a dictionary mapping keys to values
    - value_to_key: a dictionary mapping values to keys
    '''

    wb = load_workbook(workbook_path, data_only=True)
    ws = wb[sheet_name]
    headers = {}
    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col).value
        if header is not None:
            headers[str(header).strip()] = col
    key_col = headers[key_field]
    value_col = headers[value_field]

    key_to_value = {}
    value_to_key = {}

    for row in range(2, ws.max_row + 1):
        key = ws.cell(row=row, column=key_col).value
        value = ws.cell(row=row, column=value_col).value
        if key is None and value is None:
            continue

        key_to_value[key] = value
        value_to_key[value] = key

    return key_to_value, value_to_key

def get_cell(ws, row, col):
    """
    Returns the cell value.
    If col is None, returns None instead of raising an error.
    """
    if col is None:
        return None

    value =  ws.cell(row=row, column=col).value

    if isinstance(value,str):
        value = value.strip()

    return value

def get_outlet_id(outlet_lookup, outlet_name, address, next_outlet_id):
    """
    Returns the outlet_id for the given outlet_name and address.
    If the outlet does not exist, creates a new outlet_id.
    """
    key = (str(outlet_name).strip().lower(), str(address).strip().lower())

    if key in outlet_lookup:
        return outlet_lookup[key], next_outlet_id

    # Create a new outlet_id
    outlet_id = f"WRG-{next_outlet_id:05d}"
    outlet_lookup[key] = outlet_id
    next_outlet_id += 1

    return outlet_id, next_outlet_id

def detect_header_row(ws):

    """
    Find the worksheet header row.

    The header row is identified by the 'Outlet Name' header
    in column B.
    """

    for row in range(5, 15):

        cell = ws.cell(row=row, column=2).value

        if (
            cell is not None
            and str(cell).strip().lower() == "outlet name"
        ):
            return row

    raise ValueError(
        f"Could not determine worksheet header: {ws.title}"
    )

CURRENT_LAYOUT = {
    7: {
        "name_source": "header",
        "fields": {
            "transaction_type_col": 0,
            "opening_stock_col": 1,
            "closing_stock_col": 2,
            "sales_qty_col": 3,
            "bad_stock_col": 4,
            "tester_col": 5,
            "transaction_value_col": 6,
        },
    },
    8: {
        "name_source": "dropdown",
        "fields": {
            "transaction_type_col": 1,
            "opening_stock_col": 2,
            "closing_stock_col": 3,
            "sales_qty_col": 4,
            "bad_stock_col": 5,
            "tester_col": 6,
            "transaction_value_col": 7,
        },
    },
}

LEGACY_LAYOUT = {
    5: {
        "name_source": "header",
        "fields": {
            "transaction_type_col": None,
            "opening_stock_col": 0,
            "closing_stock_col": 1,
            "sales_qty_col": 2,
            "bad_stock_col": 3,
            "tester_col": 4,
            "transaction_value_col": None,
        },
    },
    6: {
        "name_source": "dropdown",
        "fields": {
            "transaction_type_col": None,
            "opening_stock_col": 1,
            "closing_stock_col": 2,
            "sales_qty_col": 3,
            "bad_stock_col": 4,
            "tester_col": 5,
            "transaction_value_col": None,
        },
    },
}

def discover_product_blocks(
        ws,
        header_row,
):
    """
    Discover product blocks from merged header cells.
    """

    product_blocks = []

    for merged in ws.merged_cells.ranges:

        # Only consider merged cells on the header row
        if merged.min_row != header_row:
            continue

        start_col = merged.min_col
        width = merged.max_col - merged.min_col + 1

        # Skip empty merged headers
        if ws.cell(row=header_row, column=start_col).value is None:
            continue

        # Determine the block type by its merged width
        definition = (
            CURRENT_LAYOUT.get(width)
            or LEGACY_LAYOUT.get(width)
        )

        # Ignore merged cells that are not product blocks
        if definition is None:
            continue

        block = {
            "name_source": definition["name_source"],
            "name_col": start_col,
        }

        for field, offset in definition["fields"].items():
            if offset is None:
                block[field] = None
            else:
                block[field] = start_col + offset

        product_blocks.append(block)

    return sorted(
        product_blocks,
        key=lambda b: b["name_col"]
    )

def clean_numeric_value(
    value,
    *,
    field_name,
    workbook_name,
    worksheet_name,
    row,
    outlet_name,
    product_name,
    review_items,
):
    """
    Clean legacy numeric fields.

    Rules:
    - Blank                  -> 0
    - *                      -> 0
    - x / X / -              -> 0
    - 0*                     -> 0
    - 10.0 / 10.             -> 10
    - *10 / 10*              -> 10
    - Tester date            -> None (historical marker, not a tester quantity)
    - Ambiguous values       -> None + review item
    """

    if value is None:
        return 0

    text = str(value).strip()

    if text == "" or text.lower() in {"none", "nan"}:
        return 0

    # Tester dates are historical markers, not tester quantities
    if field_name == "tester":
        if isinstance(value, datetime):
            return 0

        if re.fullmatch(
            r"\d{4}-\d{2}-\d{2}(?:\s+00:00:00)?",
            text
        ):
            return 0

    # Asterisk alone was a historical cash marker
    if text == "*":
        return 0

    # Values explicitly used as zero
    if text.lower() in {"x", "-"}:
        return 0

    # Remove asterisks used as decoration around a number
    cleaned = text.replace("*", "").strip()

    # Empty after removing decoration
    if cleaned == "":
        return None

    # Integer
    if re.fullmatch(r"-?\d+", cleaned):
        return int(cleaned)

    # Decimal representations such as 10.0 or 10.
    if re.fullmatch(r"-?\d+(?:\.0*)?", cleaned):
        return int(float(cleaned))

    # Anything else is ambiguous
    add_review(
        review_items,
        severity="WARNING",
        reason="Non-standard numeric value",
        workbook=workbook_name,
        worksheet=worksheet_name,
        row=row,
        outlet_name=outlet_name,
        product_name=product_name,
        details=f"{field_name} contains non-standard value: {text!r}"
    )

    return 0