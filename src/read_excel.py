from pathlib import Path
from openpyxl import load_workbook
import csv
from datetime import datetime
from config import (TABLE_SCHEMAS, PRIMARY_KEYS, INPUT_FOLDER)
from utils import (add_review, export_csv, load_lookup, get_cell, get_outlet_id, detect_header_row, 
                   discover_product_blocks, clean_numeric_value)


def get_product_name(
        ws, 
        row, 
        block, 
        header_row
    ):
    if block["name_source"] == "header":
        return ws.cell(
            row=header_row, 
            column=block["name_col"]
        ).value
    
    product_name = ws.cell(
        row=row, 
        column=block["name_col"]
        ).value

    if product_name is None:
        return None
    
    return str(product_name).strip()

def extract_product_activity(
        ws,
        row,
        block,
        visit_id,
        visit_date,
        outlet_name,
        workbook_name,
        worksheet_name,
        product_lookup,
        header_row,
        review_items
):
    
    # Read product name
    product_name = get_product_name(ws, row, block, header_row)

    if product_name is None:
        return None

    # Read activity
    transaction_type = get_cell(ws, row, block["transaction_type_col"])
    opening_stock_raw = get_cell(ws, row, block["opening_stock_col"])
    closing_stock_raw = get_cell(ws, row, block["closing_stock_col"])
    sales_qty_raw = get_cell(ws, row, block["sales_qty_col"])
    bad_stock_raw = get_cell(ws, row, block["bad_stock_col"])
    tester_raw = get_cell(ws, row, block["tester_col"])
    transaction_value_raw = get_cell(ws, row, block["transaction_value_col"])

    # Skip if there is no activity
    if closing_stock_raw is None and sales_qty_raw is None:
        return None

    # Lookup product ID
    product_id = PRODUCT_NAME_TO_ID.get(product_name)

    if product_id is None:

        add_review(
            review_items,
            severity="ERROR",
            reason="Unknown product",
            workbook=workbook_name,
            worksheet=worksheet_name,
            row=row,
            outlet_name=outlet_name,
            product_name=product_name,
            details="Product not found in Products worksheet."
        )

        return None

    # Clean numeric fields
    opening_stock = clean_numeric_value(
        opening_stock_raw,
        field_name="opening_stock",
        workbook_name=workbook_name,
        worksheet_name=worksheet_name,
        row=row,
        outlet_name=outlet_name,
        product_name=product_name,
        review_items=review_items,
    )

    closing_stock = clean_numeric_value(
        closing_stock_raw,
        field_name="closing_stock",
        workbook_name=workbook_name,
        worksheet_name=worksheet_name,
        row=row,
        outlet_name=outlet_name,
        product_name=product_name,
        review_items=review_items,
    )

    sales_qty = clean_numeric_value(
        sales_qty_raw,
        field_name="sales_qty",
        workbook_name=workbook_name,
        worksheet_name=worksheet_name,
        row=row,
        outlet_name=outlet_name,
        product_name=product_name,
        review_items=review_items,
    )

    bad_stock = clean_numeric_value(
        bad_stock_raw,
        field_name="bad_stock",
        workbook_name=workbook_name,
        worksheet_name=worksheet_name,
        row=row,
        outlet_name=outlet_name,
        product_name=product_name,
        review_items=review_items,
    )

    tester = clean_numeric_value(
        tester_raw,
        field_name="tester",
        workbook_name=workbook_name,
        worksheet_name=worksheet_name,
        row=row,
        outlet_name=outlet_name,
        product_name=product_name,
        review_items=review_items,
    )

    transaction_value = clean_numeric_value(
        transaction_value_raw,
        field_name="transaction_value",
        workbook_name=workbook_name,
        worksheet_name=worksheet_name,
        row=row,
        outlet_name=outlet_name,
        product_name=product_name,
        review_items=review_items,
    )

    visit_detail = {
        "visit_id": visit_id,
        "visit_date": visit_date,
        "worksheet": worksheet_name,
        "outlet_name": outlet_name,
        "product_id": product_id,
        "product_name": product_name,
        "transaction_type": transaction_type,
        "opening_stock": opening_stock,
        "closing_stock": closing_stock,
        "sales_qty": sales_qty,
        "bad_stock": bad_stock,
        "tester": tester,
        "transaction_value": transaction_value,
    }

    return visit_detail

PRODUCT_NAME_TO_ID, PRODUCT_ID_TO_NAME = load_lookup(
    workbook_path=Path("reference-data/products.xlsx"),
    sheet_name="products",
    key_field="product_name",
    value_field="product_id"
)

# Path to current folder
excel_files = list(INPUT_FOLDER.glob("*.xlsx"))
                   
if not excel_files:
    print("No Excel files found in the current folder.")
    exit()

if not excel_files:
    print("No Excel files found in the current folder.")
    exit()

outlet_count = 0

outlets = []
visits = []
visit_details = []
review_items = []

outlet_lookup = {}
outlet_records = {}
next_outlet_id = 1

next_visit_id = 1

for workbook_path in excel_files:
    print(f"\nReading workbook: {workbook_path.name}")
    wb = load_workbook(workbook_path, data_only=True)

    for sheet in wb.sheetnames:
        print(f"Worksheet: {sheet}")
        ws = wb[sheet]
        header_row = detect_header_row(ws)

        has_status_column = (
            ws.cell(row=header_row, column=4).value is not None
            and str(ws.cell(row=header_row, column=4).value).strip().lower() == "status"
        )


        product_blocks = discover_product_blocks(
            ws,
            header_row,
        )

        visit_date = datetime.strptime(ws.title, "%m-%d-%Y").date()

        blank_rows = 0

        for row in range (header_row + 1, ws.max_row + 1):
            outlet_name = ws[f"B{row}"].value
            address = ws[f"C{row}"].value

            if has_status_column:
                status = ws[f"D{row}"].value
            else:
                status = None
                

            outlet_id, next_outlet_id = get_outlet_id(
                outlet_lookup, 
                outlet_name, 
                address, 
                next_outlet_id
            )

            # Stop after 2 consecutive blank rows
            if outlet_name is None or str(outlet_name).strip() == "":
                blank_rows += 1

                if blank_rows >= 2:
                    print(f"Stopping at row {row} due to consecutive blank rows.")
                    break

                continue

            blank_rows = 0

            outlet_count += 1

            filename = workbook_path.stem
            route = filename.removeprefix("Form ").split(" - ", 1)[0]

            visit_id = f"VIS{next_visit_id:06d}"
            next_visit_id += 1

            # Store outlets
            outlet = {
                "outlet_id": outlet_id,
                "worksheet": ws.title,
                "row": row,
                "outlet_name": outlet_name,
                "address": address,
                }

            outlet_records.setdefault(outlet_id, outlet)

            visit = {
                "outlet_id": outlet_id,
                "visit_id": visit_id,
                "route": route,
                "visit_date": visit_date,
                "worksheet": ws.title,
                "row": row,
                "outlet_name": outlet_name,
                "address": address,
                "status": status
                }
            
            visits.append(visit)


            # Read every product block
            for block in product_blocks:

                visit_detail = extract_product_activity(
                    ws=ws,
                    row=row,
                    block=block,
                    visit_id=visit_id,
                    visit_date=visit_date,
                    outlet_name=outlet_name,
                    workbook_name=workbook_path.name,
                    worksheet_name=sheet,
                    product_lookup=PRODUCT_NAME_TO_ID,
                    header_row=header_row,
                    review_items=review_items
                )

                if visit_detail is not None:
                    visit_details.append(visit_detail)
            
outlets = list(outlet_records.values())

print(f"\nTotal outlets: {len(outlets)}")
print(f"Total visits: {len(visits)}")
print(f"Total visit details: {len(visit_details)}")

export_csv("outlets.csv", outlets, TABLE_SCHEMAS["outlets"])
export_csv("visits.csv", visits, TABLE_SCHEMAS["visits"])
export_csv("visit_details.csv", visit_details, TABLE_SCHEMAS["visit_details"])
export_csv("review_items.csv", review_items, TABLE_SCHEMAS["review_items"]) 