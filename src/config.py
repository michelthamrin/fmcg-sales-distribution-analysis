from pathlib import Path

INPUT_FOLDER = Path("input")

TABLE_SCHEMAS = {
    "outlets": [
        "outlet_id",
        "worksheet", 
        "row", 
        "outlet_name", 
        "address", 
        "status"
    ],
    "visits": [
        "outlet_id",
        "visit_id", 
        "route",
        "visit_date", 
        "worksheet", 
        "row", 
        "outlet_name", 
        "address", 
        "status"
    ],
    "visit_details": [
        "visit_id",
        "visit_date",
        "worksheet",
        "outlet_name",
        "product_id",
        "product_name",
        "transaction_type",
        "opening_stock",
        "closing_stock",
        "sales_qty",
        "bad_stock",
        "tester",
        "transaction_value"
    ],
    "review_items": [
        "severity",
        "reason",
        "workbook",
        "worksheet",
        "row",
        "outlet_name",
        "product_name",
        "details"
    ]
}

PRIMARY_KEYS = {
    "outlets": ["worksheet", "row"],
    "visits": ["visit_id"],
    "visit_details": ["visit_id", "product_id"]
}

FOREIGN_KEYS = [
    {
    "child_table": "visit_details",
    "child_field": "visit_id",
    "parent_table": "visits",
    "parent_field": "visit_id"
    }
]

REQUIRED_FIELDS = {
    "outlets": [
        "worksheet",
        "row",
        "outlet_name"
    ],
    "visits": [
        "visit_id",
        "visit_date",
        "worksheet",
        "row"
    ],
    "visit_details": [
        "visit_id",
        "visit_date",
        "product_id",
        "product_name"
    ]
}
