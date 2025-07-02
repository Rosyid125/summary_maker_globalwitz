# Constants for Excel Summary Maker
# Matches the JavaScript constants exactly

MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]

# Column mappings
DEFAULT_COLUMN_MAPPING = {
    'date': ["DATE", "CUSTOMS CLEARANCE DATE"],
    'hs_code': ["HS CODE"],
    'item_desc': ["ITEM DESC", "PRODUCT DESCRIPTION(EN)"],
    'gsm': ["GSM"],
    'item': ["ITEM"],
    'add_on': ["ADD ON"],
    'importer': ["IMPORTER", "PURCHASER"],
    'supplier': ["SUPPLIER"],
    'origin_country': ["ORIGIN COUNTRY"],
    'unit_price': ["CIF KG Unit In USD", "USD Qty Unit", "UNIT PRICE(USD)"],
    'quantity': ["Net KG Wt", "qty", "BUSINESS QUANTITY (KG)"]
}

# Default values
DEFAULT_INPUT_FOLDER = "original_excel"
DEFAULT_OUTPUT_FOLDER = "processed_excel"
DEFAULT_SHEET_NAME = "DATA OLAH"
