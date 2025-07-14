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
import os
import sys

def get_app_data_dir():
    """Get the application data directory that works in both dev and built versions"""
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        app_dir = os.path.dirname(sys.executable)
    else:
        # Running in development
        app_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return app_dir

def get_safe_output_dir():
    """Get a safe output directory that works in both dev and built versions"""
    # Always use the processed_excel folder relative to the application directory
    app_dir = get_app_data_dir()
    output_dir = os.path.join(app_dir, "processed_excel")
    
    # Ensure directory exists
    os.makedirs(output_dir, exist_ok=True)
    return output_dir

DEFAULT_INPUT_FOLDER = os.path.join(get_app_data_dir(), "original_excel")
DEFAULT_OUTPUT_FOLDER = get_safe_output_dir()
DEFAULT_SHEET_NAME = "DATA OLAH"
