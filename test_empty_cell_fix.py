#!/usr/bin/env python3
"""
Test script to verify the empty cell fix in total rows
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.js_excel_reader import JSStyleExcelReader
from src.core.js_processor import JSStyleProcessor
from src.utils.logger import setup_logger

def test_empty_cell_fix():
    """Test that empty cells in TOTAL QTY rows show '-' instead of blank"""
    logger = setup_logger()
    reader = JSStyleExcelReader(logger)
    processor = JSStyleProcessor(logger)
    
    input_file = "original_excel/done team export/India-Import-jan-jun-2025.xlsx"
    
    print("=== TESTING EMPTY CELL FIX IN TOTAL ROWS ===")
    
    # GUI-style column mapping
    gui_column_mapping = {
        'date': 'Date',
        'hs_code': 'HS Code', 
        'item_description': 'Product Description',
        'importer': 'Consignee Name',
        'supplier': 'Shipper Name',
        'origin_country': 'Country of Origin',
        'quantity': 'Standard Qty',
        'unit_price': 'Standard Unit Rate $'
    }
    
    # Read data
    raw_data = reader.read_and_preprocess_data(
        input_file,
        sheet_name="DATA OLAH",
        date_format="auto",
        number_format="EUROPEAN", 
        column_mapping=gui_column_mapping
    )
    
    if raw_data:
        print(f"✅ Read {len(raw_data)} rows")
        
        # Process and create output
        output_path = processor.process_data_like_javascript(
            raw_data,
            "2025",
            "CIF",
            "test_empty_cell_fix_output.xlsx"
        )
        
        if output_path:
            print(f"✅ Output created: {output_path}")
            print("\nSekarang cek file output untuk memastikan:")
            print("- Baris 'TOTAL QTY PER MO' menampilkan '-' untuk cell kosong")
            print("- Baris 'TOTAL QTY PER QUARTAL' menampilkan '-' untuk cell kosong") 
            print("- Tidak ada cell yang benar-benar kosong (blank)")
        else:
            print("❌ Failed to create output")
    else:
        print("❌ Failed to read data")

if __name__ == "__main__":
    test_empty_cell_fix()
