#!/usr/bin/env python3
"""
Test script for processing real Excel file with auto date format
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.js_excel_reader import JSStyleExcelReader
from src.core.js_processor import JSStyleProcessor
from src.utils.logger import setup_logger

def test_real_excel_processing():
    """Test processing the real Excel file"""
    logger = setup_logger()
    reader = JSStyleExcelReader(logger)
    processor = JSStyleProcessor(logger)
    
    input_file = "original_excel/US-Import-jan-jun-2025.xlsx"
    
    print(f"Testing processing of real Excel file: {input_file}")
    
    # Test with auto date format
    print("Using auto date format...")
    
    try:
        # Create column mapping for the real Excel structure
        column_mapping = {
            'date': 'Arrival Date',
            'hsCode': 'HS Code',
            'itemDesc': 'Product Description',
            'importer': 'Consignee Name',
            'supplier': 'Shipper Name',
            'quantity': 'Std. Quantity',
            'unitPrice': 'Value CIF US$'
        }
        
        # Read data
        raw_data = reader.read_and_preprocess_data(
            input_file,
            sheet_name="DATA OLAH",  # Use the correct sheet name
            date_format="auto",  # Use auto detection
            number_format="EUROPEAN",
            column_mapping=column_mapping  # Pass the column mapping
        )
        
        if not raw_data:
            print("❌ Failed to read data")
            return
        
        print(f"✅ Successfully read {len(raw_data)} rows")
        
        # Print sample data to verify dates are parsed
        print("\nSample of parsed data:")
        for i, row in enumerate(raw_data[:5]):
            date_val = row.get('date')
            month_val = row.get('month')
            print(f"Row {i+1}: Date={date_val}, Month={month_val}")
        
        # Process data
        output_file = "test_real_output.xlsx"  # Simple filename without path
        result = processor.process_data_like_javascript(
            raw_data,
            "2025",  # period year
            "-",     # global_incoterm - Use "-" as default
            output_file
        )
        
        if result:
            print(f"✅ Processing successful! Output saved to: {result}")
        else:
            print("❌ Processing failed")
            
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_real_excel_processing()
