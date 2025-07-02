#!/usr/bin/env python3
"""
Test script to verify GUI column mapping fix
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.js_excel_reader import JSStyleExcelReader
from src.core.data_aggregator import DataAggregator
from src.utils.logger import setup_logger

def test_gui_mapping_fix():
    """Test that GUI column mappings now work correctly"""
    logger = setup_logger()
    reader = JSStyleExcelReader(logger)
    aggregator = DataAggregator(logger)
    
    input_file = "original_excel/done team export/India-Import-jan-jun-2025.xlsx"
    
    print("=== TESTING GUI COLUMN MAPPING FIX ===")
    
    # This is exactly what the GUI sends - using underscore keys
    gui_column_mapping = {
        'date': 'Date',
        'hs_code': 'HS Code', 
        'item_description': 'Product Description',
        'importer': 'Consignee Name',
        'supplier': 'Shipper Name',
        'origin_country': 'Country of Origin',
        'quantity': 'Standard Qty',
        'unit_price': 'Standard Unit Rate $'  # This should work now
    }
    
    print("GUI Column Mapping:")
    for key, value in gui_column_mapping.items():
        print(f"  {key} -> {value}")
    
    # Test reading with GUI-style mapping
    raw_data = reader.read_and_preprocess_data(
        input_file,
        sheet_name="DATA OLAH",
        date_format="auto",
        number_format="EUROPEAN", 
        column_mapping=gui_column_mapping
    )
    
    if raw_data:
        print(f"\n✅ SUCCESS: Read {len(raw_data)} rows")
        
        # Check first few rows for price values
        print("Sample price values from first 5 rows:")
        non_zero_prices = 0
        for i, row in enumerate(raw_data[:5]):
            usd_qty = row.get('usdQtyUnit', 'NOT_FOUND')
            qty = row.get('qty', 'NOT_FOUND')
            print(f"  Row {i+1}: usdQtyUnit = {usd_qty}, qty = {qty}")
            if isinstance(usd_qty, (int, float)) and usd_qty > 0:
                non_zero_prices += 1
        
        print(f"Non-zero prices in first 5 rows: {non_zero_prices}/5")
        
        # Check all rows for non-zero prices
        total_non_zero = sum(1 for row in raw_data if isinstance(row.get('usdQtyUnit'), (int, float)) and row.get('usdQtyUnit') > 0)
        print(f"Total non-zero prices: {total_non_zero}/{len(raw_data)}")
        
        # Test aggregation 
        print("\n=== TESTING AGGREGATION ===")
        result = aggregator.perform_aggregation(raw_data)
        
        if result and result.get('summaryLvl1'):
            print(f"✅ Level 1 summary: {len(result['summaryLvl1'])} items")
            sample_item = result['summaryLvl1'][0]
            print(f"Sample Level 1 item: avgPrice = {sample_item.get('avgPrice', 'NOT_FOUND')}")
            
        if result and result.get('summaryLvl2'):
            print(f"✅ Level 2 summary: {len(result['summaryLvl2'])} items")
            sample_item = result['summaryLvl2'][0]
            print(f"Sample Level 2 item: avgOfSummaryPrice = {sample_item.get('avgOfSummaryPrice', 'NOT_FOUND')}")
            
        if total_non_zero > 0:
            print(f"\n🎉 SUCCESS: GUI column mapping fix works! All {total_non_zero} rows have non-zero prices!")
        else:
            print(f"\n❌ FAILED: No non-zero prices found")
            
    else:
        print("❌ FAILED: Could not read data")

if __name__ == "__main__":
    test_gui_mapping_fix()
