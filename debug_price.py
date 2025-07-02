#!/usr/bin/env python3
"""
Debug script for price column issue
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
from src.core.js_excel_reader import JSStyleExcelReader
from src.core.data_aggregator import DataAggregator
from src.utils.logger import setup_logger

def debug_price_issue():
    """Debug the price column processing issue"""
    logger = setup_logger()
    reader = JSStyleExcelReader(logger)
    aggregator = DataAggregator(logger)
    
    input_file = "original_excel/done team export/India-Import-jan-jun-2025.xlsx"
    
    print("=== DEBUGGING PRICE COLUMN ISSUE WITH NEW FILE ===")
    
    # First, let's check what columns are available in the Excel file
    try:
        df = pd.read_excel(input_file, sheet_name="DATA OLAH", nrows=5)
        print("\nAvailable columns in Excel file:")
        for i, col in enumerate(df.columns):
            print(f"  {i+1}. {col}")
        
        print("\nSample data from first 3 rows:")
        for i, row in df.iterrows():
            if i >= 3:
                break
            print(f"\nRow {i+1}:")
            for col in df.columns:
                if 'value' in col.lower() or 'price' in col.lower() or 'usd' in col.lower():
                    print(f"  {col}: {row[col]} (type: {type(row[col])})")
    
        # Test with different column mappings
        print("\n=== TESTING DIFFERENT COLUMN MAPPINGS ===")
        
        # Test 1: Using 'Standard Unit Rate $' column
        print("\nTest 1: Using 'Standard Unit Rate $' column")
        column_mapping = {
            'date': 'Date',
            'hsCode': 'HS Code', 
            'itemDesc': 'Product Description',
            'importer': 'Consignee Name',
            'supplier': 'Shipper Name',
            'quantity': 'Standard Qty',
            'unitPrice': 'Standard Unit Rate $'  # This should be the price column
        }
        
        raw_data = reader.read_and_preprocess_data(
            input_file,
            sheet_name="DATA OLAH",
            date_format="auto",
            number_format="EUROPEAN", 
            column_mapping=column_mapping
        )
        
        if raw_data:
            # Check first few rows for price values
            print(f"Read {len(raw_data)} rows")
            print("Sample price values from first 5 rows:")
            for i, row in enumerate(raw_data[:5]):
                usd_qty = row.get('usdQtyUnit', 'NOT_FOUND')
                qty = row.get('qty', 'NOT_FOUND')
                print(f"  Row {i+1}: usdQtyUnit = {usd_qty}, qty = {qty}")
            
            # Test aggregation 
            print("\n=== TESTING AGGREGATION ===")
            result = aggregator.perform_aggregation(raw_data)
            
            if result and result.get('summaryLvl1'):
                print(f"Level 1 summary has {len(result['summaryLvl1'])} items")
                print("Sample Level 1 items:")
                for i, item in enumerate(result['summaryLvl1'][:3]):
                    print(f"  Item {i+1}: avgPrice = {item.get('avgPrice', 'NOT_FOUND')}, totalQty = {item.get('totalQty', 'NOT_FOUND')}")
            
            if result and result.get('summaryLvl2'):
                print(f"Level 2 summary has {len(result['summaryLvl2'])} items")
                print("Sample Level 2 items:")
                for i, item in enumerate(result['summaryLvl2'][:3]):
                    print(f"  Item {i+1}: avgOfSummaryPrice = {item.get('avgOfSummaryPrice', 'NOT_FOUND')}, totalOfSummaryQty = {item.get('totalOfSummaryQty', 'NOT_FOUND')}")
        
        else:
            print("Failed to read data")
            
        # Test number parsing directly
        print("\n=== TESTING NUMBER PARSING ===")
        sample_values = ['123.45', '1,234.56', '1.234,56', '0', '', None, 100.50]
        for val in sample_values:
            parsed = reader.parse_number(val, 'EUROPEAN')
            print(f"  '{val}' -> {parsed}")
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_price_issue()
