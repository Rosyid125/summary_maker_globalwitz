#!/usr/bin/env python3 
"""
Test full process from reading to output generation
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.js_excel_reader import JSStyleExcelReader
from src.core.data_aggregator import DataAggregator
from src.core.js_output_formatter import OutputFormatter
from src.utils.logger import setup_logger

def test_full_process():
    """Test complete process with new file"""
    logger = setup_logger()
    reader = JSStyleExcelReader(logger)
    aggregator = DataAggregator(logger)
    formatter = OutputFormatter(logger)
    
    input_file = "original_excel/done team export/India-Import-jan-jun-2025.xlsx"
    
    print("=== TESTING FULL PROCESS ===")
    
    # Column mapping for Estimated CIF Value
    column_mapping = {
        'date': 'Date',
        'hsCode': 'HS Code', 
        'itemDesc': 'Product Description',
        'importer': 'Consignee Name',
        'supplier': 'Shipper Name',
        'quantity': 'Standard Qty',
        'unitPrice': 'Estimated CIF Value $'  # Total value, not unit price
    }
    
    print("Step 1: Reading data...")
    raw_data = reader.read_and_preprocess_data(
        input_file,
        sheet_name="DATA OLAH",
        date_format="auto",
        number_format="AMERICAN",
        column_mapping=column_mapping
    )
    
    if not raw_data:
        print("Failed to read data")
        return
        
    print(f"Read {len(raw_data)} rows")
    
    # Sample check
    print("\nSample raw data (first 3 rows):")
    for i, row in enumerate(raw_data[:3]):
        print(f"Row {i+1}: hsCode={row.get('hsCode')}, month={row.get('month')}, "
              f"usdQtyUnit={row.get('usdQtyUnit')}, qty={row.get('qty')}")
    
    print("\nStep 2: Grouping by supplier...")
    # Group by supplier like in main process
    grouped_by_supplier = {}
    for row in raw_data:
        supplier = row.get('supplier', 'Unknown')
        if supplier not in grouped_by_supplier:
            grouped_by_supplier[supplier] = []
        grouped_by_supplier[supplier].append(row)
    
    print(f"Found {len(grouped_by_supplier)} suppliers")
    
    # Test with first supplier
    first_supplier = list(grouped_by_supplier.keys())[0]
    first_supplier_data = grouped_by_supplier[first_supplier]
    
    print(f"\nStep 3: Processing supplier '{first_supplier}' ({len(first_supplier_data)} rows)")
    
    # Aggregate 
    result = aggregator.perform_aggregation(first_supplier_data)
    
    if not result:
        print("Aggregation failed")
        return
        
    summary_lvl1 = result.get('summaryLvl1', [])
    summary_lvl2 = result.get('summaryLvl2', [])
    
    print(f"Level 1 summary: {len(summary_lvl1)} items")
    print(f"Level 2 summary: {len(summary_lvl2)} items")
    
    # Check sample data  
    print("\nSample Level 1 data:")
    for i, item in enumerate(summary_lvl1[:3]):
        print(f"  {i+1}. month={item.get('month')}, hsCode={item.get('hsCode')}, "
              f"avgPrice={item.get('avgPrice')}, totalQty={item.get('totalQty')}")
    
    print("\nSample Level 2 data:")
    for i, item in enumerate(summary_lvl2[:3]):
        print(f"  {i+1}. hsCode={item.get('hsCode')}, "
              f"avgOfSummaryPrice={item.get('avgOfSummaryPrice')}, "
              f"totalOfSummaryQty={item.get('totalOfSummaryQty')}")
    
    print("\nStep 4: Preparing group block...")
    
    # Test prepareGroupBlock equivalent
    group_block = formatter.prepare_group_block(
        first_supplier,
        summary_lvl1,
        summary_lvl2,
        "CIF"  # incoterm value
    )
    
    print(f"Group block prepared with {len(group_block.get('groupBlockRows', []))} rows")
    
    # Check the actual data in the group block 
    group_rows = group_block.get('groupBlockRows', [])
    if len(group_rows) >= 3:
        print("\nSample group block rows:")
        print("Header 1:", group_rows[0])
        print("Header 2:", group_rows[1])
        if len(group_rows) > 2:
            print("Data row 1:", group_rows[2])
    
    print("\n=== ANALYSIS ===")
    # Let's specifically check if avgPrice values are being correctly passed
    print("Checking if avgPrice values are preserved in group block...")
    
    for i, row in enumerate(group_rows[2:5]):  # Skip headers, check first 3 data rows
        if len(row) > 6:  # Make sure row has enough columns
            print(f"Data row {i+1}:")
            print(f"  Full row: {row}")
            # Check price columns (every 2nd column starting from index 5 for months)
            month_cols = []
            for month_idx in range(12):  # 12 months
                price_col_idx = 5 + (month_idx * 2)
                qty_col_idx = 5 + (month_idx * 2) + 1
                if price_col_idx < len(row) and qty_col_idx < len(row):
                    price_val = row[price_col_idx]
                    qty_val = row[qty_col_idx]
                    if price_val != "-" and qty_val != "-":
                        month_cols.append(f"Month{month_idx+1}: price={price_val}, qty={qty_val}")
            print(f"  Month data: {month_cols}")

if __name__ == "__main__":
    test_full_process()
