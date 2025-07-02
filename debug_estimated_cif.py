#!/usr/bin/env python3
"""
Specific debug script for Estimated CIF Value issue
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
from src.core.js_excel_reader import JSStyleExcelReader
from src.core.data_aggregator import DataAggregator
from src.utils.logger import setup_logger

def debug_estimated_cif_value():
    """Debug the Estimated CIF Value column processing issue"""
    logger = setup_logger()
    reader = JSStyleExcelReader(logger)
    aggregator = DataAggregator(logger)
    
    input_file = "original_excel/done team export/India-Import-jan-jun-2025.xlsx"
    
    print("=== DEBUGGING ESTIMATED CIF VALUE ISSUE ===")
    
    # Test with different price column mappings
    test_mappings = [
        {
            'name': 'Standard Unit Rate $',
            'mapping': {
                'date': 'Date',
                'hsCode': 'HS Code', 
                'itemDesc': 'Product Description',
                'importer': 'Consignee Name',
                'supplier': 'Shipper Name',
                'quantity': 'Standard Qty',
                'unitPrice': 'Standard Unit Rate $'
            }
        },
        {
            'name': 'Estimated CIF Value $',
            'mapping': {
                'date': 'Date',
                'hsCode': 'HS Code', 
                'itemDesc': 'Product Description',  
                'importer': 'Consignee Name',
                'supplier': 'Shipper Name',
                'quantity': 'Standard Qty',
                'unitPrice': 'Estimated CIF Value $'
            }
        },
        {
            'name': 'Unit Rate $',
            'mapping': {
                'date': 'Date',
                'hsCode': 'HS Code', 
                'itemDesc': 'Product Description',
                'importer': 'Consignee Name',
                'supplier': 'Shipper Name',
                'quantity': 'Standard Qty',
                'unitPrice': 'Unit Rate $'
            }
        }
    ]
    
    for test_case in test_mappings:
        print(f"\n=== TESTING WITH {test_case['name']} COLUMN ===")
        
        try:
            raw_data = reader.read_and_preprocess_data(
                input_file,
                sheet_name="DATA OLAH",
                date_format="auto",
                number_format="AMERICAN",  # Using AMERICAN format as you mentioned
                column_mapping=test_case['mapping']
            )
            
            if raw_data:
                print(f"Read {len(raw_data)} rows")
                print("Sample price values from first 5 rows:")
                for i, row in enumerate(raw_data[:5]):
                    usd_qty = row.get('usdQtyUnit', 'NOT_FOUND')
                    qty = row.get('qty', 'NOT_FOUND')
                    month = row.get('month', 'NOT_FOUND')
                    hs_code = row.get('hsCode', 'NOT_FOUND')
                    print(f"  Row {i+1}: month={month}, hsCode={hs_code}, usdQtyUnit={usd_qty}, qty={qty}")
                
                # Test aggregation 
                print("\n=== TESTING AGGREGATION ===")
                result = aggregator.perform_aggregation(raw_data)
                
                if result and result.get('summaryLvl1'):
                    print(f"Level 1 summary has {len(result['summaryLvl1'])} items")
                    print("Sample Level 1 items (first 3):")
                    for i, item in enumerate(result['summaryLvl1'][:3]):
                        print(f"  Item {i+1}: month={item.get('month')}, hsCode={item.get('hsCode')}, "
                              f"avgPrice={item.get('avgPrice', 'NOT_FOUND')}, totalQty={item.get('totalQty', 'NOT_FOUND')}")
                
                if result and result.get('summaryLvl2'):
                    print(f"Level 2 summary has {len(result['summaryLvl2'])} items")  
                    print("Sample Level 2 items (first 3):")
                    for i, item in enumerate(result['summaryLvl2'][:3]):
                        print(f"  Item {i+1}: hsCode={item.get('hsCode')}, "
                              f"avgOfSummaryPrice={item.get('avgOfSummaryPrice', 'NOT_FOUND')}, "
                              f"totalOfSummaryQty={item.get('totalOfSummaryQty', 'NOT_FOUND')}")
                        
            else:
                print("Failed to read data")
                
        except Exception as e:
            print(f"Error with {test_case['name']}: {e}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    debug_estimated_cif_value()
