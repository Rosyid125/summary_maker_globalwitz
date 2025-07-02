#!/usr/bin/env python3
"""
Test dengan file Excel yang sama seperti yang Anda gunakan
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.js_excel_reader import JSStyleExcelReader
from src.core.data_aggregator import DataAggregator
from src.core.js_output_formatter import OutputFormatter
from src.utils.logger import setup_logger

def test_with_your_file():
    """Test dengan file dan mapping yang sama"""
    logger = setup_logger()
    reader = JSStyleExcelReader(logger)
    aggregator = DataAggregator(logger) 
    formatter = OutputFormatter(logger)
    
    input_file = "original_excel/done team export/India-Import-jan-jun-2025.xlsx"
    
    print("=== TESTING WITH YOUR EXACT SETUP ===")
    
    # Test different mapping scenarios
    test_cases = [
        {
            'name': 'Case 1: Estimated CIF Value $ (Total Value)',
            'mapping': {
                'date': 'Date',
                'hsCode': 'HS Code',
                'itemDesc': 'Product Description', 
                'item': 'ITEM',
                'gsm': 'GSM',
                'addOn': 'ADD ON',
                'importer': 'Consignee Name',
                'supplier': 'Shipper Name',
                'quantity': 'Standard Qty',
                'unitPrice': 'Estimated CIF Value $'
            }
        },
        {
            'name': 'Case 2: Standard Unit Rate $ (Unit Price)',
            'mapping': {
                'date': 'Date',
                'hsCode': 'HS Code',
                'itemDesc': 'Product Description',
                'item': 'ITEM',
                'gsm': 'GSM', 
                'addOn': 'ADD ON',
                'importer': 'Consignee Name',
                'supplier': 'Shipper Name',
                'quantity': 'Standard Qty',
                'unitPrice': 'Standard Unit Rate $'
            }
        },
        {
            'name': 'Case 3: Calculate Unit Price (CIF Value / Quantity)',
            'mapping': {
                'date': 'Date',
                'hsCode': 'HS Code',
                'itemDesc': 'Product Description',
                'item': 'ITEM',
                'gsm': 'GSM',
                'addOn': 'ADD ON', 
                'importer': 'Consignee Name',
                'supplier': 'Shipper Name',
                'quantity': 'Standard Qty',
                'unitPrice': 'CALCULATED'  # We'll calculate CIF/Qty
            }
        }
    ]
    
    for test_case in test_cases:
        print(f"\n{'='*60}")
        print(f"TESTING: {test_case['name']}")
        print(f"{'='*60}")
        
        try:
            if test_case['mapping']['unitPrice'] == 'CALCULATED':
                # Special case: calculate unit price
                raw_data = reader.read_and_preprocess_data_with_calc(
                    input_file,
                    sheet_name="DATA OLAH",
                    date_format="auto",
                    number_format="AMERICAN",
                    column_mapping=test_case['mapping']
                )
            else:
                # Normal case
                raw_data = reader.read_and_preprocess_data(
                    input_file,
                    sheet_name="DATA OLAH", 
                    date_format="auto",
                    number_format="AMERICAN",
                    column_mapping=test_case['mapping']
                )
            
            if not raw_data:
                print("Failed to read data")
                continue
                
            print(f"Read {len(raw_data)} rows")
            
            # Sample the data
            print("Sample price values (first 5 rows):")
            for i, row in enumerate(raw_data[:5]):
                usd_qty = row.get('usdQtyUnit', 'NOT_FOUND')
                qty = row.get('qty', 'NOT_FOUND')
                month = row.get('month', 'NOT_FOUND')
                supplier = row.get('supplier', 'NOT_FOUND')
                print(f"  Row {i+1}: supplier={supplier[:30]}..., month={month}, usdQtyUnit={usd_qty}, qty={qty}")
            
            # Test with one supplier
            grouped_by_supplier = {}
            for row in raw_data:
                supplier = row.get('supplier', 'Unknown')
                if supplier not in grouped_by_supplier:
                    grouped_by_supplier[supplier] = []
                grouped_by_supplier[supplier].append(row)
            
            # Get a supplier with reasonable amount of data
            target_supplier = None
            target_data = None
            for supplier, data in grouped_by_supplier.items():
                if len(data) >= 5:  # At least 5 records
                    target_supplier = supplier
                    target_data = data
                    break
            
            if not target_supplier:
                target_supplier = list(grouped_by_supplier.keys())[0]
                target_data = grouped_by_supplier[target_supplier]
            
            print(f"\nTesting with supplier: {target_supplier} ({len(target_data)} records)")
            
            # Aggregate
            result = aggregator.perform_aggregation(target_data)
            
            if result:
                summary_lvl1 = result.get('summaryLvl1', [])
                summary_lvl2 = result.get('summaryLvl2', [])
                
                print(f"Aggregation result: Level1={len(summary_lvl1)}, Level2={len(summary_lvl2)}")
                
                if summary_lvl1:
                    print("Sample Level 1 (first 3):")
                    for i, item in enumerate(summary_lvl1[:3]):
                        print(f"  {i+1}. {item.get('month')} - {item.get('hsCode')} - avgPrice={item.get('avgPrice')} - totalQty={item.get('totalQty')}")
                
                if summary_lvl2:
                    print("Sample Level 2 (first 3):")
                    for i, item in enumerate(summary_lvl2[:3]):
                        print(f"  {i+1}. {item.get('hsCode')} - avgOfSummaryPrice={item.get('avgOfSummaryPrice')} - totalOfSummaryQty={item.get('totalOfSummaryQty')}")
                
                # Test group block creation
                group_block = formatter.prepare_group_block(
                    target_supplier,
                    summary_lvl1,
                    summary_lvl2,
                    "CIF"
                )
                
                group_rows = group_block.get('groupBlockRows', [])
                print(f"\nGroup block created with {len(group_rows)} rows")
                
                # Check first data row (skip headers)
                if len(group_rows) > 2:
                    data_row = group_rows[2]
                    print("First data row preview:")
                    print(f"  Supplier: {data_row[0]}")
                    print(f"  HS Code: {data_row[1]}")
                    print(f"  Item: {data_row[2]}")
                    print(f"  GSM: {data_row[3]}")
                    print(f"  Add On: {data_row[4]}")
                    
                    # Check month data
                    for month_idx in range(min(3, 12)):  # Check first 3 months
                        price_idx = 5 + (month_idx * 2)
                        qty_idx = 5 + (month_idx * 2) + 1
                        if price_idx < len(data_row) and qty_idx < len(data_row):
                            price_val = data_row[price_idx]
                            qty_val = data_row[qty_idx]
                            if price_val != "-" and qty_val != "-":
                                month_name = ['Jan', 'Feb', 'Mar'][month_idx]
                                print(f"  {month_name}: price={price_val}, qty={qty_val}")
                    
                    # Check recap
                    recap_price_idx = len(data_row) - 3
                    recap_qty_idx = len(data_row) - 1
                    if recap_price_idx >= 0 and recap_qty_idx >= 0:
                        print(f"  RECAP: avgPrice={data_row[recap_price_idx]}, totalQty={data_row[recap_qty_idx]}")
            
        except Exception as e:
            print(f"Error in {test_case['name']}: {e}")
            import traceback
            traceback.print_exc()

def add_calculated_unit_price_support():
    """Add support for calculated unit price to reader"""
    # This is a simple extension, not production code
    
    def read_and_preprocess_data_with_calc(self, input_file_path, sheet_name="DATA OLAH", 
                                         date_format='DD/MM/YYYY', number_format='EUROPEAN',
                                         column_mapping=None):
        """Extended version that can calculate unit price from total value and quantity"""
        import pandas as pd
        
        try:
            df = pd.read_excel(input_file_path, sheet_name=sheet_name)
            self.logger.info(f"Reading {len(df)} rows from sheet '{sheet_name}' with calculated unit price...")
            
            processed_data = []
            
            for index, row in df.iterrows():
                # Get total value and quantity
                total_value = self.parse_number(row.get('Estimated CIF Value $', 0), number_format)
                quantity = self.parse_number(row.get('Standard Qty', 0), number_format)
                
                # Calculate unit price
                unit_price = total_value / quantity if quantity > 0 else 0
                
                # Process other fields normally
                processed_row = {
                    'date': self.parse_date(row.get(column_mapping.get('date', ''), ''), date_format),
                    'month': None,  # Will be set from date
                    'hsCode': str(row.get(column_mapping.get('hsCode', ''), '')),
                    'itemDesc': str(row.get(column_mapping.get('itemDesc', ''), '')),
                    'item': str(row.get(column_mapping.get('item', ''), '')),
                    'gsm': str(row.get(column_mapping.get('gsm', ''), '')),
                    'addOn': str(row.get(column_mapping.get('addOn', ''), '')),
                    'importer': str(row.get(column_mapping.get('importer', ''), '')),
                    'supplier': str(row.get(column_mapping.get('supplier', ''), '')),
                    'qty': quantity,
                    'usdQtyUnit': unit_price  # Calculated unit price
                }
                
                # Set month from date
                if processed_row['date']:
                    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 
                                 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des']
                    processed_row['month'] = month_names[processed_row['date'].month - 1]
                
                # Log sample
                if index < 5:
                    self.logger.info(f"Row {index+1}: total_value={total_value}, qty={quantity}, calculated_unit_price={unit_price}")
                
                processed_data.append(processed_row)
            
            self.logger.info(f"Processed {len(processed_data)} rows with calculated unit prices")
            return processed_data
            
        except Exception as e:
            self.logger.error(f"Error in calculated unit price processing: {e}")
            return None
    
    # Monkey patch the method
    JSStyleExcelReader.read_and_preprocess_data_with_calc = read_and_preprocess_data_with_calc

if __name__ == "__main__":
    add_calculated_unit_price_support()
    test_with_your_file()
