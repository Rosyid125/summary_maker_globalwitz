"""
Debug GUI Column Mapping
Test the exact same column mapping that GUI uses to see if it works correctly
"""

import logging
import sys
import os

# Add the src directory to the path so we can import the modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__)))

from src.core.js_excel_reader import JSStyleExcelReader
from src.core.data_aggregator import DataAggregator
from src.core.js_processor import JSStyleProcessor

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_gui_column_mapping():
    """Test the GUI column mapping exactly like it would be done in the GUI"""
    
    file_path = r"original_excel\done team export\India-Import-jan-jun-2025.xlsx"
    sheet_name = "DATA OLAH"
    
    # This is the column mapping that user would set in GUI
    # Based on your testing, you said "price" works when mapped to these columns
    column_mappings = {
        'date': 'CUSTOMS CLEARANCE DATE',
        'hs_code': 'HS Code',  # Note: GUI uses 'hs_code' but JSStyleExcelReader expects 'hsCode'
        'item_description': 'PRODUCT DESCRIPTION(EN)',
        'gsm': '',  # Empty - will use default fallback
        'item': 'ITEM',
        'add_on': 'ADD ON', 
        'importer': 'PURCHASER',
        'supplier': 'SUPPLIER',
        'origin_country': 'ORIGIN COUNTRY',
        'unit_price': 'Estimated CIF Value $',  # This is what user maps as price
        'quantity': 'BUSINESS QUANTITY (KG)'
    }
    
    # Initialize components like GUI does
    js_excel_reader = JSStyleExcelReader(logger)
    data_aggregator = DataAggregator(logger)
    js_processor = JSStyleProcessor(logger)
    
    print(f"Testing GUI column mapping with file: {file_path}")
    print(f"Sheet: {sheet_name}")
    print(f"Column mappings: {column_mappings}")
    print("-" * 80)
    
    try:
        # Step 1: Read and preprocess data like GUI does
        print("Step 1: Reading and preprocessing data...")
        
        # Convert GUI column mapping to what JSStyleExcelReader expects
        # GUI uses different keys than JSStyleExcelReader internal keys
        reader_column_mapping = {
            'date': column_mappings.get('date'),
            'hsCode': column_mappings.get('hs_code'),  # Convert hs_code -> hsCode
            'itemDesc': column_mappings.get('item_description'),
            'gsm': column_mappings.get('gsm'),
            'item': column_mappings.get('item'),
            'addOn': column_mappings.get('add_on'),
            'importer': column_mappings.get('importer'), 
            'supplier': column_mappings.get('supplier'),
            'originCountry': column_mappings.get('origin_country'),
            'unitPrice': column_mappings.get('unit_price'),  # Convert unit_price -> unitPrice
            'quantity': column_mappings.get('quantity')
        }
        
        # Remove empty mappings
        reader_column_mapping = {k: v for k, v in reader_column_mapping.items() if v}
        
        print(f"Converted mapping for reader: {reader_column_mapping}")
        
        all_raw_data = js_excel_reader.read_and_preprocess_data(
            file_path,
            sheet_name,
            "auto",  # date_format
            "auto",  # number_format  
            reader_column_mapping
        )
        
        if not all_raw_data:
            print("ERROR: No data was read from the file!")
            return
            
        print(f"Read {len(all_raw_data)} rows")
        
        # Show first few rows to verify price parsing
        print("\nFirst 3 rows after preprocessing:")
        for i, row in enumerate(all_raw_data[:3]):
            print(f"  Row {i+1}: month='{row.get('month')}', usdQtyUnit={row.get('usdQtyUnit')}, qty={row.get('qty')}")
        
        # Step 2: Process data like GUI does
        print("\nStep 2: Processing data with JSStyleProcessor...")
        
        output_path = js_processor.process_data_like_javascript(
            all_raw_data,
            "2025",  # period_year
            "FOB",   # global_incoterm
            "test_gui_mapping_output.xlsx"
        )
        
        if output_path:
            print(f"SUCCESS: Output created at {output_path}")
        else:
            print("ERROR: Failed to create output")
            
        # Step 3: Check if price values are non-zero in the raw data
        print("\nStep 3: Checking price values in processed data...")
        non_zero_prices = [row for row in all_raw_data if row.get('usdQtyUnit', 0) > 0]
        print(f"Found {len(non_zero_prices)} rows with non-zero prices out of {len(all_raw_data)} total rows")
        
        if non_zero_prices:
            print("Sample non-zero price rows:")
            for i, row in enumerate(non_zero_prices[:5]):
                print(f"  Row: usdQtyUnit={row.get('usdQtyUnit')}, qty={row.get('qty')}, supplier='{row.get('supplier')}'")
        else:
            print("WARNING: No non-zero price values found!")
            
        # Step 4: Test aggregation directly
        print("\nStep 4: Testing aggregation directly...")
        agg_result = data_aggregator.perform_aggregation(all_raw_data)
        
        print(f"Aggregation result: Level1={len(agg_result['summaryLvl1'])}, Level2={len(agg_result['summaryLvl2'])}")
        
        if agg_result['summaryLvl1']:
            print("Sample Level1 aggregation results:")
            for i, row in enumerate(agg_result['summaryLvl1'][:3]):
                print(f"  Level1 Row {i+1}: month='{row.get('month')}', avgPrice={row.get('avgPrice')}, totalQty={row.get('totalQty')}")
        
        if agg_result['summaryLvl2']:
            print("Sample Level2 aggregation results:")
            for i, row in enumerate(agg_result['summaryLvl2'][:3]):
                print(f"  Level2 Row {i+1}: avgOfSummaryPrice={row.get('avgOfSummaryPrice')}, totalOfSummaryQty={row.get('totalOfSummaryQty')}")
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_gui_column_mapping()
