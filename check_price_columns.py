#!/usr/bin/env python3
"""
Check all potential price columns in the Excel file
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import pandas as pd

def check_price_columns():
    """Check which columns might contain price data"""
    
    input_file = "original_excel/US-Import-jan-jun-2025.xlsx"
    
    print("=== CHECKING ALL POTENTIAL PRICE COLUMNS ===")
    
    try:
        df = pd.read_excel(input_file, sheet_name="DATA OLAH")
        print(f"Total rows: {len(df)}")
        
        # Look for columns that might contain price data
        price_columns = []
        for col in df.columns:
            col_lower = col.lower()
            if any(keyword in col_lower for keyword in ['value', 'price', 'unit', 'rate', 'usd', 'cost']):
                price_columns.append(col)
        
        print(f"\nFound {len(price_columns)} potential price columns:")
        
        for col in price_columns:
            print(f"\n--- Column: {col} ---")
            values = df[col].dropna()
            if len(values) > 0:
                print(f"  Non-null values: {len(values)}")
                print(f"  Data type: {values.dtype}")
                print(f"  Sample values: {values.head(10).tolist()}")
                
                # Only calculate numeric stats for numeric columns
                if values.dtype in ['int64', 'float64']:
                    print(f"  Min: {values.min()}")
                    print(f"  Max: {values.max()}")
                    print(f"  Mean: {values.mean():.2f}")
                    
                    # Count non-zero values
                    non_zero = values[values != 0]
                    print(f"  Non-zero values: {len(non_zero)}")
                    if len(non_zero) > 0:
                        print(f"  Non-zero sample: {non_zero.head(10).tolist()}")
                else:
                    print(f"  This is a text column - checking unique values")
                    unique_vals = values.unique()
                    print(f"  Unique values ({len(unique_vals)}): {unique_vals[:10]}")
            else:
                print("  No values found")
        
        # Also check a few other potentially relevant columns
        other_cols = ['BUSINESS QUANTITY (KG)', 'Std. Quantity', 'Quantity']
        print(f"\n=== CHECKING QUANTITY COLUMNS ===")
        for col in other_cols:
            if col in df.columns:
                print(f"\n--- Column: {col} ---")
                values = df[col].dropna()
                if len(values) > 0:
                    print(f"  Non-null values: {len(values)}")
                    print(f"  Data type: {values.dtype}")
                    print(f"  Min: {values.min()}")
                    print(f"  Max: {values.max()}")
                    print(f"  Sample values: {values.head(10).tolist()}")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_price_columns()
