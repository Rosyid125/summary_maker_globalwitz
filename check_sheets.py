#!/usr/bin/env python3
"""
Check sheet names in Excel file
"""

import pandas as pd

def check_sheets():
    try:
        excel_file = pd.ExcelFile('original_excel/US-Import-jan-jun-2025.xlsx')
        print('Available sheets:', excel_file.sheet_names)
        
        # Also check the first sheet content
        first_sheet = excel_file.sheet_names[0]
        print(f'\nReading first sheet: {first_sheet}')
        df = pd.read_excel('original_excel/US-Import-jan-jun-2025.xlsx', sheet_name=first_sheet)
        print(f'Shape: {df.shape}')
        print('Columns:', df.columns.tolist())
        
        # Check if there's a Date column and show sample dates
        if 'Date' in df.columns:
            print('\nSample dates from Date column:')
            print(df['Date'].head(10).tolist())
        else:
            print('\nNo "Date" column found. Looking for date-like columns...')
            for col in df.columns:
                if 'date' in col.lower() or 'tanggal' in col.lower():
                    print(f'Found date-like column: {col}')
                    print(df[col].head(5).tolist())
        
    except Exception as e:
        print('Error:', e)

if __name__ == "__main__":
    check_sheets()
