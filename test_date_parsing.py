#!/usr/bin/env python3
"""
Test script for date parsing functionality
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.js_excel_reader import JSStyleExcelReader
from src.utils.logger import setup_logger

def test_date_parsing():
    """Test date parsing with various formats"""
    logger = setup_logger()
    reader = JSStyleExcelReader(logger)
    
    # Test dates from the Excel file
    test_dates = [
        "12-Apr-2025",
        "08-May-2025", 
        "12-May-2025",
        "01-Jan-2025",
        "28-Apr-2025",
        "03-Mar-2025",
        "29-May-2025",
        "17-Apr-2025",
        "23-Mar-2025"
    ]
    
    print("Testing date parsing with 'auto' format...")
    for date_str in test_dates:
        parsed = reader.parse_date(date_str, 'auto')
        if parsed:
            month_name = reader.get_month_name(parsed)
            print(f"✅ {date_str} -> {parsed.strftime('%Y-%m-%d')} (Month: {month_name})")
        else:
            print(f"❌ {date_str} -> Failed to parse")
    
    print("\nTesting date parsing with 'DD-MONTH-YYYY' format...")
    for date_str in test_dates:
        parsed = reader.parse_date(date_str, 'DD-MONTH-YYYY')
        if parsed:
            month_name = reader.get_month_name(parsed)
            print(f"✅ {date_str} -> {parsed.strftime('%Y-%m-%d')} (Month: {month_name})")
        else:
            print(f"❌ {date_str} -> Failed to parse")

if __name__ == "__main__":
    test_date_parsing()
