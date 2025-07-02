#!/usr/bin/env python3
"""
Test script for number parsing logic
"""

import re

def test_parse_number(value, number_format='EUROPEAN'):
    """Test the parse_number logic"""
    if isinstance(value, str):
        cleaned_value = value.strip()
        if cleaned_value == "":
            return 0
        
        num_str = cleaned_value
        print(f"Input: {value}, Format: {number_format}")
        
        if number_format == 'AMERICAN':
            # American format: dot as decimal, comma as thousand separator
            american_regex = re.compile(r'^-?\d{1,3}(,\d{3})*(\.\d+)?$')
            print(f"  Testing against American regex: {american_regex.pattern}")
            match = american_regex.match(num_str)
            print(f"  Match result: {match}")
            if match:
                print(f"  Before replacement: {num_str}")
                num_str = num_str.replace(',', '')
                print(f"  After replacement: {num_str}")
        else:
            # European format: comma as decimal, dot as thousand separator
            european_regex = re.compile(r'^-?\d{1,3}(\.\d{3})*(,\d+)?$')
            print(f"  Testing against European regex: {european_regex.pattern}")
            match = european_regex.match(num_str)
            print(f"  Match result: {match}")
            if match:
                print(f"  Before replacement: {num_str}")
                num_str = num_str.replace('.', '').replace(',', '.')
                print(f"  After replacement: {num_str}")
        
        try:
            result = float(num_str)
            print(f"  Final result: {result}")
            return result
        except ValueError as e:
            print(f"  ValueError: {e}")
            return 0
    
    return 0

if __name__ == "__main__":
    print("=== TESTING NUMBER PARSING LOGIC ===")
    
    test_cases = [
        ('123.45', 'EUROPEAN'),
        ('1,234.56', 'EUROPEAN'),
        ('1.234,56', 'EUROPEAN'),
        ('1,234.56', 'AMERICAN'),
        ('1.234,56', 'AMERICAN'),
        ('123.45', 'AMERICAN'),
    ]
    
    for value, format_type in test_cases:
        print()
        result = test_parse_number(value, format_type)
        print(f"Final: '{value}' with {format_type} -> {result}")
        print("-" * 50)
