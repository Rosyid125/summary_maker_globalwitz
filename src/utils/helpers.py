"""
Utility functions for Excel Summary Maker
"""

import re
from datetime import datetime, timedelta
from dateutil import parser
import pandas as pd
from typing import Union, Optional, List, Dict

class DateParser:
    """Handles various date format parsing"""
    
    @staticmethod
    def parse_date(date_value, date_format="auto"):
        """
        Parse date from various formats
        
        Args:
            date_value: The date value to parse
            date_format: Expected date format ("DD/MM/YYYY", "MM/DD/YYYY", "DD-MONTH-YYYY", "auto")
        
        Returns:
            datetime or None: Parsed date
        """
        if pd.isna(date_value) or date_value is None:
            return None
            
        # Handle Excel serial numbers
        if isinstance(date_value, (int, float)):
            try:
                # Excel epoch starts at 1900-01-01
                excel_epoch = datetime(1900, 1, 1)
                return excel_epoch + timedelta(days=date_value - 2)  # -2 for Excel bug
            except:
                return None
        
        # Convert to string
        date_str = str(date_value).strip()
        if not date_str:
            return None
        
        try:
            # Try specific format first
            if date_format == "DD/MM/YYYY":
                return datetime.strptime(date_str, "%d/%m/%Y")
            elif date_format == "MM/DD/YYYY":
                return datetime.strptime(date_str, "%m/%d/%Y")
            elif date_format == "DD-MONTH-YYYY":
                return DateParser._parse_month_name_format(date_str)
            
            # Auto-detect format
            return DateParser._auto_parse_date(date_str)
            
        except Exception:
            return None
    
    @staticmethod
    def _parse_month_name_format(date_str):
        """Parse date with month names (e.g., 15-JAN-2024)"""
        month_mapping = {
            'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MAY': 5, 'JUN': 6,
            'JUL': 7, 'AUG': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12,
            'JANUARI': 1, 'FEBRUARI': 2, 'MARET': 3, 'APRIL': 4, 'MEI': 5, 'JUNI': 6,
            'JULI': 7, 'AGUSTUS': 8, 'SEPTEMBER': 9, 'OKTOBER': 10, 'NOVEMBER': 11, 'DESEMBER': 12
        }
        
        # Try different separators
        for sep in ['-', ' ', '/']:
            if sep in date_str:
                parts = date_str.upper().split(sep)
                if len(parts) == 3:
                    try:
                        day = int(parts[0])
                        month = month_mapping.get(parts[1])
                        year = int(parts[2])
                        if month and 1 <= day <= 31 and year > 1900:
                            return datetime(year, month, day)
                    except (ValueError, TypeError):
                        continue
        return None
    
    @staticmethod
    def _auto_parse_date(date_str):
        """Auto-detect and parse date format"""
        try:
            # Try dateutil parser first
            return parser.parse(date_str, dayfirst=True)
        except:
            pass
        
        # Try common patterns
        patterns = [
            r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})',  # DD/MM/YYYY or MM/DD/YYYY
            r'(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})',  # YYYY/MM/DD
            r'(\d{6})',  # YYYYMM
        ]
        
        for pattern in patterns:
            match = re.match(pattern, date_str)
            if match:
                groups = match.groups()
                try:
                    if len(groups) == 3:
                        if len(groups[0]) == 4:  # YYYY first
                            year, month, day = map(int, groups)
                        else:  # Assume DD/MM/YYYY
                            day, month, year = map(int, groups)
                        return datetime(year, month, day)
                    elif len(groups) == 1 and len(groups[0]) == 6:  # YYYYMM
                        yyyymm = groups[0]
                        year = int(yyyymm[:4])
                        month = int(yyyymm[4:])
                        return datetime(year, month, 1)
                except (ValueError, TypeError):
                    continue
        
        return None

class NumberParser:
    """Handles number parsing with different locale formats"""
    
    @staticmethod
    def parse_number(value, number_format="auto"):
        """
        Parse number from string with locale support
        
        Args:
            value: The value to parse
            number_format: "european" (1.234,56) or "american" (1,234.56) or "auto"
        
        Returns:
            float or None: Parsed number
        """
        if pd.isna(value) or value is None:
            return None
        
        if isinstance(value, (int, float)):
            return float(value)
        
        # Convert to string and clean
        str_value = str(value).strip()
        if not str_value:
            return None
        
        # Remove currency symbols and spaces
        str_value = re.sub(r'[^\d,.\-+]', '', str_value)
        
        if not str_value or str_value in ['-', '+']:
            return None
        
        try:
            if number_format == "european":
                return NumberParser._parse_european_format(str_value)
            elif number_format == "american":
                return NumberParser._parse_american_format(str_value)
            else:
                return NumberParser._auto_parse_number(str_value)
        except:
            return None
    
    @staticmethod
    def _parse_european_format(str_value):
        """Parse European format: 1.234,56"""
        # Replace comma with dot for decimal
        if ',' in str_value and '.' in str_value:
            # Both present - comma should be decimal
            str_value = str_value.replace('.', '').replace(',', '.')
        elif ',' in str_value:
            # Only comma - could be thousands or decimal
            comma_pos = str_value.rfind(',')
            if len(str_value) - comma_pos - 1 <= 2:  # Decimal comma
                str_value = str_value.replace(',', '.')
            else:  # Thousands comma
                str_value = str_value.replace(',', '')
        
        return float(str_value)
    
    @staticmethod
    def _parse_american_format(str_value):
        """Parse American format: 1,234.56"""
        # Remove commas (thousands separators)
        str_value = str_value.replace(',', '')
        return float(str_value)
    
    @staticmethod
    def _auto_parse_number(str_value):
        """Auto-detect number format"""
        # If only one separator, determine its purpose
        if str_value.count(',') + str_value.count('.') == 1:
            if ',' in str_value:
                comma_pos = str_value.rfind(',')
                if len(str_value) - comma_pos - 1 <= 2:
                    # Likely decimal comma
                    return float(str_value.replace(',', '.'))
                else:
                    # Likely thousands comma
                    return float(str_value.replace(',', ''))
            else:
                # Only dot
                return float(str_value)
        
        # Multiple separators - assume American format
        return NumberParser._parse_american_format(str_value)

def get_month_name(month_num, language="id"):
    """
    Get month name in specified language
    
    Args:
        month_num (int): Month number (1-12)
        language (str): "id" for Indonesian, "en" for English
    
    Returns:
        str: Month name
    """
    if language == "id":
        months = [
            "Januari", "Februari", "Maret", "April", "Mei", "Juni",
            "Juli", "Agustus", "September", "Oktober", "November", "Desember"
        ]
    else:
        months = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]
    
    if 1 <= month_num <= 12:
        return months[month_num - 1]
    return f"Month{month_num}"

def safe_average(values):
    """
    Calculate average of numeric values, ignoring None/NaN
    
    Args:
        values: List of values
    
    Returns:
        float or None: Average value
    """
    numeric_values = [v for v in values if v is not None and not pd.isna(v)]
    if not numeric_values:
        return None
    return sum(numeric_values) / len(numeric_values)

def format_currency(value, currency="USD"):
    """
    Format number as currency
    
    Args:
        value: Numeric value
        currency: Currency code
    
    Returns:
        str: Formatted currency string
    """
    if value is None or pd.isna(value):
        return ""
    
    try:
        return f"{currency} {value:,.2f}"
    except:
        return str(value)

def format_american_number(value, decimals=2):
    """
    Format number using American format: comma as thousands separator, dot as decimal separator
    Replace zero values with "-" for better readability
    
    Args:
        value: Numeric value to format
        decimals: Number of decimal places (default: 2, or 'auto' for full precision)
    
    Returns:
        str: Formatted number string (e.g., "10,000.00" or "2.096666666666667") or "-" for zero values
    """
    if value is None or pd.isna(value):
        return "-"
    
    try:
        # Convert to float to ensure proper formatting
        numeric_value = float(value)
        
        # Return "-" for zero values
        if numeric_value == 0:
            return "-"
        
        # If decimals is 'auto', show full precision without rounding
        if decimals == 'auto':
            # Format with full precision, removing unnecessary trailing zeros
            formatted = f"{numeric_value:,}"
            return formatted
        
        # Format with specified decimal places
        if decimals > 0:
            return f"{numeric_value:,.{decimals}f}"
        else:
            return f"{numeric_value:,.0f}"
    except (ValueError, TypeError):
        return "-"

def format_price_with_precision(value, max_decimals=3):
    """
    Format price with controlled precision and proper rounding
    Always shows exactly 3 decimal places and replaces zero values with "-"
    
    Args:
        value: Numeric value to format
        max_decimals: Maximum decimal places to show (default: 3)
    
    Returns:
        str: Formatted price string with exactly 3 decimal places (e.g., "2.223") or "-" for zero values
    """
    if value is None or pd.isna(value):
        return "-"
    
    try:
        # Convert to float to ensure proper formatting
        numeric_value = float(value)
        
        # Return "-" for zero values
        if numeric_value == 0:
            return "-"
        
        # Round to specified decimal places
        rounded_value = round(numeric_value, max_decimals)
        
        # Format with thousands separator and exactly max_decimals decimal places
        return f"{rounded_value:,.{max_decimals}f}"
            
    except (ValueError, TypeError):
        return "-"

def format_qty_with_precision(value, max_decimals=3):
    """
    Format quantity with controlled precision and proper rounding
    Always shows exactly 3 decimal places for non-integer values and replaces zero values with "-"
    
    Args:
        value: Numeric value to format
        max_decimals: Maximum decimal places to show (default: 3)
    
    Returns:
        str: Formatted quantity string with proper rounding (e.g., "2,457.000" or "19,170") or "-" for zero values
    """
    if value is None or pd.isna(value):
        return "-"
    
    try:
        # Convert to float to ensure proper formatting
        numeric_value = float(value)
        
        # Return "-" for zero values
        if numeric_value == 0:
            return "-"
        
        # Round to specified decimal places
        rounded_value = round(numeric_value, max_decimals)
        
        # Check if it's a whole number after rounding
        if rounded_value == int(rounded_value):
            # If it's a whole number, show without decimals but with thousand separators
            return f"{int(rounded_value):,}"
        else:
            # Format with exactly max_decimals decimal places
            return f"{rounded_value:,.{max_decimals}f}"
            
    except (ValueError, TypeError):
        return "-"

def average_greater_than_zero(arr):
    """
    Calculate average of values greater than zero
    Matches the JavaScript averageGreaterThanZero function exactly
    
    Args:
        arr: List of numbers
        
    Returns:
        float: Average of values > 0, or 0 if no valid values
    """
    if not arr:
        return 0
    
    filtered_arr = [num for num in arr if isinstance(num, (int, float)) and num > 0]
    
    if len(filtered_arr) == 0:
        return 0
    
    return sum(filtered_arr) / len(filtered_arr)
