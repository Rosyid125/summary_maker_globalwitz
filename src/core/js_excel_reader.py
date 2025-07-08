"""
JavaScript-style Excel Reader
Implements the exact logic from the original JavaScript excelReader
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import re
from typing import List, Dict, Any, Optional, Union

from ..utils.constants import MONTH_ORDER, DEFAULT_INPUT_FOLDER, DEFAULT_SHEET_NAME
from ..utils.helpers import DateParser

class JSStyleExcelReader:
    """Excel reader with JavaScript-compatible logic"""
    
    def __init__(self, logger):
        self.logger = logger
    
    def parse_date_ddmmyyyy(self, date_string: str) -> Optional[datetime]:
        """Parse date in DD/MM/YYYY format (Indonesian standard)"""
        if not isinstance(date_string, str):
            return None
        
        # Format DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY (Indonesian standard)
        parts = re.match(r'(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})', date_string)
        if parts:
            day = int(parts.group(1))
            month = int(parts.group(2))
            year = int(parts.group(3))
            
            if year < 100:
                year += 1900 if year > 50 else 2000
            
            if 1 <= month <= 12 and 1 <= day <= 31:
                try:
                    date_obj = datetime(year, month, day)
                    if date_obj.year == year and date_obj.month == month and date_obj.day == day:
                        return date_obj
                except ValueError:
                    pass
        
        # Format YYYY-MM-DD (for Vietnam imports)
        parts = re.match(r'(\d{4})-(\d{2})-(\d{2})', date_string)
        if parts:
            year = int(parts.group(1))
            month = int(parts.group(2))
            day = int(parts.group(3))
            
            if 1 <= month <= 12 and 1 <= day <= 31:
                try:
                    date_obj = datetime(year, month, day)
                    if date_obj.year == year and date_obj.month == month and date_obj.day == day:
                        return date_obj
                except ValueError:
                    pass
        
        return None
    
    def parse_date_mmddyyyy(self, date_string: str) -> Optional[datetime]:
        """Parse date in MM/DD/YYYY format (US/Global standard)"""
        if not isinstance(date_string, str):
            return None
        
        # Format MM/DD/YYYY, MM-DD-YYYY, MM.DD.YYYY (US/Global standard)
        parts = re.match(r'(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})', date_string)
        if parts:
            month = int(parts.group(1))
            day = int(parts.group(2))
            year = int(parts.group(3))
            
            if year < 100:
                year += 1900 if year > 50 else 2000
            
            if 1 <= month <= 12 and 1 <= day <= 31:
                try:
                    date_obj = datetime(year, month, day)
                    if date_obj.year == year and date_obj.month == month and date_obj.day == day:
                        return date_obj
                except ValueError:
                    pass
        
        # Format YYYY-MM-DD (ISO standard)
        parts = re.match(r'(\d{4})-(\d{2})-(\d{2})', date_string)
        if parts:
            year = int(parts.group(1))
            month = int(parts.group(2))
            day = int(parts.group(3))
            
            if 1 <= month <= 12 and 1 <= day <= 31:
                try:
                    date_obj = datetime(year, month, day)
                    if date_obj.year == year and date_obj.month == month and date_obj.day == day:
                        return date_obj
                except ValueError:
                    pass
        
        return None
    
    def parse_date_ddmonthyyyy(self, date_string: str) -> Optional[datetime]:
        """Parse date in DD-Month-YYYY format"""
        if not isinstance(date_string, str):
            return None
        
        month_names = {
            'jan': 1, 'januari': 1,
            'feb': 2, 'februari': 2,
            'mar': 3, 'maret': 3,
            'apr': 4, 'april': 4,
            'mei': 5, 'may': 5,
            'jun': 6, 'juni': 6,
            'jul': 7, 'juli': 7,
            'agu': 8, 'agustus': 8, 'aug': 8, 'august': 8,
            'sep': 9, 'september': 9,
            'okt': 10, 'oktober': 10, 'oct': 10, 'october': 10,
            'nov': 11, 'november': 11,
            'des': 12, 'desember': 12, 'dec': 12, 'december': 12
        }
        
        parts = re.match(r'(\d{1,2})[\/\-\.]([a-zA-Z]+)[\/\-\.](\d{2,4})', date_string)
        if parts:
            day = int(parts.group(1))
            month_str = parts.group(2).lower()
            year = int(parts.group(3))
            
            if year < 100:
                year += 1900 if year > 50 else 2000
            
            month = month_names.get(month_str)
            if month and 1 <= day <= 31:
                try:
                    date_obj = datetime(year, month, day)
                    if date_obj.year == year and date_obj.month == month and date_obj.day == day:
                        return date_obj
                except ValueError:
                    pass
        
        return None
    
    def parse_date(self, date_string: str, date_format: str = 'DD/MM/YYYY') -> Optional[datetime]:
        """Parse date according to specified format"""
        if date_format == 'auto':
            # Try all formats in order of preference
            # First try DD-MONTH-YYYY (12-Apr-2025, 08-May-2025, etc.)
            result = self.parse_date_ddmonthyyyy(date_string)
            if result:
                return result
            
            # Then try DD/MM/YYYY
            result = self.parse_date_ddmmyyyy(date_string)
            if result:
                return result
            
            # Finally try MM/DD/YYYY
            result = self.parse_date_mmddyyyy(date_string)
            if result:
                return result
            
            return None
        elif date_format == 'MM/DD/YYYY':
            return self.parse_date_mmddyyyy(date_string)
        elif date_format == 'DD-MONTH-YYYY':
            return self.parse_date_ddmonthyyyy(date_string)
        else:
            return self.parse_date_ddmmyyyy(date_string)
    
    def excel_serial_number_to_date(self, serial: Union[int, float]) -> Optional[datetime]:
        """Convert Excel serial number to date"""
        if not isinstance(serial, (int, float)) or np.isnan(serial):
            return None
        
        # Check if serial is in reasonable range for Excel dates
        if serial < 1 or serial > 2958465:  # 2958465 is for 31/12/9999
            return None
        
        try:
            # Excel epoch adjustment
            utc_days = int(serial - 25569)
            utc_value = utc_days * 86400
            date_info = datetime.utcfromtimestamp(utc_value)
            
            # Handle fractional day
            fractional_day = serial - int(serial) + 0.0000001
            total_seconds = int(86400 * fractional_day)
            seconds = total_seconds % 60
            total_seconds -= seconds
            hours = total_seconds // 3600
            minutes = (total_seconds // 60) % 60
            
            return datetime(date_info.year, date_info.month, date_info.day, hours, minutes, seconds)
        except:
            return None
    
    def get_month_name(self, date_obj: datetime) -> str:
        """Get month name from date object"""
        if not date_obj or not isinstance(date_obj, datetime):
            return "N/A"
        
        try:
            return MONTH_ORDER[date_obj.month - 1]
        except (IndexError, AttributeError):
            return "N/A"
    
    def parse_number(self, value: Any, number_format: str = 'EUROPEAN') -> float:
        """Parse number according to format"""
        if isinstance(value, (int, float)):
            return 0 if np.isnan(value) else value
        
        if isinstance(value, str):
            cleaned_value = value.strip()
            if cleaned_value == "":
                return 0
            
            num_str = cleaned_value
            
            if number_format == 'AMERICAN':
                # American format: dot as decimal, comma as thousand separator
                american_regex = re.compile(r'^-?\d{1,3}(,\d{3})*(\.\d+)?$')
                if american_regex.match(num_str):
                    num_str = num_str.replace(',', '')
            else:
                # European format: comma as decimal, dot as thousand separator
                european_regex = re.compile(r'^-?\d{1,3}(\.\d{3})*(,\d+)?$')
                if european_regex.match(num_str):
                    num_str = num_str.replace('.', '').replace(',', '.')
            
            try:
                return float(num_str)
            except ValueError:
                return 0
        
        return 0
    
    def read_and_preprocess_data(self, input_file_path: str, sheet_name: str = DEFAULT_SHEET_NAME, 
                               date_format: str = 'DD/MM/YYYY', number_format: str = 'EUROPEAN',
                               column_mapping: Dict[str, str] = None) -> Optional[List[Dict[str, Any]]]:
        """
        Read and preprocess Excel data exactly like JavaScript version
        
        Args:
            input_file_path: Path to input Excel file
            sheet_name: Name of sheet to process
            date_format: Date format to use
            number_format: Number format to use
            column_mapping: Column mapping dictionary
            
        Returns:
            List of dictionaries with processed data
        """
        if not os.path.exists(input_file_path):
            self.logger.error(f"Error: Input file '{input_file_path}' not found.")
            return None
        
        try:
            # Read Excel file
            excel_file = pd.ExcelFile(input_file_path)
            
            if sheet_name not in excel_file.sheet_names:
                self.logger.error(f"Error: Sheet '{sheet_name}' not found in file {input_file_path}")
                return None
            
            # Read sheet
            df = pd.read_excel(input_file_path, sheet_name=sheet_name)
            
            self.logger.info(f"Reading {len(df)} rows from sheet '{sheet_name}' with date format {date_format} and number format {number_format}...")
            
            # Helper function to get column value with mapping
            def get_column_value(row, mapping_key, default_columns):
                if column_mapping and column_mapping.get(mapping_key):
                    mapped_col = column_mapping[mapping_key]
                    if mapped_col in row:
                        return row[mapped_col]
                
                # Fallback to default columns (case-insensitive matching)
                row_columns = {col.lower(): col for col in row.index}
                for col in default_columns:
                    # Try exact match first
                    if col in row:
                        return row[col]
                    # Try case-insensitive match
                    col_lower = col.lower()
                    if col_lower in row_columns:
                        actual_col = row_columns[col_lower]
                        return row[actual_col]
                return None
            
            def safe_string_value(value):
                """Convert value to string safely, handling NaN and None"""
                import pandas as pd
                import numpy as np
                
                if value is None:
                    return "-"
                if pd.isna(value) or (isinstance(value, float) and np.isnan(value)):
                    return "-"
                if str(value).strip() == "":
                    return "-"
                return str(value).strip()
            
            # Process each row
            processed_data = []
            for index, row in df.iterrows():
                # Process date
                date_value = get_column_value(row, 'date', ["Arrival Date", "DATE", "CUSTOMS CLEARANCE DATE"])
                month = "-"
                
                if date_value is not None:
                    parsed_date = None
                    
                    # Try parsing as Excel serial number first
                    if isinstance(date_value, (int, float)):
                        parsed_date = self.excel_serial_number_to_date(date_value)
                    elif isinstance(date_value, str):
                        parsed_date = self.parse_date(date_value, date_format)
                    elif isinstance(date_value, datetime):
                        parsed_date = date_value
                    
                    if parsed_date:
                        month = self.get_month_name(parsed_date)
                
                # Process other fields - support both GUI mapping keys and original keys
                hs_code = safe_string_value(get_column_value(row, 'hs_code', ["HS Code", "HS CODE"]) or
                          get_column_value(row, 'hsCode', ["HS Code", "HS CODE"]))
                item_desc = safe_string_value(get_column_value(row, 'item_description', ["Product Description", "ITEM DESC", "PRODUCT DESCRIPTION(EN)"]) or
                            get_column_value(row, 'itemDesc', ["Product Description", "ITEM DESC", "PRODUCT DESCRIPTION(EN)"]))
                gsm = safe_string_value(get_column_value(row, 'gsm', ["GSM"]))
                item = safe_string_value(get_column_value(row, 'item', ["ITEM"]))
                add_on = safe_string_value(get_column_value(row, 'add_on', ["ADD ON"]) or
                         get_column_value(row, 'addOn', ["ADD ON"]))
                importer = safe_string_value(get_column_value(row, 'importer', ["Consignee Name", "IMPORTER", "PURCHASER"]))
                supplier = safe_string_value(get_column_value(row, 'supplier', ["Shipper Name", "SUPPLIER"]))
                origin_country = safe_string_value(get_column_value(row, 'origin_country', ["Country of Origin", "ORIGIN COUNTRY"]) or
                                 get_column_value(row, 'originCountry', ["Country of Origin", "ORIGIN COUNTRY"]))
                incoterms = safe_string_value(get_column_value(row, 'incoterms', ["INCOTERMS", "Incoterms", "INCOTERM", "Incoterm"]))
                
                # Process numeric fields - improved price column detection
                # Support both GUI mapping keys (unit_price/quantity) and original keys (unitPrice/quantity)
                unit_price_raw = (get_column_value(row, 'unit_price', ["Standard Unit Rate $", "Value CIF US$", "CIF KG Unit In USD", "USD Qty Unit", "UNIT PRICE(USD)"]) or
                                 get_column_value(row, 'unitPrice', ["Standard Unit Rate $", "Value CIF US$", "CIF KG Unit In USD", "USD Qty Unit", "UNIT PRICE(USD)"]))
                quantity_raw = (get_column_value(row, 'quantity', ["Standard Qty", "Std. Quantity", "Net KG Wt", "qty", "BUSINESS QUANTITY (KG)"]) or
                               get_column_value(row, 'quantity', ["Standard Qty", "Std. Quantity", "Net KG Wt", "qty", "BUSINESS QUANTITY (KG)"]))
                
                usd_qty_unit = self.parse_number(unit_price_raw, number_format)
                qty = self.parse_number(quantity_raw, number_format)
                
                # Log price processing for first few rows to help debugging
                if index < 5:
                    self.logger.info(f"Row {index+1}: unit_price_raw='{unit_price_raw}' -> parsed={usd_qty_unit}, qty_raw='{quantity_raw}' -> parsed={qty}")
                
                processed_row = {
                    'month': month,
                    'hsCode': hs_code,
                    'itemDesc': item_desc,
                    'gsm': gsm,
                    'item': item,
                    'addOn': add_on,
                    'importer': importer,
                    'supplier': supplier,
                    'originCountry': origin_country,
                    'incoterms': incoterms,
                    'usdQtyUnit': usd_qty_unit,
                    'qty': qty
                }
                
                processed_data.append(processed_row)
            
            self.logger.info(f"Processed {len(processed_data)} rows successfully")
            return processed_data
            
        except Exception as error:
            self.logger.error(f"Error reading Excel file '{input_file_path}': {str(error)}")
            return None
    
    def get_excel_info(self, input_file_path: str) -> Optional[Dict[str, Any]]:
        """Get Excel file information"""
        if not os.path.exists(input_file_path):
            self.logger.error(f"Error: Input file '{input_file_path}' not found.")
            return None
        
        try:
            excel_file = pd.ExcelFile(input_file_path)
            sheet_names = excel_file.sheet_names
            
            # Get column names from first sheet
            column_names = []
            if sheet_names:
                df = pd.read_excel(input_file_path, sheet_name=sheet_names[0], nrows=0)
                column_names = df.columns.tolist()
            
            return {
                'sheetNames': sheet_names,
                'columnNames': column_names
            }
            
        except Exception as error:
            self.logger.error(f"Error reading Excel structure '{input_file_path}': {str(error)}")
            return None
    
    def get_sheet_column_names(self, input_file_path: str, sheet_name: str) -> List[str]:
        """Get column names from specific sheet"""
        if not os.path.exists(input_file_path):
            return []
        
        try:
            excel_file = pd.ExcelFile(input_file_path)
            if sheet_name not in excel_file.sheet_names:
                return []
            
            df = pd.read_excel(input_file_path, sheet_name=sheet_name, nrows=0)
            return df.columns.tolist()
            
        except Exception as error:
            self.logger.error(f"Error reading columns from sheet '{sheet_name}': {str(error)}")
            return []
    
    def scan_excel_files(self, input_folder: str = DEFAULT_INPUT_FOLDER) -> List[Dict[str, Any]]:
        """Scan and list all Excel files in input folder"""
        try:
            folder_path = os.path.abspath(input_folder)
            
            if not os.path.exists(folder_path):
                self.logger.warning(f"Folder '{folder_path}' does not exist")
                return []
            
            files = os.listdir(folder_path)
            excel_files = [f for f in files if f.lower().endswith(('.xlsx', '.xls'))]
            
            result = []
            for file in excel_files:
                file_path = os.path.join(folder_path, file)
                file_stat = os.stat(file_path)
                
                result.append({
                    'name': file,
                    'path': file_path,
                    'size': f"{file_stat.st_size / 1024:.1f} KB",
                    'modified': datetime.fromtimestamp(file_stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                })
            
            return result
            
        except Exception as error:
            self.logger.error(f"Error scanning folder {input_folder}: {str(error)}")
            return []
        
    def find_best_price_column(self, df: pd.DataFrame, column_mapping: Dict[str, str] = None) -> Optional[str]:
        """
        Find the best column that contains actual price data (non-zero values)
        
        Args:
            df: DataFrame to analyze
            column_mapping: User-specified column mapping
            
        Returns:
            str: Name of the best price column, or None if not found
        """
        # If user specified a price column, check if it has data
        if column_mapping and column_mapping.get('unitPrice'):
            user_col = column_mapping['unitPrice']
            if user_col in df.columns:
                non_zero_count = (df[user_col].dropna() != 0).sum()
                if non_zero_count > 0:
                    self.logger.info(f"Using user-specified price column '{user_col}' with {non_zero_count} non-zero values")
                    return user_col
                else:
                    self.logger.warning(f"User-specified price column '{user_col}' contains only zeros!")
        
        # List of potential price columns to check
        potential_price_columns = [
            "Value CIF US$", "CIF KG Unit In USD", "USD Qty Unit", "UNIT PRICE(USD)",
            "Std. Unit Rate $", "Unit Rate", "Price", "Value", "Amount",
            "CIF Value", "FOB Value", "Total Value", "Unit Cost"
        ]
        
        best_column = None
        best_score = 0
        
        for col_name in potential_price_columns:
            if col_name in df.columns:
                try:
                    # Convert to numeric and count non-zero, non-null values
                    numeric_col = pd.to_numeric(df[col_name], errors='coerce')
                    non_zero_count = (numeric_col.dropna() != 0).sum()
                    total_count = len(numeric_col.dropna())
                    
                    if non_zero_count > 0:
                        # Score based on percentage of non-zero values
                        score = non_zero_count / len(df) if len(df) > 0 else 0
                        self.logger.info(f"Price column '{col_name}': {non_zero_count}/{total_count} non-zero values (score: {score:.2f})")
                        
                        if score > best_score:
                            best_score = score
                            best_column = col_name
                    else:
                        self.logger.debug(f"Price column '{col_name}' contains only zeros")
                except Exception as e:
                    self.logger.debug(f"Could not analyze column '{col_name}': {e}")
        
        if best_column:
            self.logger.info(f"Selected best price column: '{best_column}' with score {best_score:.2f}")
        else:
            self.logger.warning("No price column with non-zero values found!")
            
        return best_column

    def calculate_unit_price_from_total(self, df: pd.DataFrame, total_value_col: str, quantity_col: str) -> pd.Series:
        """
        Calculate unit price by dividing total value by quantity
        
        Args:
            df: DataFrame
            total_value_col: Column containing total value
            quantity_col: Column containing quantity
            
        Returns:
            pd.Series: Calculated unit prices
        """
        try:
            total_values = pd.to_numeric(df[total_value_col], errors='coerce')
            quantities = pd.to_numeric(df[quantity_col], errors='coerce')
            
            # Avoid division by zero
            unit_prices = total_values / quantities.replace(0, pd.NA)
            
            valid_prices = unit_prices.dropna()
            self.logger.info(f"Calculated {len(valid_prices)} unit prices from {total_value_col}/{quantity_col}")
            
            return unit_prices.fillna(0)
            
        except Exception as e:
            self.logger.error(f"Error calculating unit price: {e}")
            return pd.Series([0] * len(df))
