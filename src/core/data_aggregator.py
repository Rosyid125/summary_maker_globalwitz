"""
Data Aggregator Module - Handles data aggregation and summarization
Implements the exact logic from the original JavaScript aggregator
"""

import pandas as pd
from typing import Dict, List, Any, Optional
from collections import defaultdict
from datetime import datetime

from ..utils.helpers import safe_average, get_month_name, average_greater_than_zero

class DataAggregator:
    """Handles data aggregation and summarization with JavaScript-compatible logic"""
    
    def __init__(self, logger):
        self.logger = logger
    
    def _safe_string_value(self, value):
        """Convert value to string safely, handling NaN and None"""
        if value is None:
            return ""
        if pd.isna(value) or (isinstance(value, float) and pd.isna(value)):
            return ""
        if str(value).strip() == "":
            return ""
        return str(value).strip()
    def perform_aggregation(self, data: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Perform aggregation exactly like the JavaScript version
        
        Args:
            data: List of dictionaries with raw data
            
        Returns:
            Dict with 'summaryLvl1' and 'summaryLvl2' keys
        """
        try:
            monthly_summary = {}
            
            self.logger.info(f"    Starting aggregation for {len(data)} rows")
            
            if not data:
                self.logger.warning("    No data to aggregate")
                return {'summaryLvl1': [], 'summaryLvl2': []}
            
            # Debug: Show sample data being processed
            for i, row in enumerate(data[:3]):
                self.logger.info(f"    Sample row {i}: month='{row.get('month')}', hsCode='{row.get('hsCode')}', date='{row.get('date')}'")
            
            valid_rows_processed = 0
            
            for index, row in enumerate(data):
                # Required columns: month, hsCode
                # gsm, item, addOn can be '-' or empty string and are valid for grouping
                if not row.get('month') or row.get('month') == "-" or not row.get('hsCode') or row.get('hsCode') == "-":
                    self.logger.debug(f"    Skipping row {index}: month='{row.get('month')}', hsCode='{row.get('hsCode')}'")
                    continue
                
                # If GSM, ITEM, or ADD ON don't exist (None/undefined), treat as empty string for grouping
                # If value is string "-", it will be treated as unique "-" value
                gsm_value = self._safe_string_value(row.get('gsm'))
                item_value = self._safe_string_value(row.get('item'))
                add_on_value = self._safe_string_value(row.get('addOn'))
                
                key = f"{row['month']}-{row['hsCode']}-{item_value}-{gsm_value}-{add_on_value}"
                
                if key not in monthly_summary:
                    monthly_summary[key] = {
                        'month': row['month'],
                        'hsCode': self._safe_string_value(row['hsCode']),
                        'item': item_value,
                        'gsm': gsm_value,
                        'addOn': add_on_value,
                        'usdQtyUnits': [],
                        'totalQty': 0
                    }
                
                usd_qty = row.get('usdQtyUnit', 0)  # Fixed field name to match Excel reader
                qty = row.get('qty', 0)
                
                # Ensure numeric values
                try:
                    usd_qty = float(usd_qty) if usd_qty is not None else 0
                except (ValueError, TypeError):
                    usd_qty = 0
                    
                try:
                    qty = float(qty) if qty is not None else 0
                except (ValueError, TypeError):
                    qty = 0
                
                # Debug: Log price values for first few rows
                if index < 5:
                    self.logger.info(f"    Row {index}: usdQtyUnit={usd_qty}, qty={qty}")
                
                if usd_qty > 0:  # Only add positive prices
                    monthly_summary[key]['usdQtyUnits'].append(usd_qty)
                monthly_summary[key]['totalQty'] += qty
                valid_rows_processed += 1
            
            self.logger.info(f"    Processed {valid_rows_processed} valid rows out of {len(data)} total rows")
            self.logger.info(f"    Created {len(monthly_summary)} monthly groups")
            
            # Create summaryLvl1Data
            summary_lvl1_data = []
            for group in monthly_summary.values():
                summary_lvl1_data.append({
                    'month': group['month'],
                    'hsCode': self._safe_string_value(group['hsCode']),
                    'item': self._safe_string_value(group['item']),
                    'gsm': self._safe_string_value(group['gsm']),
                    'addOn': self._safe_string_value(group['addOn']),
                    'avgPrice': average_greater_than_zero(group['usdQtyUnits']),
                    'totalQty': group['totalQty']
                })
            
            # Create recapSummary
            recap_summary = {}
            for row in summary_lvl1_data:
                key = f"{row['hsCode']}-{row['item']}-{row['gsm']}-{row['addOn']}"
                if key not in recap_summary:
                    recap_summary[key] = {
                        'hsCode': self._safe_string_value(row['hsCode']),
                        'item': self._safe_string_value(row['item']),
                        'gsm': self._safe_string_value(row['gsm']),
                        'addOn': self._safe_string_value(row['addOn']),
                        'avgPrices': [],
                        'totalQty': 0
                    }
                if row['avgPrice'] and row['avgPrice'] > 0:
                    recap_summary[key]['avgPrices'].append(row['avgPrice'])
                recap_summary[key]['totalQty'] += row['totalQty']
            
            # Create summaryLvl2Data
            summary_lvl2_data = []
            for group in recap_summary.values():
                summary_lvl2_data.append({
                    'hsCode': self._safe_string_value(group['hsCode']),
                    'item': self._safe_string_value(group['item']), 
                    'gsm': self._safe_string_value(group['gsm']),
                    'addOn': self._safe_string_value(group['addOn']),
                    'avgOfSummaryPrice': average_greater_than_zero(group['avgPrices']),
                    'totalOfSummaryQty': group['totalQty']
                })
            
            self.logger.info(f"    Final result: Level1={len(summary_lvl1_data)}, Level2={len(summary_lvl2_data)}")
            
            return {
                'summaryLvl1': summary_lvl1_data,
                'summaryLvl2': summary_lvl2_data
            }
            
        except Exception as e:
            self.logger.error(f"    Error in perform_aggregation: {str(e)}")
            return {'summaryLvl1': [], 'summaryLvl2': []}

    def aggregate_data(self, df: pd.DataFrame, year: int = None) -> Dict[str, Any]:
        """
        Aggregate data by importer and create summary tables
        This method maintains compatibility with the existing GUI code
        
        Args:
            df (pd.DataFrame): Input data
            year (int): Target year for filtering
            
        Returns:
            Dict: Aggregated data structure
        """
        try:
            if df.empty:
                self.logger.warning("Empty DataFrame provided for aggregation")
                return {}
            
            # Filter by year if specified
            if year:
                df = self._filter_by_year(df, year)
                if df.empty:
                    self.logger.warning(f"No data found for year {year}")
                    return {}
            
            # Group by importer
            importer_data = {}
            
            if 'importer' in df.columns:
                for importer in df['importer'].dropna().unique():
                    importer_df = df[df['importer'] == importer].copy()
                    importer_data[importer] = self._aggregate_importer_data(importer_df)
            else:
                # No importer column, treat all data as single group
                importer_data['All Data'] = self._aggregate_importer_data(df)
            
            self.logger.info(f"Aggregated data for {len(importer_data)} importers")
            
            return importer_data
            
        except Exception as e:
            self.logger.error(f"Error during data aggregation: {str(e)}")
            return {}
    
    def _filter_by_year(self, df: pd.DataFrame, year: int) -> pd.DataFrame:
        """Filter DataFrame by year"""
        if 'date' not in df.columns:
            return df
        
        # Filter rows where date year matches target year
        date_mask = df['date'].apply(
            lambda x: x.year == year if isinstance(x, datetime) else False
        )
        
        return df[date_mask].copy()
    
    def _aggregate_importer_data(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Aggregate data for a single importer
        
        Args:
            df (pd.DataFrame): Data for single importer
            
        Returns:
            Dict: Aggregated data structure
        """
        result = {
            'monthly_summary': self._create_monthly_summary(df),
            'overall_summary': self._create_overall_summary(df),
            'supplier_summary': self._create_supplier_summary(df),
            'item_summary': self._create_item_summary(df),
            'raw_data': df.to_dict('records')
        }
        
        return result
    
    def _create_monthly_summary(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """Create monthly summary grouped by HS Code, Item, GSM, Add On"""
        try:
            if 'date' not in df.columns:
                return []
            
            # Add month-year column
            df_copy = df.copy()
            df_copy['month_year'] = df_copy['date'].apply(
                lambda x: f"{x.year}-{x.month:02d}" if isinstance(x, datetime) else None
            )
            
            # Group by key fields and month
            group_columns = ['month_year']
            for col in ['hs_code', 'item', 'gsm', 'add_on']:
                if col in df_copy.columns:
                    group_columns.append(col)
            
            # Remove None values from group columns
            df_clean = df_copy.dropna(subset=[col for col in group_columns if col != 'month_year'])
            
            if df_clean.empty:
                return []
            
            monthly_groups = df_clean.groupby(group_columns, dropna=False)
            
            monthly_summary = []
            for group_key, group_data in monthly_groups:
                if isinstance(group_key, tuple):
                    month_year = group_key[0]
                    other_keys = group_key[1:]
                else:
                    month_year = group_key
                    other_keys = []
                
                if not month_year:
                    continue
                
                summary_row = {
                    'month_year': month_year,
                    'month_name': self._get_month_name_from_key(month_year),
                    'total_quantity': group_data['quantity'].sum() if 'quantity' in group_data else 0,
                    'avg_unit_price': safe_average(group_data['unit_price'].dropna()) if 'unit_price' in group_data else None,
                    'total_value': group_data['unit_price'].sum() * group_data['quantity'].sum() if all(col in group_data for col in ['unit_price', 'quantity']) else 0,
                    'record_count': len(group_data)
                }
                
                # Add grouped field values
                for i, col in enumerate(['hs_code', 'item', 'gsm', 'add_on']):
                    if i < len(other_keys):
                        summary_row[col] = other_keys[i]
                    elif col in df_copy.columns:
                        # Use most common value in group
                        mode_val = group_data[col].mode()
                        summary_row[col] = mode_val.iloc[0] if not mode_val.empty else None
                
                monthly_summary.append(summary_row)
            
            # Sort by month_year
            monthly_summary.sort(key=lambda x: x['month_year'] or '')
            
            return monthly_summary
            
        except Exception as e:
            self.logger.error(f"Error creating monthly summary: {str(e)}")
            return []
    
    def _create_overall_summary(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Create overall summary across all months"""
        try:
            summary = {
                'total_records': len(df),
                'total_quantity': df['quantity'].sum() if 'quantity' in df.columns else 0,
                'avg_unit_price': safe_average(df['unit_price'].dropna()) if 'unit_price' in df.columns else None,
                'total_value': 0,
                'date_range': self._get_date_range(df),
                'unique_suppliers': len(df['supplier'].dropna().unique()) if 'supplier' in df.columns else 0,
                'unique_items': len(df['item'].dropna().unique()) if 'item' in df.columns else 0,
                'unique_hs_codes': len(df['hs_code'].dropna().unique()) if 'hs_code' in df.columns else 0
            }
            
            # Calculate total value
            if all(col in df.columns for col in ['unit_price', 'quantity']):
                df_clean = df.dropna(subset=['unit_price', 'quantity'])
                summary['total_value'] = (df_clean['unit_price'] * df_clean['quantity']).sum()
            
            return summary
            
        except Exception as e:
            self.logger.error(f"Error creating overall summary: {str(e)}")
            return {}
    
    def _create_supplier_summary(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """Create summary grouped by supplier"""
        try:
            if 'supplier' not in df.columns:
                return []
            
            supplier_groups = df.groupby('supplier', dropna=False)
            supplier_summary = []
            
            for supplier, group_data in supplier_groups:
                if pd.isna(supplier):
                    continue
                
                summary_row = {
                    'supplier': supplier,
                    'total_quantity': group_data['quantity'].sum() if 'quantity' in group_data else 0,
                    'avg_unit_price': safe_average(group_data['unit_price'].dropna()) if 'unit_price' in group_data else None,
                    'total_value': 0,
                    'record_count': len(group_data),
                    'unique_items': len(group_data['item'].dropna().unique()) if 'item' in group_data else 0,
                    'date_range': self._get_date_range(group_data)
                }
                
                # Calculate total value
                if all(col in group_data.columns for col in ['unit_price', 'quantity']):
                    clean_data = group_data.dropna(subset=['unit_price', 'quantity'])
                    summary_row['total_value'] = (clean_data['unit_price'] * clean_data['quantity']).sum()
                
                supplier_summary.append(summary_row)
            
            # Sort by total value descending
            supplier_summary.sort(key=lambda x: x['total_value'], reverse=True)
            
            return supplier_summary
            
        except Exception as e:
            self.logger.error(f"Error creating supplier summary: {str(e)}")
            return []
    
    def _create_item_summary(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """Create summary grouped by item"""
        try:
            if 'item' not in df.columns:
                return []
            
            item_groups = df.groupby('item', dropna=False)
            item_summary = []
            
            for item, group_data in item_groups:
                if pd.isna(item):
                    continue
                
                summary_row = {
                    'item': item,
                    'total_quantity': group_data['quantity'].sum() if 'quantity' in group_data else 0,
                    'avg_unit_price': safe_average(group_data['unit_price'].dropna()) if 'unit_price' in group_data else None,
                    'total_value': 0,
                    'record_count': len(group_data),
                    'unique_suppliers': len(group_data['supplier'].dropna().unique()) if 'supplier' in group_data else 0,
                    'avg_gsm': safe_average(group_data['gsm'].dropna()) if 'gsm' in group_data else None
                }
                
                # Calculate total value
                if all(col in group_data.columns for col in ['unit_price', 'quantity']):
                    clean_data = group_data.dropna(subset=['unit_price', 'quantity'])
                    summary_row['total_value'] = (clean_data['unit_price'] * clean_data['quantity']).sum()
                
                item_summary.append(summary_row)
            
            # Sort by total quantity descending
            item_summary.sort(key=lambda x: x['total_quantity'], reverse=True)
            
            return item_summary
            
        except Exception as e:
            self.logger.error(f"Error creating item summary: {str(e)}")
            return []
    
    def _get_date_range(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Get date range from DataFrame"""
        try:
            if 'date' not in df.columns:
                return {'start': None, 'end': None, 'span_days': 0}
            
            dates = df['date'].dropna()
            if dates.empty:
                return {'start': None, 'end': None, 'span_days': 0}
            
            min_date = dates.min()
            max_date = dates.max()
            span_days = (max_date - min_date).days if isinstance(min_date, datetime) and isinstance(max_date, datetime) else 0
            
            return {
                'start': min_date,
                'end': max_date,
                'span_days': span_days
            }
            
        except Exception as e:
            self.logger.error(f"Error calculating date range: {str(e)}")
            return {'start': None, 'end': None, 'span_days': 0}
    
    def _get_month_name_from_key(self, month_year_key: str) -> str:
        """Extract month name from month_year key"""
        try:
            if not month_year_key or '-' not in month_year_key:
                return month_year_key or ''
            
            year, month = month_year_key.split('-')
            month_num = int(month)
            return f"{get_month_name(month_num, 'id')} {year}"
            
        except Exception:
            return month_year_key or ''
