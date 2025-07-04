"""
JavaScript-style data processor
Implements the exact logic from the original JavaScript index.js
"""

import pandas as pd
from typing import Dict, List, Any, Optional
import os

from .data_aggregator import DataAggregator
from .js_output_formatter import OutputFormatter
from ..utils.constants import MONTH_ORDER, DEFAULT_OUTPUT_FOLDER
from ..utils.helpers import average_greater_than_zero

class JSStyleProcessor:
    """Processes data using JavaScript-compatible logic"""
    
    def __init__(self, logger):
        self.logger = logger
        self.aggregator = DataAggregator(logger)
        self.formatter = OutputFormatter(logger)
    
    def process_sheet_data(self, data_to_process: List[Dict], sheet_base_name: str, 
                          incoterm_value: str, incoterm_mode: str = "manual") -> Optional[Dict[str, Any]]:
        """
        Process sheet data exactly like JavaScript processSheetData function
        
        Args:
            data_to_process: List of data dictionaries
            sheet_base_name: Base name for the sheet
            incoterm_value: INCOTERM value to use (for manual mode)
            incoterm_mode: Mode for incoterm handling ("manual" or "from_column")
            
        Returns:
            Dict with sheet data or None if no data
        """
        self.logger.info(f"Processing data for sheet based on '{sheet_base_name}' with INCOTERM: {incoterm_value}, mode: {incoterm_mode}...")
        
        # Group by supplier or origin
        grouped_by_supplier_or_origin = {}
        for row in data_to_process:
            # Try supplier first, then origin_country, then "Unknown"
            group_key = row.get('supplier') or row.get('originCountry') or "Unknown"
            if group_key not in grouped_by_supplier_or_origin:
                grouped_by_supplier_or_origin[group_key] = []
            grouped_by_supplier_or_origin[group_key].append(row)
        
        all_rows_for_sheet_content = []
        supplier_groups_meta = []
        sheet_overall_monthly_totals = [0] * 12
        item_summary_data_for_sheet = {}
        
        total_columns = 5 + len(MONTH_ORDER) * 2 + 3
        
        # Process each supplier group
        group_keys = sorted(grouped_by_supplier_or_origin.keys())
        for group_index, group_name in enumerate(group_keys):
            self.logger.info(f"  - Processing supplier/origin group: {group_name}")
            group_data = grouped_by_supplier_or_origin[group_name]
            
            # Perform aggregation
            aggregation_result = self.aggregator.perform_aggregation(group_data)
            summary_lvl1 = aggregation_result['summaryLvl1']
            summary_lvl2 = aggregation_result['summaryLvl2']
            
            self.logger.info(f"    Aggregation result: Level1={len(summary_lvl1)}, Level2={len(summary_lvl2)}")
            
            if summary_lvl2:
                # Prepare group block
                group_block = self.formatter.prepare_group_block(group_name, summary_lvl1, summary_lvl2, 
                                                               incoterm_value, incoterm_mode, group_data)
                all_rows_for_sheet_content.extend(group_block['groupBlockRows'])
                
                supplier_groups_meta.append({
                    'name': group_name,
                    'productRowCount': group_block['distinctCombinationsCount'],
                    'headerRowCount': group_block['headerRowCount'],
                    'hasFollowingGroup': group_index < len(group_keys) - 1
                })
                
                # Update sheet overall monthly totals and item summary
                for lvl1_row in summary_lvl1:
                    try:
                        month_index = MONTH_ORDER.index(lvl1_row['month'])
                        qty_to_add = lvl1_row['totalQty'] if isinstance(lvl1_row['totalQty'], (int, float)) else 0
                        sheet_overall_monthly_totals[month_index] += qty_to_add
                        
                        # Update item summary
                        item_key = f"{lvl1_row['item']}-{lvl1_row['gsm']}-{lvl1_row['addOn']}"
                        if item_key not in item_summary_data_for_sheet:
                            item_summary_data_for_sheet[item_key] = {
                                'item': lvl1_row['item'],
                                'gsm': lvl1_row['gsm'],
                                'addOn': lvl1_row['addOn'],
                                'monthlyQtys': [0] * 12,
                                'totalQtyRecap': 0
                            }
                        
                        item_summary_data_for_sheet[item_key]['monthlyQtys'][month_index] += qty_to_add
                        item_summary_data_for_sheet[item_key]['totalQtyRecap'] += qty_to_add
                        
                    except ValueError:
                        # Month not in MONTH_ORDER
                        continue
                
                # Add separator if not last group
                if group_index < len(group_keys) - 1:
                    all_rows_for_sheet_content.append([])
        
        if all_rows_for_sheet_content:
            # Add separator
            all_rows_for_sheet_content.append([])
            
            # Add "TOTAL ALL SUPPLIER" section
            total_all_header_month_row = ["Month", None, None, None, None]
            for month in MONTH_ORDER:
                total_all_header_month_row.extend([month, None])
            total_all_header_month_row.extend(["RECAP", None, None])
            all_rows_for_sheet_content.append(total_all_header_month_row)
            
            # Calculate grand total
            grand_total_all_suppliers = sum(sheet_overall_monthly_totals)
            total_all_mo_row = ["TOTAL ALL SUPPLIER PER MO", None, None, None, None]
            for total in sheet_overall_monthly_totals:
                total_all_mo_row.extend([total, None])
            total_all_mo_row.extend([grand_total_all_suppliers, None, None])
            all_rows_for_sheet_content.append(total_all_mo_row)
            
            # Calculate quarterly totals
            quarterly_totals_all = [0, 0, 0, 0]
            for i, total in enumerate(sheet_overall_monthly_totals):
                if i < 3:
                    quarterly_totals_all[0] += total
                elif i < 6:
                    quarterly_totals_all[1] += total
                elif i < 9:
                    quarterly_totals_all[2] += total
                else:
                    quarterly_totals_all[3] += total
            
            total_all_quartal_row = ["TOTAL ALL SUPPLIER PER QUARTAL", None, None, None, None]
            total_all_quartal_row.extend([quarterly_totals_all[0], None, None, None, None, None])
            total_all_quartal_row.extend([quarterly_totals_all[1], None, None, None, None, None])
            total_all_quartal_row.extend([quarterly_totals_all[2], None, None, None, None, None])
            total_all_quartal_row.extend([quarterly_totals_all[3], None, None, None, None, None])
            total_all_quartal_row.extend([None, None, None])
            all_rows_for_sheet_content.append(total_all_quartal_row)
            
            # Add separator
            all_rows_for_sheet_content.append([])
            
            # Add "TOTAL PER ITEM" section
            item_table_main_title_row = ["TOTAL PER ITEM"]
            all_rows_for_sheet_content.append(item_table_main_title_row)
            
            item_table_header_month_row = ["Month", None, None, None, None]
            for month in MONTH_ORDER:
                item_table_header_month_row.extend([month, None])
            item_table_header_month_row.extend(["RECAP", None, None])
            all_rows_for_sheet_content.append(item_table_header_month_row)
            
            # Add item rows
            for item_key in sorted(item_summary_data_for_sheet.keys()):
                item_data = item_summary_data_for_sheet[item_key]
                item_row = [f"{item_data['item']} {item_data['gsm']} {item_data['addOn']}", None, None, None, None]
                for qty in item_data['monthlyQtys']:
                    item_row.extend([qty, None])
                item_row.extend([item_data['totalQtyRecap'], None, None])
                all_rows_for_sheet_content.append(item_row)
            
            return {
                'name': sheet_base_name,
                'allRowsForSheetContent': all_rows_for_sheet_content,
                'supplierGroupsMeta': supplier_groups_meta,
                'totalColumns': total_columns
            }
        
        return None
    
    def process_data_like_javascript(self, all_raw_data: List[Dict], period_year: str, 
                                   global_incoterm: str, incoterm_mode: str = "manual",
                                   output_filename: str = "summary_output.xlsx") -> str:
        """
        Process all data like the JavaScript main function
        
        Args:
            all_raw_data: All raw data
            period_year: Year for the period
            global_incoterm: Global INCOTERM value (for manual mode)
            incoterm_mode: Mode for incoterm handling ("manual" or "from_column")
            output_filename: Output filename
            
        Returns:
            str: Path to output file
        """
        workbook_data_for_excel_js = []
        
        # Separate data with valid importer vs blank/NA importer
        data_with_valid_importer = []
        data_with_blank_or_na_importer = []
        
        for row in all_raw_data:
            importer = row.get('importer')
            if not importer or importer == "N/A" or importer == "":
                data_with_blank_or_na_importer.append(row)
            else:
                data_with_valid_importer.append(row)
        
        # Process data without importer
        if data_with_blank_or_na_importer:
            sheet_name_for_blank = "Data_Tanpa_Importer"
            sheet_result = self.process_sheet_data(data_with_blank_or_na_importer, sheet_name_for_blank, 
                                                 global_incoterm, incoterm_mode)
            if sheet_result:
                workbook_data_for_excel_js.append(sheet_result)
        
        # Process data by importer
        if data_with_valid_importer:
            # Get unique importers
            unique_importers = list(set(row['importer'] for row in data_with_valid_importer if row.get('importer')))
            unique_importers.sort()
            
            for importer in unique_importers:
                importer_data = [row for row in data_with_valid_importer if row.get('importer') == importer]
                if importer_data:
                    # Clean sheet name (replace invalid characters)
                    base_sheet_name = importer.replace('*', '_').replace('?', '_').replace(':', '_').replace('\\', '_').replace('/', '_').replace('[', '_').replace(']', '_')
                    base_sheet_name = base_sheet_name[:30]  # Limit to 30 characters
                    
                    sheet_result = self.process_sheet_data(importer_data, base_sheet_name, 
                                                          global_incoterm, incoterm_mode)
                    if sheet_result:
                        workbook_data_for_excel_js.append(sheet_result)
        
        # Write output to file
        if workbook_data_for_excel_js:
            output_file = self.formatter.write_output_to_file(workbook_data_for_excel_js, output_filename, period_year)
            self.logger.info(f"Process completed. Output saved to: {output_file}")
            return output_file
        else:
            self.logger.warning("No data was processed for output Excel.")
            return ""
