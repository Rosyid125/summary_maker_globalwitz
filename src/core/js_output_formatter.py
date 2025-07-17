"""
Output Formatter Module - Handles Excel output formatting
Implements the exact logic from the original JavaScript outputFormatter
Now using xlsxwriter for advanced formatting to match ExcelJS output
"""

import xlsxwriter
import os
from typing import Dict, List, Any, Optional

from ..utils.constants import MONTH_ORDER, DEFAULT_OUTPUT_FOLDER
from ..utils.helpers import average_greater_than_zero, format_american_number, format_price_with_precision, format_qty_with_precision

class OutputFormatter:
    """Handles Excel output formatting with JavaScript-compatible logic"""
    
    def __init__(self, logger):
        self.logger = logger
    
    def extract_incoterm_from_value(self, incoterm_value: str) -> str:
        """Extract first 3 uppercase characters from incoterm value"""
        if not incoterm_value or not isinstance(incoterm_value, str):
            return "-"
        
        incoterm_clean = incoterm_value.strip().upper()
        if len(incoterm_clean) >= 3:
            return incoterm_clean[:3]
        else:
            return "-"
    
    def get_incoterm_for_combination(self, combo: Dict, raw_data: List[Dict], 
                                   incoterm_mode: str, default_incoterm: str) -> str:
        """Get incoterm value for a specific combination based on mode"""
        if incoterm_mode == "manual":
            return default_incoterm
        
        # For from_column mode, find the first matching row and extract incoterm
        for row in raw_data:
            if (row.get('hsCode') == combo['hsCode'] and 
                row.get('item') == combo['item'] and 
                row.get('gsm') == combo['gsm'] and 
                row.get('addOn') == combo['addOn']):
                
                incoterm_raw = row.get('incoterms', '')
                return self.extract_incoterm_from_value(incoterm_raw)
        
        return "-"
    
    def prepare_group_block(self, group_name: str, summary_lvl1_data: List[Dict], 
                          summary_lvl2_data: List[Dict], incoterm_value: str, 
                          incoterm_mode: str = "manual", raw_data: List[Dict] = None,
                          supplier_as_sheet: str = "tidak") -> Dict[str, Any]:
        """
        Prepare group block exactly like JavaScript prepareGroupBlock function
        
        Args:
            group_name: Name of the supplier/group
            summary_lvl1_data: Monthly summary data
            summary_lvl2_data: Overall summary data
            incoterm_value: INCOTERM value to use (for manual mode)
            incoterm_mode: Mode for incoterm handling ("manual" or "from_column")
            raw_data: Raw data for extracting incoterms per row (for from_column mode)
            supplier_as_sheet: Whether supplier is used as sheet ("ya" or "tidak")
            
        Returns:
            Dict with group block data
        """
        group_block_rows = []
        header_row_count = 2
        
        # Create header rows - adjust based on supplier_as_sheet mode
        if supplier_as_sheet == "ya":
            header_row1 = ["IMPORTER", "HS CODE", "ITEM", "GSM", "ADD ON"]
        else:
            header_row1 = ["SUPPLIER", "HS CODE", "ITEM", "GSM", "ADD ON"]
        
        header_row2 = [None, None, None, None, None]
        
        for month in MONTH_ORDER:
            header_row1.extend([month, None])
            header_row2.extend(["PRICE", "QTY"])
        
        header_row1.extend(["RECAP", None, None])
        header_row2.extend(["AVG PRICE", "INCOTERM", "TOTAL QTY"])
        
        group_block_rows.append(header_row1)
        group_block_rows.append(header_row2)
        
        monthly_totals = [0] * 12
        
        # Get distinct combinations
        distinct_combinations = []
        for item in summary_lvl2_data:
            combo = {
                'hsCode': item['hsCode'],
                'item': item['item'],
                'gsm': item['gsm'],
                'addOn': item['addOn']
            }
            if combo not in distinct_combinations:
                distinct_combinations.append(combo)
        
        # Sort distinct combinations - ensure all values are strings to avoid comparison errors
        def safe_sort_key(x):
            return (
                str(x['hsCode']) if x['hsCode'] is not None else "",
                str(x['item']) if x['item'] is not None else "",
                str(x['gsm']) if x['gsm'] is not None else "",
                str(x['addOn']) if x['addOn'] is not None else ""
            )
        
        distinct_combinations.sort(key=safe_sort_key)
        
        # Create data rows
        for index, combo in enumerate(distinct_combinations):
            data_row = []
            data_row.append(group_name if index == 0 else None)
            data_row.append(combo['hsCode'])
            data_row.append(combo['item'])
            data_row.append(combo['gsm'])
            data_row.append(combo['addOn'])
            
            # Add monthly data
            for month_index, month in enumerate(MONTH_ORDER):
                month_data = None
                for d in summary_lvl1_data:
                    if (d['hsCode'] == combo['hsCode'] and 
                        d['item'] == combo['item'] and 
                        d['gsm'] == combo['gsm'] and 
                        d['addOn'] == combo['addOn'] and 
                        d['month'] == month):
                        month_data = d
                        break
                
                if month_data:
                    avg_price = format_price_with_precision(month_data['avgPrice']) if month_data['avgPrice'] else "-"
                    qty = format_qty_with_precision(month_data['totalQty'])
                    data_row.extend([avg_price, qty])
                    monthly_totals[month_index] += month_data['totalQty'] if month_data['totalQty'] else 0
                else:
                    data_row.extend(["-", "-"])
            
            # Add recap data
            recap_data = None
            for d in summary_lvl2_data:
                if (d['hsCode'] == combo['hsCode'] and 
                    d['item'] == combo['item'] and 
                    d['gsm'] == combo['gsm'] and 
                    d['addOn'] == combo['addOn']):
                    recap_data = d
                    break
            
            if recap_data:
                avg_price = format_price_with_precision(recap_data['avgOfSummaryPrice']) if recap_data['avgOfSummaryPrice'] else "-"
                # Get incoterm based on mode
                combo_incoterm = self.get_incoterm_for_combination(combo, raw_data or [], incoterm_mode, incoterm_value)
                data_row.extend([avg_price, combo_incoterm, format_qty_with_precision(recap_data['totalOfSummaryQty'])])
            else:
                data_row.extend(["-", "-", "-"])
            
            group_block_rows.append(data_row)
        
        # Calculate overall total
        overall_total_qty = sum(qty for qty in monthly_totals if isinstance(qty, (int, float)))
        
        if distinct_combinations:
            # Add total qty per month row
            total_qty_per_mo_row = ["TOTAL QTY PER MO", "-", "-", "-", "-"]
            for total in monthly_totals:
                # If total is 0 or None, show "-" instead of empty cell
                display_total = format_qty_with_precision(total) if (isinstance(total, (int, float)) and total > 0) else "-"
                total_qty_per_mo_row.extend([display_total, "-"])
            total_qty_per_mo_row.extend([format_qty_with_precision(overall_total_qty), "-", "-"])
            group_block_rows.append(total_qty_per_mo_row)
            
            # Add quarterly totals
            quarterly_totals = [0, 0, 0, 0]
            for index, total in enumerate(monthly_totals):
                num_total = total if isinstance(total, (int, float)) else 0
                if index < 3:
                    quarterly_totals[0] += num_total
                elif index < 6:
                    quarterly_totals[1] += num_total
                elif index < 9:
                    quarterly_totals[2] += num_total
                else:
                    quarterly_totals[3] += num_total
            
            total_qty_per_quartal_row = ["TOTAL QTY PER QUARTAL", "-", "-", "-", "-"]
            # Q1 (Jan-Mar)
            q1_display = format_qty_with_precision(quarterly_totals[0]) if quarterly_totals[0] > 0 else "-"
            total_qty_per_quartal_row.extend([q1_display, "-", "-", "-", "-", "-"])
            # Q2 (Apr-Jun)  
            q2_display = format_qty_with_precision(quarterly_totals[1]) if quarterly_totals[1] > 0 else "-"
            total_qty_per_quartal_row.extend([q2_display, "-", "-", "-", "-", "-"])
            # Q3 (Jul-Sep)
            q3_display = format_qty_with_precision(quarterly_totals[2]) if quarterly_totals[2] > 0 else "-"
            total_qty_per_quartal_row.extend([q3_display, "-", "-", "-", "-", "-"])
            # Q4 (Oct-Dec)
            q4_display = format_qty_with_precision(quarterly_totals[3]) if quarterly_totals[3] > 0 else "-"
            total_qty_per_quartal_row.extend([q4_display, "-", "-", "-", "-", "-"])
            total_qty_per_quartal_row.extend(["-", "-", "-"])
            group_block_rows.append(total_qty_per_quartal_row)
        
        return {
            'groupBlockRows': group_block_rows,
            'overallTotalQtyForGroup': overall_total_qty,
            'distinctCombinationsCount': len(distinct_combinations),
            'headerRowCount': header_row_count,
            'header1Length': len(header_row1)
        }
    
    def write_output_to_file(self, workbook_data: List[Dict], output_filename: str = "summary_output.xlsx", 
                           period_year: str = None, supplier_as_sheet: str = "tidak") -> str:
        """
        Write output to Excel file with advanced formatting using xlsxwriter
        Matches the ExcelJS formatting from the original JavaScript version
        
        Args:
            workbook_data: List of sheet data
            output_filename: Output filename
            period_year: Year for the period title
            supplier_as_sheet: Whether supplier is used as sheet ("ya" or "tidak")
            
        Returns:
            str: Path to output file
        """
        try:
            self.logger.info(f"Starting write_output_to_file with {len(workbook_data)} sheets")
            
            # Validate workbook data
            if not workbook_data:
                raise ValueError("No workbook data provided")
            
            # Check that each sheet has required structure
            for i, sheet_info in enumerate(workbook_data):
                if not isinstance(sheet_info, dict):
                    raise ValueError(f"Sheet {i} is not a dictionary")
                if 'name' not in sheet_info:
                    raise ValueError(f"Sheet {i} missing 'name' field")
                if 'allRowsForSheetContent' not in sheet_info:
                    raise ValueError(f"Sheet {i} missing 'allRowsForSheetContent' field")
                if 'totalColumns' not in sheet_info:
                    raise ValueError(f"Sheet {i} missing 'totalColumns' field")
                self.logger.info(f"Sheet {i}: '{sheet_info['name']}' - {len(sheet_info['allRowsForSheetContent'])} rows, {sheet_info['totalColumns']} columns")
            
            # Get absolute path for output folder with fallback
            try:
                output_folder = os.path.abspath(DEFAULT_OUTPUT_FOLDER)
                self.logger.info(f"Primary output folder: {output_folder}")
                
                # Test write access
                test_file = os.path.join(output_folder, "test_write.tmp")
                with open(test_file, 'w') as f:
                    f.write("test")
                os.remove(test_file)
                self.logger.info("Write permissions confirmed for primary folder")
                
            except Exception as primary_error:
                self.logger.warning(f"Primary output folder failed: {primary_error}")
                # Fallback to user's desktop
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                output_folder = os.path.join(desktop_path, "ExcelSummaryMaker_Output")
                self.logger.info(f"Using fallback output folder: {output_folder}")
                
                try:
                    os.makedirs(output_folder, exist_ok=True)
                    # Test write access to fallback
                    test_file = os.path.join(output_folder, "test_write.tmp")
                    with open(test_file, 'w') as f:
                        f.write("test")
                    os.remove(test_file)
                    self.logger.info("Write permissions confirmed for fallback folder")
                except Exception as fallback_error:
                    raise Exception(f"No writable output folder found. Primary: {primary_error}, Fallback: {fallback_error}")
            
            if not os.path.exists(output_folder):
                self.logger.info(f"Creating output directory: {output_folder}")
                os.makedirs(output_folder, exist_ok=True)
            
            output_file = os.path.join(output_folder, output_filename)
            self.logger.info(f"Output file path: {output_file}")
            
            # Define colors (matching JavaScript colors)
            colors = {
                'period': '#7030A0',
                'supplierCols': '#002060',
                'q1': '#FFC000',
                'q2': '#00B050',
                'q3': '#FFFF00',
                'q4': '#00B0F0',
                'recap': '#002060',
                'totalPerItemTitle': '#FF0000',
                'textWhite': '#FFFFFF',
                'textBlack': '#000000'
            }
            
            # Create workbook with xlsxwriter
            self.logger.info("Creating xlsxwriter workbook...")
            workbook = xlsxwriter.Workbook(output_file)
            
            # Define formats
            formats = self._create_formats(workbook, colors)
            
            for i, sheet_info in enumerate(workbook_data):
                self.logger.info(f"Processing sheet {i+1}/{len(workbook_data)}: {sheet_info['name']}")
                worksheet = workbook.add_worksheet(sheet_info['name'])
                
                # Add period title
                if period_year:
                    period_title = f"{period_year} PERIODE"
                    worksheet.merge_range(0, 0, 0, sheet_info['totalColumns'] - 1, period_title, formats['period_title'])
                    worksheet.set_row(0, 20)
                
                # Add sheet content
                start_row = 1 if period_year else 0
                current_row = start_row
                
                # Apply advanced formatting (this will also write the data)
                self.logger.info(f"Applying formatting to sheet: {sheet_info['name']}")
                self._apply_advanced_formatting(worksheet, sheet_info, formats, start_row)
            
            self.logger.info("Closing workbook...")
            workbook.close()
            
            # Verify the file was created
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                self.logger.info(f"Output file created successfully: {output_file} (size: {file_size} bytes)")
                return output_file
            else:
                raise Exception(f"Output file was not created: {output_file}")
                
        except Exception as e:
            self.logger.error(f"Error in write_output_to_file: {str(e)}")
            raise Exception(f"Failed to write output file: {str(e)}")
    
    def _create_formats(self, workbook, colors):
        """Create all the formats needed for the Excel output"""
        formats = {}
        
        # Period title format
        formats['period_title'] = workbook.add_format({
            'bg_color': colors['period'],
            'font_color': colors['textWhite'],
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Supplier columns format
        formats['supplier_cols'] = workbook.add_format({
            'bg_color': colors['supplierCols'],
            'font_color': colors['textWhite'],
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        # Quarter formats
        formats['q1'] = workbook.add_format({
            'bg_color': colors['q1'],
            'font_color': colors['textBlack'],
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        formats['q2'] = workbook.add_format({
            'bg_color': colors['q2'],
            'font_color': colors['textWhite'],
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        formats['q3'] = workbook.add_format({
            'bg_color': colors['q3'],
            'font_color': colors['textBlack'],
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        formats['q4'] = workbook.add_format({
            'bg_color': colors['q4'],
            'font_color': colors['textWhite'],
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        # Recap format
        formats['recap'] = workbook.add_format({
            'bg_color': colors['recap'],
            'font_color': colors['textWhite'],
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        # Total per item title format
        formats['total_per_item_title'] = workbook.add_format({
            'bg_color': colors['totalPerItemTitle'],
            'font_color': colors['textWhite'],
            'bold': True,
            'font_size': 12,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Data cell format
        formats['data_cell'] = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': '#,##0.00'  # American number format
        })
        
        # Price cell format with controlled precision (for price values)
        formats['price_cell'] = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': '#,##0.000'  # Exactly 3 decimal places
        })
        
        # Data cell format without border (for separator rows)
        formats['no_border_cell'] = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#,##0.000'  # Exactly 3 decimal places
        })
        
        # Text format for GSM and other string fields
        formats['text_cell'] = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': '@'  # Text format
        })
        
        # Number format for quantities (no decimal places for whole numbers, 3 for decimals)
        formats['qty_cell'] = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': '#,##0'  # American number format for whole numbers
        })
        
        # Bold data format
        formats['bold_data'] = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': '#,##0.000'  # Exactly 3 decimal places
        })
        
        # Bold price format with controlled precision
        formats['bold_price'] = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': '#,##0.000'  # Exactly 3 decimal places
        })
        
        # Total all supplier/importer formats
        formats['total_all_period'] = workbook.add_format({
            'bg_color': colors['period'],
            'font_color': colors['textWhite'],
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        return formats
    
    def _apply_advanced_formatting(self, worksheet, sheet_info, formats, start_row):
        """Apply advanced formatting to match ExcelJS output"""
        # Apply borders and center alignment to data cells, but skip empty separator rows
        total_rows = len(sheet_info['allRowsForSheetContent'])
        total_cols = sheet_info['totalColumns']
        
        # First pass: identify which rows should have borders (non-empty content rows)
        rows_with_content = set()
        for row_idx in range(total_rows):
            row_data = sheet_info['allRowsForSheetContent'][row_idx]
            # Check if row has any meaningful content (not just empty strings or None)
            has_content = any(cell is not None and str(cell).strip() != "" for cell in row_data[:total_cols])
            if has_content:
                rows_with_content.add(row_idx)
        
        # Apply formatting only to rows with content
        for row_idx in range(total_rows):
            actual_row = start_row + row_idx
            row_data = sheet_info['allRowsForSheetContent'][row_idx]
            
            if row_idx in rows_with_content:
                # Apply format with borders for content rows
                for col_idx in range(total_cols):
                    cell_value = row_data[col_idx] if col_idx < len(row_data) else ""
                    
                    # Use text format for GSM column (column 3, zero-indexed)
                    if col_idx == 3:  # GSM column
                        worksheet.write(actual_row, col_idx, cell_value, formats['text_cell'])
                    else:
                        worksheet.write(actual_row, col_idx, cell_value, formats['data_cell'])
            else:
                # Apply format without borders for separator rows
                for col_idx in range(total_cols):
                    cell_value = row_data[col_idx] if col_idx < len(row_data) else ""
                    
                    # Use text format for GSM column (column 3, zero-indexed)
                    if col_idx == 3:  # GSM column
                        worksheet.write(actual_row, col_idx, cell_value, formats['text_cell'])
                    else:
                        worksheet.write(actual_row, col_idx, cell_value, formats['no_border_cell'])
        
        # Format supplier group headers
        current_row = start_row
        for group_meta in sheet_info.get('supplierGroupsMeta', []):
            header_start_row = current_row
            data_start_row = header_start_row + group_meta['headerRowCount']
            product_rows = group_meta['productRowCount']
            
            # Format header rows
            self._format_group_headers(worksheet, header_start_row, formats, sheet_info, start_row)
            
            # Merge supplier name cell
            if product_rows > 0:
                supplier_name = ""
                if data_start_row - start_row < len(sheet_info['allRowsForSheetContent']):
                    supplier_name = sheet_info['allRowsForSheetContent'][data_start_row - start_row][0] or ""
                
                worksheet.merge_range(
                    data_start_row, 0, 
                    data_start_row + product_rows - 1, 0, 
                    supplier_name,
                    formats['data_cell']
                )
            
            # Format total rows
            if product_rows > 0:
                total_qty_per_mo_row = data_start_row + product_rows
                quartal_row = total_qty_per_mo_row + 1
                
                # Format total qty per month row
                worksheet.merge_range(total_qty_per_mo_row, 0, total_qty_per_mo_row, 4, 
                                    "TOTAL QTY PER MO", formats['bold_data'])
                
                # Merge monthly total cells
                col = 5
                for i in range(12):  # 12 months
                    value = ""
                    if (total_qty_per_mo_row - start_row < len(sheet_info['allRowsForSheetContent']) and 
                        col < len(sheet_info['allRowsForSheetContent'][total_qty_per_mo_row - start_row])):
                        value = sheet_info['allRowsForSheetContent'][total_qty_per_mo_row - start_row][col] or ""
                    
                    worksheet.merge_range(total_qty_per_mo_row, col, total_qty_per_mo_row, col + 1, 
                                        value, formats['data_cell'])
                    col += 2
                
                # Format quarterly total row
                worksheet.merge_range(quartal_row, 0, quartal_row, 4, 
                                    "TOTAL QTY PER QUARTAL", formats['bold_data'])
                
                # Merge quarterly cells
                col = 5
                for q in range(4):  # 4 quarters
                    value = ""
                    if (quartal_row - start_row < len(sheet_info['allRowsForSheetContent']) and 
                        col < len(sheet_info['allRowsForSheetContent'][quartal_row - start_row])):
                        value = sheet_info['allRowsForSheetContent'][quartal_row - start_row][col] or ""
                    
                    worksheet.merge_range(quartal_row, col, quartal_row, col + 5, 
                                        value, formats['data_cell'])
                    col += 6
                
                # Merge recap cell for totals
                recap_start_col = 5 + 12 * 2  # 5 + 24 = 29
                recap_value = ""
                if (total_qty_per_mo_row - start_row < len(sheet_info['allRowsForSheetContent']) and 
                    recap_start_col < len(sheet_info['allRowsForSheetContent'][total_qty_per_mo_row - start_row])):
                    recap_value = sheet_info['allRowsForSheetContent'][total_qty_per_mo_row - start_row][recap_start_col] or ""
                
                worksheet.merge_range(total_qty_per_mo_row, recap_start_col, quartal_row, recap_start_col + 2, 
                                    recap_value, formats['bold_data'])
            
            # Move to next group
            current_row += group_meta['headerRowCount'] + product_rows + (2 if product_rows > 0 else 0)
            if group_meta.get('hasFollowingGroup'):
                current_row += 1
        
        # Format "TOTAL ALL SUPPLIER/IMPORTER" section
        self._format_total_all_supplier_section(worksheet, sheet_info, formats, start_row)
        
        # Format "TOTAL PER ITEM" section
        self._format_total_per_item_section(worksheet, sheet_info, formats, start_row)
    
    def _format_group_headers(self, worksheet, header_start_row, formats, sheet_info, start_row):
        """Format the group headers with proper colors and merging"""
        # Get header row data
        header_row_data = []
        if header_start_row - start_row < len(sheet_info['allRowsForSheetContent']):
            header_row_data = sheet_info['allRowsForSheetContent'][header_start_row - start_row]
        
        # Merge cells for supplier columns
        for col in range(5):
            value = header_row_data[col] if col < len(header_row_data) else ""
            worksheet.merge_range(header_start_row, col, header_start_row + 1, col, 
                                value, formats['supplier_cols'])
        
        # Format monthly columns with quarterly colors
        col = 5
        quarter_formats = ['q1', 'q2', 'q3', 'q4']
        
        for q in range(4):  # 4 quarters
            q_format = formats[quarter_formats[q]]
            for i in range(3):  # 3 months per quarter
                # Get month name
                month_value = header_row_data[col] if col < len(header_row_data) else ""
                
                # Merge month header
                worksheet.merge_range(header_start_row, col, header_start_row, col + 1, 
                                    month_value, q_format)
                # Format price and qty cells
                worksheet.write(header_start_row + 1, col, "PRICE", q_format)
                worksheet.write(header_start_row + 1, col + 1, "QTY", q_format)
                col += 2
        
        # Format recap columns
        worksheet.merge_range(header_start_row, col, header_start_row, col + 2, "RECAP", formats['recap'])
        worksheet.write(header_start_row + 1, col, "AVG PRICE", formats['recap'])
        worksheet.write(header_start_row + 1, col + 1, "INCOTERM", formats['recap'])
        worksheet.write(header_start_row + 1, col + 2, "TOTAL QTY", formats['recap'])
    
    def _format_total_all_supplier_section(self, worksheet, sheet_info, formats, start_row):
        """Format the TOTAL ALL SUPPLIER/IMPORTER section"""
        # Find the total all supplier/importer section
        total_all_rows = []
        for i, row_data in enumerate(sheet_info['allRowsForSheetContent']):
            if row_data and str(row_data[0]).strip() == "Month":
                total_all_rows = [i, i + 1, i + 2]  # Month header, Total per month, Total per quarter
                break
        
        if not total_all_rows:
            return
        
        # Format Month header row
        month_header_row = start_row + total_all_rows[0]
        worksheet.merge_range(month_header_row, 0, month_header_row, 4, "Month", formats['supplier_cols'])
        
        # Format monthly columns with quarterly colors
        col = 5
        quarter_formats = ['q1', 'q2', 'q3', 'q4']
        
        for q in range(4):
            q_format = formats[quarter_formats[q]]
            for i in range(3):
                worksheet.merge_range(month_header_row, col, month_header_row, col + 1, 
                                    sheet_info['allRowsForSheetContent'][total_all_rows[0]][col], q_format)
                col += 2
        
        # Format recap
        recap_start_col = 5 + 12 * 2
        worksheet.merge_range(month_header_row, recap_start_col, month_header_row, recap_start_col + 2, 
                            "RECAP", formats['recap'])
        
        # Format total per month row - get the actual text from the data
        total_per_month_row = start_row + total_all_rows[1]
        month_text = sheet_info['allRowsForSheetContent'][total_all_rows[1]][0]
        worksheet.merge_range(total_per_month_row, 0, total_per_month_row, 4, 
                            month_text, formats['total_all_period'])
        
        col = 5
        for i in range(12):
            worksheet.merge_range(total_per_month_row, col, total_per_month_row, col + 1, 
                                sheet_info['allRowsForSheetContent'][total_all_rows[1]][col], 
                                formats['total_all_period'])
            col += 2
        
        # Format total per quarter row - get the actual text from the data
        total_per_quarter_row = start_row + total_all_rows[2]
        quarter_text = sheet_info['allRowsForSheetContent'][total_all_rows[2]][0]
        worksheet.merge_range(total_per_quarter_row, 0, total_per_quarter_row, 4, 
                            quarter_text, formats['total_all_period'])
        
        col = 5
        for q in range(4):
            worksheet.merge_range(total_per_quarter_row, col, total_per_quarter_row, col + 5, 
                                sheet_info['allRowsForSheetContent'][total_all_rows[2]][col], 
                                formats['total_all_period'])
            col += 6
        
        # Merge recap cell
        worksheet.merge_range(total_per_month_row, recap_start_col, total_per_quarter_row, recap_start_col + 2, 
                            sheet_info['allRowsForSheetContent'][total_all_rows[1]][recap_start_col], 
                            formats['total_all_period'])
    
    def _format_total_per_item_section(self, worksheet, sheet_info, formats, start_row):
        """Format the TOTAL PER ITEM section"""
        # Find the total per item section
        total_per_item_start = -1
        total_per_item_header = -1
        
        for i, row_data in enumerate(sheet_info['allRowsForSheetContent']):
            if row_data and str(row_data[0]).strip() == "TOTAL PER ITEM":
                total_per_item_start = i
                total_per_item_header = i + 1
                break
        
        if total_per_item_start == -1:
            return
        
        # Format title row
        title_row = start_row + total_per_item_start
        worksheet.merge_range(title_row, 0, title_row, sheet_info['totalColumns'] - 1, 
                            "TOTAL PER ITEM", formats['total_per_item_title'])
        worksheet.set_row(title_row, 18)
        
        # Format header row
        header_row = start_row + total_per_item_header
        worksheet.merge_range(header_row, 0, header_row, 4, 
                            sheet_info['allRowsForSheetContent'][total_per_item_header][0], 
                            formats['supplier_cols'])
        
        # Format monthly columns
        col = 5
        quarter_formats = ['q1', 'q2', 'q3', 'q4']
        
        for q in range(4):
            q_format = formats[quarter_formats[q]]
            for i in range(3):
                worksheet.merge_range(header_row, col, header_row, col + 1, 
                                    sheet_info['allRowsForSheetContent'][total_per_item_header][col], q_format)
                col += 2
        
        # Format recap
        recap_start_col = 5 + 12 * 2
        worksheet.merge_range(header_row, recap_start_col, header_row, recap_start_col + 2, 
                            "RECAP", formats['recap'])
        
        # Format item rows
        current_item_row = header_row + 1
        while current_item_row < start_row + len(sheet_info['allRowsForSheetContent']):
            row_data = sheet_info['allRowsForSheetContent'][current_item_row - start_row]
            if not row_data or not str(row_data[0]).strip():
                break
            
            # Check if this is a valid item row
            first_cell = str(row_data[0]).strip()
            if (first_cell.startswith("TOTAL ALL") or 
                first_cell.startswith("Month") or 
                first_cell.startswith("TOTAL QTY PER MO") or 
                first_cell.startswith("TOTAL QTY PER QUARTAL") or 
                first_cell.startswith("TOTAL PER ITEM")):
                break
            
            # Merge item columns
            worksheet.merge_range(current_item_row, 0, current_item_row, 4, first_cell, formats['data_cell'])
            
            # Merge monthly columns
            col = 5
            for i in range(12):
                worksheet.merge_range(current_item_row, col, current_item_row, col + 1, 
                                    row_data[col] if col < len(row_data) else "", formats['data_cell'])
                col += 2
            
            # Merge recap columns
            worksheet.merge_range(current_item_row, recap_start_col, current_item_row, recap_start_col + 2, 
                                row_data[recap_start_col] if recap_start_col < len(row_data) else "", 
                                formats['data_cell'])
            
            current_item_row += 1
    
    def extract_incoterm_from_value(self, incoterm_value: str) -> str:
        """
        Extract first 3 uppercase characters from incoterm value
        
        Args:
            incoterm_value: Raw incoterm value from data
            
        Returns:
            str: First 3 uppercase characters or "-" if invalid
        """
        if not incoterm_value or not isinstance(incoterm_value, str):
            return "-"
        
        # Extract first 3 characters and convert to uppercase
        incoterm_clean = incoterm_value.strip().upper()
        if len(incoterm_clean) >= 3:
            return incoterm_clean[:3]
        else:
            return "-"
    
    def get_incoterm_for_combination(self, combo: Dict, raw_data: List[Dict], 
                                   incoterm_mode: str, default_incoterm: str) -> str:
        """
        Get incoterm value for a specific combination based on mode
        
        Args:
            combo: Combination dict with hsCode, item, gsm, addOn
            raw_data: Raw data to search for incoterm
            incoterm_mode: "manual" or "from_column"
            default_incoterm: Default incoterm for manual mode
            
        Returns:
            str: Incoterm value to use
        """
        if incoterm_mode == "manual":
            return default_incoterm
        
        # For from_column mode, find the first matching row and extract incoterm
        for row in raw_data:
            if (row.get('hsCode') == combo['hsCode'] and 
                row.get('item') == combo['item'] and 
                row.get('gsm') == combo['gsm'] and 
                row.get('addOn') == combo['addOn']):
                
                incoterm_raw = row.get('incoterms', '')
                return self.extract_incoterm_from_value(incoterm_raw)
        
        return "-"
