"""
Output Formatter Module - Handles Excel output generation with formatting
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import os
from datetime import datetime
from typing import Dict, List, Any, Optional

from ..utils.helpers import format_currency, get_month_name, format_american_number, format_qty_with_precision

class OutputFormatter:
    """Handles Excel output generation with complex formatting"""
    
    def __init__(self, logger):
        self.logger = logger
        self.workbook = None
        self.quarter_colors = {
            1: PatternFill(start_color="FFE6E6", end_color="FFE6E6", fill_type="solid"),  # Light Red
            2: PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid"),  # Light Blue
            3: PatternFill(start_color="E6FFE6", end_color="E6FFE6", fill_type="solid"),  # Light Green
            4: PatternFill(start_color="FFFFE6", end_color="FFFFE6", fill_type="solid"),  # Light Yellow
        }
        
        # Define styles
        self.header_font = Font(bold=True, size=12)
        self.title_font = Font(bold=True, size=14)
        self.border_thin = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        self.center_alignment = Alignment(horizontal='center', vertical='center')
    
    def create_output_file(self, aggregated_data: Dict[str, Any], output_path: str, 
                          incoterm: str = "FOB", year: int = None) -> bool:
        """
        Create formatted Excel output file
        
        Args:
            aggregated_data (Dict): Aggregated data from DataAggregator
            output_path (str): Output file path
            incoterm (str): INCOTERM for pricing
            year (int): Year for the report
            
        Returns:
            bool: True if successful
        """
        try:
            self.workbook = Workbook()
            
            # Remove default sheet
            if 'Sheet' in self.workbook.sheetnames:
                self.workbook.remove(self.workbook['Sheet'])
            
            # Create sheets for each importer
            for importer_name, importer_data in aggregated_data.items():
                self._create_importer_sheet(importer_name, importer_data, incoterm, year)
            
            # Create summary sheet
            self._create_summary_sheet(aggregated_data, incoterm, year)
            
            # Save workbook
            self.workbook.save(output_path)
            self.logger.info(f"Output file created: {output_path}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error creating output file: {str(e)}")
            return False
    
    def _create_importer_sheet(self, importer_name: str, importer_data: Dict[str, Any], 
                              incoterm: str, year: int):
        """Create sheet for individual importer"""
        try:
            # Create worksheet
            sheet_name = self._sanitize_sheet_name(importer_name)
            ws = self.workbook.create_sheet(title=sheet_name)
            
            current_row = 1
            
            # Title
            ws.cell(row=current_row, column=1, value=f"Import Summary - {importer_name}")
            ws.cell(row=current_row, column=1).font = self.title_font
            current_row += 2
            
            # Overall summary
            if 'overall_summary' in importer_data:
                current_row = self._add_overall_summary(ws, current_row, importer_data['overall_summary'], incoterm)
                current_row += 2
            
            # Monthly summary
            if 'monthly_summary' in importer_data and importer_data['monthly_summary']:
                current_row = self._add_monthly_summary(ws, current_row, importer_data['monthly_summary'], incoterm)
                current_row += 2
            
            # Supplier summary
            if 'supplier_summary' in importer_data and importer_data['supplier_summary']:
                current_row = self._add_supplier_summary(ws, current_row, importer_data['supplier_summary'], incoterm)
                current_row += 2
            
            # Item summary
            if 'item_summary' in importer_data and importer_data['item_summary']:
                current_row = self._add_item_summary(ws, current_row, importer_data['item_summary'], incoterm)
            
            # Auto-adjust column widths
            self._auto_adjust_columns(ws)
            
        except Exception as e:
            self.logger.error(f"Error creating sheet for {importer_name}: {str(e)}")
    
    def _add_overall_summary(self, ws, start_row: int, summary: Dict[str, Any], incoterm: str) -> int:
        """Add overall summary section"""
        current_row = start_row
        
        # Section title
        ws.cell(row=current_row, column=1, value="Overall Summary")
        ws.cell(row=current_row, column=1).font = self.header_font
        current_row += 1
        
        # Summary data
        summary_items = [
            ("Total Records", summary.get('total_records', 0)),
            ("Total Quantity", format_qty_with_precision(summary.get('total_quantity', 0))),
            ("Average Unit Price", f"{incoterm} {format_american_number(summary.get('avg_unit_price', 0))}" if summary.get('avg_unit_price') else "N/A"),
            ("Total Value", f"{incoterm} {format_american_number(summary.get('total_value', 0))}"),
            ("Unique Suppliers", summary.get('unique_suppliers', 0)),
            ("Unique Items", summary.get('unique_items', 0)),
            ("Unique HS Codes", summary.get('unique_hs_codes', 0))
        ]
        
        for label, value in summary_items:
            ws.cell(row=current_row, column=1, value=label)
            ws.cell(row=current_row, column=2, value=value)
            ws.cell(row=current_row, column=1).font = Font(bold=True)
            current_row += 1
        
        return current_row
    
    def _add_monthly_summary(self, ws, start_row: int, monthly_data: List[Dict[str, Any]], incoterm: str) -> int:
        """Add monthly summary table"""
        current_row = start_row
        
        # Section title
        ws.cell(row=current_row, column=1, value="Monthly Summary")
        ws.cell(row=current_row, column=1).font = self.header_font
        current_row += 1
        
        # Headers
        headers = [
            "Month", "HS Code", "Item", "GSM", "Add On", 
            "Total Quantity", f"Avg Unit Price ({incoterm})", f"Total Value ({incoterm})", "Records"
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=col, value=header)
            cell.font = self.header_font
            cell.border = self.border_thin
            cell.alignment = self.center_alignment
        
        current_row += 1
        
        # Data rows
        for row_data in monthly_data:
            quarter = self._get_quarter_from_month_year(row_data.get('month_year', ''))
            quarter_fill = self.quarter_colors.get(quarter, None)
            
            values = [
                row_data.get('month_name', ''),
                row_data.get('hs_code', ''),
                row_data.get('item', ''),
                row_data.get('gsm', ''),
                row_data.get('add_on', ''),
                format_american_number(row_data.get('total_quantity', 0), 0),
                format_american_number(row_data.get('avg_unit_price', 0)) if row_data.get('avg_unit_price') else "N/A",
                format_american_number(row_data.get('total_value', 0)),
                row_data.get('record_count', 0)
            ]
            
            for col, value in enumerate(values, 1):
                cell = ws.cell(row=current_row, column=col, value=value)
                cell.border = self.border_thin
                if quarter_fill:
                    cell.fill = quarter_fill
                
                # Right align numbers
                if col >= 6:  # Numeric columns
                    cell.alignment = Alignment(horizontal='right')
            
            current_row += 1
        
        return current_row
    
    def _add_supplier_summary(self, ws, start_row: int, supplier_data: List[Dict[str, Any]], incoterm: str) -> int:
        """Add supplier summary table"""
        current_row = start_row
        
        # Section title
        ws.cell(row=current_row, column=1, value="Supplier Summary")
        ws.cell(row=current_row, column=1).font = self.header_font
        current_row += 1
        
        # Headers
        headers = [
            "Supplier", "Total Quantity", f"Avg Unit Price ({incoterm})", 
            f"Total Value ({incoterm})", "Records", "Unique Items"
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=col, value=header)
            cell.font = self.header_font
            cell.border = self.border_thin
            cell.alignment = self.center_alignment
        
        current_row += 1
        
        # Data rows
        for row_data in supplier_data:
            values = [
                row_data.get('supplier', ''),
                format_qty_with_precision(row_data.get('total_quantity', 0)),
                format_american_number(row_data.get('avg_unit_price', 0)) if row_data.get('avg_unit_price') else "N/A",
                format_american_number(row_data.get('total_value', 0)),
                row_data.get('record_count', 0),
                row_data.get('unique_items', 0)
            ]
            
            for col, value in enumerate(values, 1):
                cell = ws.cell(row=current_row, column=col, value=value)
                cell.border = self.border_thin
                
                # Right align numbers
                if col >= 2:  # Numeric columns
                    cell.alignment = Alignment(horizontal='right')
            
            current_row += 1
        
        return current_row
    
    def _add_item_summary(self, ws, start_row: int, item_data: List[Dict[str, Any]], incoterm: str) -> int:
        """Add item summary table"""
        current_row = start_row
        
        # Section title
        ws.cell(row=current_row, column=1, value="Item Summary")
        ws.cell(row=current_row, column=1).font = self.header_font
        current_row += 1
        
        # Headers
        headers = [
            "Item", "Total Quantity", f"Avg Unit Price ({incoterm})", 
            f"Total Value ({incoterm})", "Records", "Unique Suppliers", "Avg GSM"
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=col, value=header)
            cell.font = self.header_font
            cell.border = self.border_thin
            cell.alignment = self.center_alignment
        
        current_row += 1
        
        # Data rows
        for row_data in item_data:
            values = [
                row_data.get('item', ''),
                format_qty_with_precision(row_data.get('total_quantity', 0)),
                format_american_number(row_data.get('avg_unit_price', 0)) if row_data.get('avg_unit_price') else "N/A",
                format_american_number(row_data.get('total_value', 0)),
                row_data.get('record_count', 0),
                row_data.get('unique_suppliers', 0),
                format_american_number(row_data.get('avg_gsm', 0), 1) if row_data.get('avg_gsm') else "N/A"
            ]
            
            for col, value in enumerate(values, 1):
                cell = ws.cell(row=current_row, column=col, value=value)
                cell.border = self.border_thin
                
                # Right align numbers
                if col >= 2:  # Numeric columns
                    cell.alignment = Alignment(horizontal='right')
            
            current_row += 1
        
        return current_row
    
    def _create_summary_sheet(self, aggregated_data: Dict[str, Any], incoterm: str, year: int):
        """Create overall summary sheet"""
        try:
            ws = self.workbook.create_sheet(title="Overall Summary", index=0)
            
            current_row = 1
            
            # Title
            ws.cell(row=current_row, column=1, value=f"Import Summary Report - {year or 'All Years'}")
            ws.cell(row=current_row, column=1).font = self.title_font
            current_row += 2
            
            # Summary by importer
            ws.cell(row=current_row, column=1, value="Summary by Importer")
            ws.cell(row=current_row, column=1).font = self.header_font
            current_row += 1
            
            # Headers
            headers = [
                "Importer", "Total Records", "Total Quantity", 
                f"Total Value ({incoterm})", "Unique Suppliers", "Unique Items"
            ]
            
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=current_row, column=col, value=header)
                cell.font = self.header_font
                cell.border = self.border_thin
                cell.alignment = self.center_alignment
            
            current_row += 1
            
            # Data rows
            total_records = 0
            total_quantity = 0
            total_value = 0
            
            for importer_name, importer_data in aggregated_data.items():
                overall = importer_data.get('overall_summary', {})
                
                values = [
                    importer_name,
                    overall.get('total_records', 0),
                    format_qty_with_precision(overall.get('total_quantity', 0)),
                    format_american_number(overall.get('total_value', 0)),
                    overall.get('unique_suppliers', 0),
                    overall.get('unique_items', 0)
                ]
                
                # Update totals
                total_records += overall.get('total_records', 0)
                total_quantity += overall.get('total_quantity', 0)
                total_value += overall.get('total_value', 0)
                
                for col, value in enumerate(values, 1):
                    cell = ws.cell(row=current_row, column=col, value=value)
                    cell.border = self.border_thin
                    
                    # Right align numbers
                    if col >= 2:
                        cell.alignment = Alignment(horizontal='right')
                
                current_row += 1
            
            # Total row
            total_values = [
                "TOTAL",
                total_records,
                format_qty_with_precision(total_quantity),
                format_american_number(total_value),
                "",  # Can't sum unique suppliers across importers
                ""   # Can't sum unique items across importers
            ]
            
            for col, value in enumerate(total_values, 1):
                cell = ws.cell(row=current_row, column=col, value=value)
                cell.font = Font(bold=True)
                cell.border = self.border_thin
                cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
                
                if col >= 2:
                    cell.alignment = Alignment(horizontal='right')
            
            # Auto-adjust column widths
            self._auto_adjust_columns(ws)
            
        except Exception as e:
            self.logger.error(f"Error creating summary sheet: {str(e)}")
    
    def _get_quarter_from_month_year(self, month_year: str) -> int:
        """Get quarter number from month_year string"""
        try:
            if not month_year or '-' not in month_year:
                return 1
            
            month = int(month_year.split('-')[1])
            return ((month - 1) // 3) + 1
            
        except Exception:
            return 1
    
    def _sanitize_sheet_name(self, name: str) -> str:
        """Sanitize sheet name for Excel compatibility"""
        if not name:
            return "Sheet1"
        
        # Remove invalid characters
        invalid_chars = ['\\', '/', '*', '[', ']', ':', '?']
        for char in invalid_chars:
            name = name.replace(char, '_')
        
        # Limit length
        return name[:31]
    
    def _auto_adjust_columns(self, ws):
        """Auto-adjust column widths based on content"""
        try:
            for column in ws.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                
                for cell in column:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                
                # Set column width with some padding
                adjusted_width = min(max_length + 2, 50)  # Max width of 50
                ws.column_dimensions[column_letter].width = adjusted_width
                
        except Exception as e:
            self.logger.error(f"Error adjusting column widths: {str(e)}")
