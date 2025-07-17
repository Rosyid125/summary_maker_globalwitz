"""
Main Window GUI for Excel Summary Maker
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional

from ..core.excel_reader import ExcelReader
from ..core.data_aggregator import DataAggregator
from ..core.output_formatter import OutputFormatter
from ..core.js_excel_reader import JSStyleExcelReader
from ..core.js_processor import JSStyleProcessor

class MainWindow:
    """Main application window"""
    
    def __init__(self, root, logger):
        self.root = root
        self.logger = logger
        
        # Initialize core components
        self.excel_reader = ExcelReader(logger)
        self.data_aggregator = DataAggregator(logger)
        self.output_formatter = OutputFormatter(logger)
        
        # Initialize JavaScript-style components
        self.js_excel_reader = JSStyleExcelReader(logger)
        self.js_processor = JSStyleProcessor(logger)
        
        # Initialize variables
        self.current_file_path = tk.StringVar()
        self.selected_sheet = tk.StringVar()
        self.date_format = tk.StringVar(value="auto")
        self.number_format = tk.StringVar(value="auto")
        self.target_year = tk.StringVar(value=str(datetime.now().year))
        self.incoterm = tk.StringVar(value="-")
        self.incoterm_mode = tk.StringVar(value="manual")  # "manual" or "from_column"
        self.supplier_as_sheet = tk.StringVar(value="tidak")  # "ya" or "tidak"
        self.output_filename = tk.StringVar()
        
        # Column mapping variables
        self.column_mappings = {
            'date': tk.StringVar(),
            'hs_code': tk.StringVar(),
            'item_description': tk.StringVar(),
            'gsm': tk.StringVar(),
            'item': tk.StringVar(),
            'add_on': tk.StringVar(),
            'importer': tk.StringVar(),
            'supplier': tk.StringVar(),
            'origin_country': tk.StringVar(),
            'unit_price': tk.StringVar(),
            'quantity': tk.StringVar(),
            'incoterms': tk.StringVar()
        }
        
        self.available_columns = []
        self.processing = False
        
        self.setup_ui()
    
    def on_incoterm_mode_change(self, event=None):
        """Handle incoterm mode change to update UI visibility"""
        mode = self.incoterm_mode.get()
        if mode == "manual":
            self.incoterm_combo.config(state="normal")
            self.incoterm_info_label.config(text="(Manual entry - applied to all rows)")
        else:  # from_column
            self.incoterm_combo.config(state="disabled")
            self.incoterm_info_label.config(text="(Read from incoterms column - first 3 chars)")
    
    def update_field_descriptions(self):
        """Update field descriptions based on supplier_as_sheet setting"""
        self.field_descriptions = {
            'date': 'Date/Invoice Date',
            'hs_code': 'HS Code',
            'item_description': 'Item Description',
            'gsm': 'GSM (grams per square meter)',
            'item': 'Item/Product Name',
            'add_on': 'Add On/Additional Info',
            'importer': 'Importer Name',
            'supplier': 'Supplier Name',
            'origin_country': 'Origin Country',
            'unit_price': 'Unit Price',
            'quantity': 'Quantity',
            'incoterms': 'Incoterms (for auto-read mode)'
        }
        
        # If supplier as sheet is enabled, descriptions stay the same
        # The actual swapping happens during processing, not in the UI labels
        return self.field_descriptions
    
    def on_supplier_as_sheet_change(self):
        """Handle supplier as sheet option change"""
        supplier_mode = self.supplier_as_sheet.get()
        if supplier_mode == "ya":
            self.log_message("Supplier sebagai sheet: YA - Supplier akan menjadi sheet, Importer menjadi kolom")
        else:
            self.log_message("Supplier sebagai sheet: TIDAK - Mode normal (Importer sebagai sheet, Supplier sebagai kolom)")
        
        # Update field descriptions
        self.update_field_descriptions()
    
    def setup_ui(self):
        """Setup the user interface"""
        # Create main notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # File Selection Tab
        self.setup_file_tab()
        
        # Configuration Tab
        self.setup_config_tab()
        
        # Column Mapping Tab
        self.setup_mapping_tab()
        
        # Processing Tab
        self.setup_processing_tab()
        
        # Set initial tab
        self.notebook.select(0)
    
    def setup_file_tab(self):
        """Setup file selection tab"""
        file_frame = ttk.Frame(self.notebook)
        self.notebook.add(file_frame, text="File Selection")
        
        # File selection section
        file_section = ttk.LabelFrame(file_frame, text="Excel File Selection", padding="10")
        file_section.pack(fill='x', padx=10, pady=5)
        
        # Current file display
        ttk.Label(file_section, text="Selected File:").grid(row=0, column=0, sticky='w', pady=2)
        file_label = ttk.Label(file_section, textvariable=self.current_file_path, width=60, relief='sunken')
        file_label.grid(row=0, column=1, columnspan=2, sticky='ew', padx=(10, 0), pady=2)
        
        # Browse button
        browse_btn = ttk.Button(file_section, text="Browse Files", command=self.browse_file)
        browse_btn.grid(row=1, column=0, sticky='w', pady=5)
        
        # Quick select from original_excel folder
        quick_btn = ttk.Button(file_section, text="Select from Original Excel", command=self.quick_select_file)
        quick_btn.grid(row=1, column=1, sticky='w', padx=(10, 0), pady=5)
        
        file_section.columnconfigure(1, weight=1)
        
        # Sheet selection section
        sheet_section = ttk.LabelFrame(file_frame, text="Sheet Selection", padding="10")
        sheet_section.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(sheet_section, text="Select Sheet:").grid(row=0, column=0, sticky='w', pady=2)
        self.sheet_combo = ttk.Combobox(sheet_section, textvariable=self.selected_sheet, state='readonly', width=40)
        self.sheet_combo.grid(row=0, column=1, sticky='ew', padx=(10, 0), pady=2)
        self.sheet_combo.bind('<<ComboboxSelected>>', self.on_sheet_selected)
        
        # Refresh button
        refresh_btn = ttk.Button(sheet_section, text="Refresh", command=self.refresh_sheets)
        refresh_btn.grid(row=0, column=2, padx=(10, 0), pady=2)
        
        sheet_section.columnconfigure(1, weight=1)
        
        # File info section
        self.info_section = ttk.LabelFrame(file_frame, text="File Information", padding="10")
        self.info_section.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Info text widget
        self.info_text = tk.Text(self.info_section, height=10, width=80, wrap='word')
        info_scroll = ttk.Scrollbar(self.info_section, orient='vertical', command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=info_scroll.set)
        
        self.info_text.pack(side='left', fill='both', expand=True)
        info_scroll.pack(side='right', fill='y')
    
    def setup_config_tab(self):
        """Setup configuration tab"""
        config_frame = ttk.Frame(self.notebook)
        self.notebook.add(config_frame, text="Configuration")
        
        # Date format section
        date_section = ttk.LabelFrame(config_frame, text="Date Format", padding="10")
        date_section.pack(fill='x', padx=10, pady=5)
        
        date_formats = [
            ("Auto Detect", "auto"),
            ("DD/MM/YYYY (Indonesian)", "DD/MM/YYYY"),
            ("MM/DD/YYYY (US/Global)", "MM/DD/YYYY"),
            ("DD-MONTH-YYYY (with month names)", "DD-MONTH-YYYY")
        ]
        
        for i, (text, value) in enumerate(date_formats):
            ttk.Radiobutton(date_section, text=text, variable=self.date_format, value=value).grid(
                row=i, column=0, sticky='w', pady=2
            )
        
        # Number format section
        number_section = ttk.LabelFrame(config_frame, text="Number Format", padding="10")
        number_section.pack(fill='x', padx=10, pady=5)
        
        number_formats = [
            ("Auto Detect", "auto"),
            ("American Format (1,234.56)", "american"),
            ("European Format (1.234,56)", "european")
        ]
        
        for i, (text, value) in enumerate(number_formats):
            ttk.Radiobutton(number_section, text=text, variable=self.number_format, value=value).grid(
                row=i, column=0, sticky='w', pady=2
            )
        
        # Other settings section
        other_section = ttk.LabelFrame(config_frame, text="Other Settings", padding="10")
        other_section.pack(fill='x', padx=10, pady=5)
        
        # Year setting
        ttk.Label(other_section, text="Target Year:").grid(row=0, column=0, sticky='w', pady=2)
        year_entry = ttk.Entry(other_section, textvariable=self.target_year, width=10)
        year_entry.grid(row=0, column=1, sticky='w', padx=(10, 0), pady=2)
        
        # INCOTERM setting
        ttk.Label(other_section, text="INCOTERM Mode:").grid(row=1, column=0, sticky='w', pady=2)
        self.incoterm_mode_combo = ttk.Combobox(other_section, textvariable=self.incoterm_mode, 
                                                values=["manual", "from_column"], 
                                                state="readonly", width=15)
        self.incoterm_mode_combo.grid(row=1, column=1, sticky='w', padx=(10, 0), pady=2)
        self.incoterm_mode_combo.bind('<<ComboboxSelected>>', self.on_incoterm_mode_change)
        
        ttk.Label(other_section, text="INCOTERM:").grid(row=2, column=0, sticky='w', pady=2)
        self.incoterm_combo = ttk.Combobox(other_section, textvariable=self.incoterm, 
                                     values=["FOB", "CIF", "CFR", "EXW", "FCA"], width=10)
        self.incoterm_combo.grid(row=2, column=1, sticky='w', padx=(10, 0), pady=2)
        
        # Info label for from_column mode
        self.incoterm_info_label = ttk.Label(other_section, text="(Manual entry - applied to all rows)", 
                                            font=('TkDefaultFont', 8), foreground='gray')
        self.incoterm_info_label.grid(row=2, column=2, sticky='w', padx=(10, 0), pady=2)
        
        # Output filename
        ttk.Label(other_section, text="Output Filename:").grid(row=3, column=0, sticky='w', pady=2)
        output_entry = ttk.Entry(other_section, textvariable=self.output_filename, width=50)
        output_entry.grid(row=3, column=1, columnspan=2, sticky='ew', padx=(10, 0), pady=2)
        
        # Auto-generate filename button
        auto_btn = ttk.Button(other_section, text="Auto Generate", command=self.auto_generate_filename)
        auto_btn.grid(row=3, column=3, padx=(10, 0), pady=2)
        
        # Supplier as sheet option
        ttk.Label(other_section, text="Supplier sebagai Sheet:").grid(row=4, column=0, sticky='w', pady=2)
        
        supplier_frame = ttk.Frame(other_section)
        supplier_frame.grid(row=4, column=1, columnspan=2, sticky='w', padx=(10, 0), pady=2)
        
        supplier_ya_radio = ttk.Radiobutton(supplier_frame, text="Ya", variable=self.supplier_as_sheet, value="ya", command=self.on_supplier_as_sheet_change)
        supplier_ya_radio.pack(side='left')
        supplier_tidak_radio = ttk.Radiobutton(supplier_frame, text="Tidak", variable=self.supplier_as_sheet, value="tidak", command=self.on_supplier_as_sheet_change)
        supplier_tidak_radio.pack(side='left', padx=(10, 0))
        
        # Info label for supplier as sheet option
        supplier_info_label = ttk.Label(other_section, text="Pilih Ya jika Anda membalikkan Supplier dan Importer", 
                                       font=('TkDefaultFont', 8), foreground='gray')
        supplier_info_label.grid(row=5, column=1, columnspan=2, sticky='w', padx=(10, 0), pady=2)
        
        other_section.columnconfigure(1, weight=1)
        
        # Initialize incoterm mode UI state
        self.on_incoterm_mode_change()
    
    def setup_mapping_tab(self):
        """Setup column mapping tab"""
        mapping_frame = ttk.Frame(self.notebook)
        self.notebook.add(mapping_frame, text="Column Mapping")
        
        # Instructions
        instruction_text = ("Map the columns from your Excel file to the required fields. "
                           "Select the appropriate column for each field from the dropdown menus.")
        ttk.Label(mapping_frame, text=instruction_text, wraplength=800).pack(pady=10)
        
        # Mapping section
        mapping_section = ttk.LabelFrame(mapping_frame, text="Column Mappings", padding="10")
        mapping_section.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Create scrollable frame
        canvas = tk.Canvas(mapping_section)
        scrollbar = ttk.Scrollbar(mapping_section, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Field mappings
        self.mapping_widgets = {}
        
        # Update field descriptions based on supplier_as_sheet setting
        field_descriptions = self.update_field_descriptions()
        
        for i, (field_key, description) in enumerate(field_descriptions.items()):
            ttk.Label(scrollable_frame, text=f"{description}:").grid(row=i, column=0, sticky='w', pady=2)
            
            combo = ttk.Combobox(scrollable_frame, textvariable=self.column_mappings[field_key], 
                               state='readonly', width=40)
            combo.grid(row=i, column=1, sticky='ew', padx=(10, 0), pady=2)
            
            self.mapping_widgets[field_key] = combo
        
        scrollable_frame.columnconfigure(1, weight=1)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Control buttons frame
        button_frame = ttk.Frame(mapping_frame)
        button_frame.pack(fill='x', padx=10, pady=10)
        
        # Auto-map button
        auto_map_btn = ttk.Button(button_frame, text="Auto Map Columns", command=self.auto_map_columns)
        auto_map_btn.pack(side='left', padx=(0, 10))
        
        # Refresh columns button
        refresh_map_btn = ttk.Button(button_frame, text="Refresh Columns", command=self.refresh_column_mappings)
        refresh_map_btn.pack(side='left', padx=(0, 10))
        
        # Clear mappings button
        clear_map_btn = ttk.Button(button_frame, text="Clear All", command=self.clear_column_mappings)
        clear_map_btn.pack(side='left', padx=(0, 10))
        
        # Refresh all data button  
        refresh_all_btn = ttk.Button(button_frame, text="Refresh All Data", command=self.refresh_all_data)
        refresh_all_btn.pack(side='left')
        auto_map_btn.pack(pady=10)
    
    def setup_processing_tab(self):
        """Setup processing tab"""
        process_frame = ttk.Frame(self.notebook)
        self.notebook.add(process_frame, text="Processing")
        
        # Status section
        status_section = ttk.LabelFrame(process_frame, text="Processing Status", padding="10")
        status_section.pack(fill='x', padx=10, pady=5)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_section, variable=self.progress_var, mode='determinate')
        self.progress_bar.pack(fill='x', pady=5)
        
        # Status label
        self.status_var = tk.StringVar(value="Ready to process")
        self.status_label = ttk.Label(status_section, textvariable=self.status_var)
        self.status_label.pack(pady=5)
        
        # Control buttons
        button_frame = ttk.Frame(status_section)
        button_frame.pack(fill='x', pady=10)
        
        self.process_btn = ttk.Button(button_frame, text="Start Processing", command=self.start_processing)
        self.process_btn.pack(side='left', padx=(0, 10))
        
        self.cancel_btn = ttk.Button(button_frame, text="Cancel", command=self.cancel_processing, state='disabled')
        self.cancel_btn.pack(side='left')
        
        # Log section
        log_section = ttk.LabelFrame(process_frame, text="Processing Log", padding="10")
        log_section.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Log text widget
        self.log_text = tk.Text(log_section, height=15, width=80, wrap='word')
        log_scroll = ttk.Scrollbar(log_section, orient='vertical', command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scroll.pack(side='right', fill='y')
    
    def browse_file(self):
        """Browse for Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls *.xlsm"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.load_file(file_path)
    
    def quick_select_file(self):
        """Quick select file from original_excel folder"""
        from ..utils.constants import get_app_data_dir
        
        # Get the correct path for both dev and built versions
        if getattr(sys, 'frozen', False):
            # Running as built executable - look in executable directory
            app_dir = os.path.dirname(sys.executable)
            original_excel_path = os.path.join(app_dir, "original_excel")
        else:
            # Running in development
            original_excel_path = Path("original_excel")
        
        if not os.path.exists(original_excel_path):
            # Create the directory if it doesn't exist
            try:
                os.makedirs(original_excel_path, exist_ok=True)
                messagebox.showinfo("Info", f"Created folder: {original_excel_path}\nPlease place your Excel files there and try again.")
            except Exception as e:
                messagebox.showerror("Error", f"Could not create folder {original_excel_path}: {str(e)}")
            return
        
        files = self.js_excel_reader.scan_excel_files(str(original_excel_path))
        
        if not files:
            messagebox.showinfo("Info", f"No Excel files found in {original_excel_path}!\nPlease place Excel files there first.")
            return
        
        # Create file selection dialog
        self.show_file_selection_dialog(files)
    
    def show_file_selection_dialog(self, files):
        """Show file selection dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Excel File")
        dialog.geometry("600x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # File list
        listbox_frame = ttk.Frame(dialog)
        listbox_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        listbox = tk.Listbox(listbox_frame, selectmode='single')
        scrollbar = ttk.Scrollbar(listbox_frame, orient='vertical', command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        
        for file_info in files:
            display_text = f"{file_info['name']} ({file_info['size']} bytes)"
            listbox.insert(tk.END, display_text)
        
        listbox.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill='x', padx=10, pady=10)
        
        def on_select():
            selection = listbox.curselection()
            if selection:
                selected_file = files[selection[0]]
                self.load_file(selected_file['path'])
                dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(button_frame, text="Select", command=on_select).pack(side='right', padx=(10, 0))
        ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side='right')
        
        if files:
            listbox.selection_set(0)  # Select first item
            listbox.focus()
    
    def clear_previous_file_data(self):
        """Clear all cached data from previous file to ensure fresh start"""
        try:
            # Clear current file info
            if hasattr(self, 'current_excel_info'):
                delattr(self, 'current_excel_info')
            
            # Clear sheet selection
            self.selected_sheet.set("")
            
            # Clear available columns
            self.available_columns = []
            
            # Clear all column mappings
            for var in self.column_mappings.values():
                var.set("")
            
            # Update combobox values to be empty
            if hasattr(self, 'mapping_widgets'):
                for widget in self.mapping_widgets.values():
                    widget['values'] = [""]
            
            # Clear sheet combo
            if hasattr(self, 'sheet_combo'):
                self.sheet_combo['values'] = []
            
            # Clear info displays
            if hasattr(self, 'info_text'):
                self.info_text.delete(1.0, tk.END)
            
            if hasattr(self, 'sheet_info_text'):
                self.sheet_info_text.delete(1.0, tk.END)
            
            # Clear output filename to force regeneration
            self.output_filename.set("")
            
            self.log_message("Cleared previous file data - ready for new file")
            
        except Exception as e:
            self.logger.error(f"Error clearing previous file data: {str(e)}")

    def load_file(self, file_path):
        """Load Excel file using JavaScript-style reader"""
        try:
            self.log_message(f"Loading file: {file_path}")
            
            # Clear all cached data from previous file
            self.clear_previous_file_data()
            
            # Get Excel info using JavaScript-style reader
            excel_info = self.js_excel_reader.get_excel_info(file_path)
            
            if excel_info:
                self.current_file_path.set(file_path)
                self.current_excel_info = excel_info
                self.refresh_sheets()
                self.show_file_info()
                self.auto_generate_filename()  # Auto-generate new filename
                self.log_message("File loaded successfully")
            else:
                messagebox.showerror("Error", "Failed to load Excel file!")
                
        except Exception as e:
            self.logger.error(f"Error loading file: {str(e)}")
            messagebox.showerror("Error", f"Error loading file: {str(e)}")
    
    def refresh_sheets(self):
        """Refresh sheet list"""
        if hasattr(self, 'current_excel_info'):
            sheets = self.current_excel_info.get('sheetNames', [])
            self.sheet_combo['values'] = sheets
            
            if sheets:
                self.selected_sheet.set(sheets[0])
                self.log_message(f"Refreshed sheets: {', '.join(sheets)} - Selected: {sheets[0]}")
                self.on_sheet_selected()
            else:
                self.log_message("No sheets found in Excel file")
    
    def on_sheet_selected(self, event=None):
        """Handle sheet selection"""
        if self.selected_sheet.get():
            self.log_message(f"Sheet selected: {self.selected_sheet.get()}")
            self.update_column_mappings()
            self.show_sheet_info()
    
    def update_column_mappings(self):
        """Update column mapping options"""
        sheet_name = self.selected_sheet.get()
        if not sheet_name or not hasattr(self, 'current_excel_info'):
            return
        
        # Clear existing column mappings when switching sheets
        # This prevents confusion with previous sheet's mappings
        for var in self.column_mappings.values():
            var.set("")
        
        # Get column names from selected sheet
        columns = self.js_excel_reader.get_sheet_column_names(
            self.current_file_path.get(), 
            sheet_name
        )
        
        if columns:
            self.available_columns = [""] + columns
            
            # Log column information for debugging
            self.log_message(f"Found {len(columns)} columns in sheet '{sheet_name}'")
            self.log_message(f"Columns: {', '.join(columns[:10])}{'...' if len(columns) > 10 else ''}")
            
            # Update column mapping comboboxes if they exist
            if hasattr(self, 'mapping_widgets'):
                for widget in self.mapping_widgets.values():
                    widget['values'] = self.available_columns
                    
            self.log_message("Column mappings cleared for new sheet - please remap columns")
        else:
            self.log_message(f"No column info found for sheet '{sheet_name}'")
    
    def show_file_info(self):
        """Show file information"""
        self.info_text.delete(1.0, tk.END)
        
        if not self.current_file_path.get() or not hasattr(self, 'current_excel_info'):
            return
        
        file_path = Path(self.current_file_path.get())
        sheets = self.current_excel_info.get('sheetNames', [])
        
        info = f"File: {file_path.name}\n"
        info += f"Path: {file_path}\n"
        info += f"Size: {file_path.stat().st_size:,} bytes\n"
        info += f"Sheets: {len(sheets)}\n"
        info += f"Sheet Names: {', '.join(sheets)}\n"
        
        self.info_text.insert(tk.END, info)
    
    def show_sheet_info(self):
        """Show sheet information"""
        sheet_name = self.selected_sheet.get()
        if not sheet_name or not hasattr(self, 'current_excel_info'):
            return
        
        # Get column names for the selected sheet
        columns = self.js_excel_reader.get_sheet_column_names(
            self.current_file_path.get(), 
            sheet_name
        )
        
        # Update info text
        current_info = self.info_text.get(1.0, tk.END)
        sheet_details = f"\n\nSheet: {sheet_name}\n"
        sheet_details += f"Columns: {len(columns)}\n"
        
        if columns:
            sheet_details += f"\nColumn Names: {', '.join(columns[:10])}{'...' if len(columns) > 10 else ''}\n"
        
        # Clear previous sheet info and append new
        lines = current_info.split('\n')
        # Keep only file info (before sheet info)
        file_info_lines = []
        for line in lines:
            if line.startswith('Sheet:'):
                break
            file_info_lines.append(line)
        
        new_info = '\n'.join(file_info_lines).rstrip() + sheet_details
        
        self.info_text.delete(1.0, tk.END)
        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(tk.END, new_info)
    
    def auto_map_columns(self):
        """Auto-map columns based on common patterns"""
        if not self.available_columns:
            messagebox.showinfo("Info", "Please select a sheet first!")
            return
        
        # Common mapping patterns (more comprehensive)
        mapping_patterns = {
            'date': ['date', 'invoice date', 'shipment date', 'tanggal', 'tgl', 'invoice_date', 'ship_date'],
            'hs_code': ['hs code', 'hs_code', 'hscode', 'commodity code', 'hs', 'harmonized', 'tariff'],
            'item_description': ['item description', 'description', 'product description', 'deskripsi', 'desc', 'commodity'],
            'gsm': ['gsm', 'gram', 'weight', 'gross weight', 'net weight', 'berat'],
            'item': ['item', 'product', 'produk', 'barang', 'goods', 'merchandise'],
            'add_on': ['add on', 'addon', 'additional', 'tambahan', 'remark', 'note', 'keterangan'],
            'importer': ['importer', 'buyer', 'pembeli', 'consignee', 'penerima'],
            'supplier': ['supplier', 'seller', 'penjual', 'exporter', 'shipper', 'pengirim'],
            'origin_country': ['origin', 'country', 'negara', 'asal', 'source', 'from'],
            'unit_price': ['unit price', 'price', 'harga', 'value'],
            'quantity': ['quantity', 'qty', 'jumlah', 'amount']
        }
        
        mapped_count = 0
        
        for field_key, patterns in mapping_patterns.items():
            best_match = None
            best_score = 0
            
            for column in self.available_columns[1:]:  # Skip empty option
                column_lower = column.lower()
                
                for pattern in patterns:
                    if pattern in column_lower:
                        score = len(pattern) / len(column_lower)  # Preference for exact matches
                        if score > best_score:
                            best_score = score
                            best_match = column
            
            if best_match:
                self.column_mappings[field_key].set(best_match)
                mapped_count += 1
        
        self.log_message(f"Auto-mapped {mapped_count} columns")
        messagebox.showinfo("Auto Mapping", f"Successfully mapped {mapped_count} columns!")
    
    def auto_generate_filename(self):
        """Auto-generate output filename"""
        if not self.current_file_path.get():
            messagebox.showinfo("Info", "Please select a file first!")
            return
        
        file_path = Path(self.current_file_path.get())
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # For built versions, show user where file will be saved
        from ..utils.constants import get_safe_output_dir
        output_dir = get_safe_output_dir()
        
        filename = f"Summary_{file_path.stem}_{timestamp}.xlsx"
        self.output_filename.set(filename)
        
        if getattr(sys, 'frozen', False):
            # In built version, inform user of output location
            self.log_message(f"Auto-generated filename: {filename}")
            self.log_message(f"Output will be saved to: {output_dir}")
        else:
            self.log_message(f"Auto-generated filename: {filename}")
    
    def start_processing(self):
        """Start data processing"""
        if self.processing:
            return
        
        # Validate inputs
        if not self.current_file_path.get():
            messagebox.showerror("Error", "Please select an Excel file!")
            return
        
        if not self.selected_sheet.get():
            messagebox.showerror("Error", "Please select a sheet!")
            return
        
        if not self.output_filename.get():
            messagebox.showerror("Error", "Please enter an output filename!")
            return
        
        # Check if at least some columns are mapped
        mapped_columns = sum(1 for var in self.column_mappings.values() if var.get())
        if mapped_columns < 3:
            messagebox.showerror("Error", "Please map at least 3 columns!")
            return
        
        # Start processing in background thread
        self.processing = True
        self.process_btn.config(state='disabled')
        self.cancel_btn.config(state='normal')
        self.progress_var.set(0)
        
        threading.Thread(target=self.process_data, daemon=True).start()
    
    def process_data(self):
        """Process data using JavaScript-style logic"""
        try:
            # Update status
            self.root.after(0, lambda: self.status_var.set("Reading data..."))
            self.root.after(0, lambda: self.progress_var.set(10))
            
            # Get column mappings
            column_mapping = {key: var.get() for key, var in self.column_mappings.items() if var.get()}
            
            # Read data using JavaScript-style reader
            all_raw_data = self.js_excel_reader.read_and_preprocess_data(
                self.current_file_path.get(),
                self.selected_sheet.get(),
                self.date_format.get(),
                self.number_format.get(),
                column_mapping
            )
            
            if not all_raw_data:
                raise ValueError("No data found or failed to read data")
            
            # Validate data structure
            valid_rows = 0
            for i, row in enumerate(all_raw_data):
                if row.get('month') and row.get('hsCode'):
                    valid_rows += 1
                if i < 3:  # Log first 3 rows for debugging
                    self.root.after(0, lambda r=row: self.log_message(f"Sample row: month='{r.get('month')}', hsCode='{r.get('hsCode')}', item='{r.get('item')}'"))
            
            if valid_rows == 0:
                raise ValueError("No valid data rows found (missing month or hsCode)")
            
            self.root.after(0, lambda: self.log_message(f"Read {len(all_raw_data)} rows, {valid_rows} valid for processing"))
            
            # Log sample data for debugging
            if all_raw_data:
                sample_row = all_raw_data[0]
                self.logger.info(f"Sample row: {sample_row}")
            
            self.root.after(0, lambda: self.progress_var.set(30))
            
            # Process data using JavaScript-style processor
            self.root.after(0, lambda: self.status_var.set("Processing data..."))
            
            period_year = self.target_year.get() or str(datetime.now().year)
            global_incoterm = self.incoterm.get() or "FOB"
            incoterm_mode = self.incoterm_mode.get()
            supplier_as_sheet_mode = self.supplier_as_sheet.get()
            output_filename = self.output_filename.get() or "summary_output.xlsx"
            
            self.logger.info(f"Processing with period_year='{period_year}', global_incoterm='{global_incoterm}', incoterm_mode='{incoterm_mode}', supplier_as_sheet='{supplier_as_sheet_mode}', output_filename='{output_filename}'")
            
            self.root.after(0, lambda: self.progress_var.set(60))
            
            # Generate output using JavaScript-style logic
            self.root.after(0, lambda: self.status_var.set("Generating output file..."))
            
            try:
                output_path = self.js_processor.process_data_like_javascript(
                    all_raw_data,
                    period_year,
                    global_incoterm,
                    incoterm_mode,
                    output_filename,
                    supplier_as_sheet_mode
                )
                
                if not output_path:
                    raise ValueError("Processing completed but no output file path was returned")
                    
            except Exception as processing_error:
                raise ValueError(f"Processing failed: {str(processing_error)}")
            
            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.status_var.set("Processing completed successfully!"))
            self.root.after(0, lambda: self.log_message(f"Output saved to: {output_path}"))
            
            # Show success message with file location
            def show_success():
                from ..utils.constants import get_safe_output_dir
                output_dir = get_safe_output_dir()
                success_msg = f"Processing completed successfully!\n\nOutput saved to:\n{output_path}\n\nOutput folder:\n{output_dir}"
                
                # Ask user if they want to open the output folder
                result = messagebox.askyesno(
                    "Success", 
                    success_msg + "\n\nWould you like to open the output folder?",
                    icon='question'
                )
                
                if result:
                    try:
                        # Open the output folder
                        import subprocess
                        subprocess.Popen(f'explorer "{os.path.dirname(output_path)}"')
                    except Exception as e:
                        self.log_message(f"Could not open output folder: {str(e)}")
            
            self.root.after(0, show_success)
            
        except Exception as e:
            error_msg = str(e)
            self.logger.error(f"Processing error: {error_msg}")
            self.root.after(0, lambda: self.status_var.set(f"Error: {error_msg}"))
            self.root.after(0, lambda: self.log_message(f"ERROR: {error_msg}"))
            self.root.after(0, lambda: messagebox.showerror("Error", f"Processing failed: {error_msg}"))
        
        finally:
            self.processing = False
            self.root.after(0, lambda: self.process_btn.config(state='normal'))
            self.root.after(0, lambda: self.cancel_btn.config(state='disabled'))
    
    def cancel_processing(self):
        """Cancel processing"""
        self.processing = False
        self.status_var.set("Processing cancelled")
        self.process_btn.config(state='normal')
        self.cancel_btn.config(state='disabled')
        self.log_message("Processing cancelled by user")
    
    def log_message(self, message):
        """Add message to log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.logger.info(message)
    
    def refresh_column_mappings(self):
        """Refresh column mappings from the currently selected sheet"""
        if not self.selected_sheet.get():
            messagebox.showinfo("Info", "Please select a sheet first!")
            return
        
        try:
            self.log_message("Refreshing column mappings...")
            self.update_column_mappings()
            messagebox.showinfo("Info", f"Column mappings refreshed! Found {len(self.available_columns)-1} columns.")
        except Exception as e:
            self.logger.error(f"Error refreshing column mappings: {str(e)}")
            messagebox.showerror("Error", f"Failed to refresh column mappings: {str(e)}")
    
    def clear_column_mappings(self):
        """Clear all column mappings"""
        for var in self.column_mappings.values():
            var.set("")
        self.log_message("All column mappings cleared")
        messagebox.showinfo("Info", "All column mappings have been cleared!")
    
    def refresh_all_data(self):
        """Manually refresh all data - useful for debugging or when data seems stale"""
        if not self.current_file_path.get():
            messagebox.showinfo("Info", "Please select a file first!")
            return
        
        try:
            current_file = self.current_file_path.get()
            self.log_message("Manually refreshing all data...")
            
            # Reload the file completely
            self.load_file(current_file)
            
            messagebox.showinfo("Info", "All data refreshed successfully!")
            
        except Exception as e:
            self.logger.error(f"Error refreshing all data: {str(e)}")
            messagebox.showerror("Error", f"Failed to refresh data: {str(e)}")
