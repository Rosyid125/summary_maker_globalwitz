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
from ..utils.settings import SettingsManager, get_settings_manager

class MainWindow:
    """Main application window"""
    DEFAULT_COMBINATION_MODE_LABEL = "item + gsm + addon"
    FIBER_COMBINATION_MODE_LABEL = "item + gsm + addon + denier + length + lustre"
    
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
        self.combination_mode = tk.StringVar(value=self.DEFAULT_COMBINATION_MODE_LABEL)
        self.output_filename = tk.StringVar()
        
        # Initialize settings manager
        self.settings_manager = get_settings_manager()
        
        # Settings variables for default mappings
        self.default_mapping_vars = {}
        self.auto_apply_mappings = tk.BooleanVar(value=self.settings_manager.get_auto_apply_mappings())
        
        # Column mapping variables
        self.column_mappings = {
            'date': tk.StringVar(),
            'hs_code': tk.StringVar(),
            'item_description': tk.StringVar(),
            'gsm': tk.StringVar(),
            'item': tk.StringVar(),
            'add_on': tk.StringVar(),
            'denier': tk.StringVar(),
            'length': tk.StringVar(),
            'lustre': tk.StringVar(),
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

    def get_combination_mode_value(self):
        """Return normalized combination mode for processing."""
        mode = self.combination_mode.get()
        return "fiber" if mode == self.FIBER_COMBINATION_MODE_LABEL else "default"

    def get_visible_mapping_keys(self):
        """Return mapping keys shown for the selected combination mode."""
        visible_keys = [
            'date', 'hs_code', 'item_description', 'gsm', 'item', 'add_on',
            'importer', 'supplier', 'origin_country', 'unit_price', 'quantity', 'incoterms'
        ]
        if self.get_combination_mode_value() == "fiber":
            insert_at = visible_keys.index('add_on') + 1
            visible_keys[insert_at:insert_at] = ['denier', 'length', 'lustre']
        return visible_keys
    
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
        all_field_descriptions = {
            'date': 'Date/Invoice Date',
            'hs_code': 'HS Code',
            'item_description': 'Item Description',
            'gsm': 'GSM (grams per square meter)',
            'item': 'Item/Product Name',
            'add_on': 'Add On/Additional Info',
            'denier': 'Denier',
            'length': 'Length',
            'lustre': 'Lustre',
            'importer': 'Importer Name',
            'supplier': 'Supplier Name',
            'origin_country': 'Origin Country',
            'unit_price': 'Unit Price',
            'quantity': 'Quantity',
            'incoterms': 'Incoterms (for auto-read mode)'
        }
        self.field_descriptions = {
            key: all_field_descriptions[key]
            for key in self.get_visible_mapping_keys()
        }
        
        # If supplier as sheet is enabled, descriptions stay the same
        # The actual swapping happens during processing, not in the UI labels
        return self.field_descriptions

    def on_combination_mode_change(self, event=None):
        """Handle combination mode change and rebuild mapping controls."""
        mode_value = self.get_combination_mode_value()
        self.rebuild_column_mapping_fields()
        if mode_value == "fiber":
            self.log_message("Combination Mode: fiber - Denier, Length, and Lustre are included in combinations")
        else:
            self.log_message("Combination Mode: default - Item, GSM, and Add On are used for combinations")

    def rebuild_column_mapping_fields(self):
        """Rebuild visible mapping widgets while preserving StringVar values."""
        if not hasattr(self, 'mapping_scrollable_frame'):
            return
        for child in self.mapping_scrollable_frame.winfo_children():
            child.destroy()
        self.mapping_widgets = {}
        field_descriptions = self.update_field_descriptions()
        for i, (field_key, description) in enumerate(field_descriptions.items()):
            ttk.Label(self.mapping_scrollable_frame, text=f"{description}:").grid(row=i, column=0, sticky='w', pady=2)
            combo = ttk.Combobox(
                self.mapping_scrollable_frame,
                textvariable=self.column_mappings[field_key],
                state='readonly',
                width=40
            )
            combo['values'] = self.available_columns if self.available_columns else [""]
            combo.grid(row=i, column=1, sticky='ew', padx=(10, 0), pady=2)
            self.mapping_widgets[field_key] = combo
        self.mapping_scrollable_frame.columnconfigure(1, weight=1)
    
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
        # Create header frame for title and settings button
        self.setup_header_frame()
        
        # Create main notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # File Selection Tab
        self.setup_file_tab()
        
        # Configuration Tab
        self.setup_config_tab()
        
        # Column Mapping Tab
        self.setup_mapping_tab()
        
        # Processing Tab
        self.setup_processing_tab()
        
        # Settings is now a separate dialog, not a tab
        self.setup_settings_dialog()
        
        # Set initial tab
        self.notebook.select(0)
    
    def setup_header_frame(self):
        """Setup header frame with title and settings button"""
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill='x', padx=10, pady=(10, 0))
        
        # Application title on the left
        title_label = ttk.Label(
            header_frame, 
            text="Excel Summary Maker", 
            font=('TkDefaultFont', 12, 'bold')
        )
        title_label.pack(side='left')
        
        # Settings button on the right (gear icon)
        settings_btn = ttk.Button(
            header_frame,
            text="⚙ Settings",
            command=self.open_settings_dialog,
            width=12
        )
        settings_btn.pack(side='right')
        
        self.header_frame = header_frame
    
    def open_settings_dialog(self):
        """Open the settings dialog window"""
        if hasattr(self, 'settings_dialog') and self.settings_dialog.winfo_exists():
            # Bring existing window to front
            self.settings_dialog.lift()
            self.settings_dialog.focus_force()
            return
        
        # Create new settings dialog
        self.settings_dialog = tk.Toplevel(self.root)
        self.settings_dialog.title("Settings - Default Column Mappings")
        self.settings_dialog.geometry("700x600")
        self.settings_dialog.transient(self.root)
        self.settings_dialog.grab_set()
        
        # Center the dialog on the main window
        self.settings_dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (700 // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (600 // 2)
        self.settings_dialog.geometry(f"700x600+{x}+{y}")
        
        # Build the settings UI inside the dialog
        self.build_settings_ui(self.settings_dialog)
    
    def build_settings_ui(self, parent):
        """Build the settings UI inside the given parent widget"""
        # Instructions
        instruction_text = ("Configure default column mappings. Enter column names from your Excel files "
                           "(row 1). You can enter multiple names separated by commas for each field. "
                           "The system will try to auto-map these when loading files.")
        ttk.Label(parent, text=instruction_text, wraplength=650).pack(pady=10, padx=10)
        
        # Auto-apply option
        auto_apply_frame = ttk.LabelFrame(parent, text="Auto-Apply Settings", padding="10")
        auto_apply_frame.pack(fill='x', padx=10, pady=5)
        
        auto_apply_cb = ttk.Checkbutton(
            auto_apply_frame, 
            text="Automatically apply default mappings when loading Excel files",
            variable=self.auto_apply_mappings,
            command=self.on_auto_apply_change
        )
        auto_apply_cb.pack(anchor='w')
        
        ttk.Label(auto_apply_frame, 
                 text="When enabled, default mappings will be applied automatically when you select a sheet",
                 font=('TkDefaultFont', 8), foreground='gray').pack(anchor='w', pady=(5, 0))
        
        # Create scrollable frame for default mappings
        mapping_container = ttk.LabelFrame(parent, text="Default Mapping Set", padding="10")
        mapping_container.pack(fill='both', expand=True, padx=10, pady=5)
        
        canvas = tk.Canvas(mapping_container)
        scrollbar = ttk.Scrollbar(mapping_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        self.settings_scrollable_frame = scrollable_frame
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Initialize default mapping variables and create fields
        self._create_default_mapping_fields()
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Control buttons frame
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill='x', padx=10, pady=10)
        
        # Left side buttons
        left_buttons = ttk.Frame(button_frame)
        left_buttons.pack(side='left')
        
        save_btn = ttk.Button(left_buttons, text="Save Settings", command=self.save_default_mappings)
        save_btn.pack(side='left', padx=(0, 5))
        
        load_btn = ttk.Button(left_buttons, text="Load Settings", command=self.load_default_mappings)
        load_btn.pack(side='left', padx=(0, 5))
        
        clear_btn = ttk.Button(left_buttons, text="Clear All", command=self.clear_default_mappings)
        clear_btn.pack(side='left')
        
        # Right side buttons
        right_buttons = ttk.Frame(button_frame)
        right_buttons.pack(side='right')
        
        export_btn = ttk.Button(right_buttons, text="Export", command=self.export_mappings)
        export_btn.pack(side='left', padx=(0, 5))
        
        import_btn = ttk.Button(right_buttons, text="Import", command=self.import_mappings)
        import_btn.pack(side='left', padx=(0, 5))
        
        close_btn = ttk.Button(right_buttons, text="Close", command=self.close_settings_dialog)
        close_btn.pack(side='left')
    
    def close_settings_dialog(self):
        """Close the settings dialog"""
        if hasattr(self, 'settings_dialog') and self.settings_dialog.winfo_exists():
            self.settings_dialog.destroy()
    
    def setup_settings_dialog(self):
        """Initialize settings dialog data (UI is built when dialog is opened)"""
        # Initialize default mapping variables
        self.default_mapping_vars = {}
        # Settings UI will be built when open_settings_dialog() is called
    
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
        ttk.Label(other_section, text="COMBINATION MODE:").grid(row=4, column=0, sticky='w', pady=2)
        self.combination_mode_combo = ttk.Combobox(other_section, textvariable=self.combination_mode,
                                                  values=[self.DEFAULT_COMBINATION_MODE_LABEL, self.FIBER_COMBINATION_MODE_LABEL],
                                                  state="readonly", width=45)
        self.combination_mode_combo.grid(row=4, column=1, columnspan=2, sticky='w', padx=(10, 0), pady=2)
        self.combination_mode_combo.bind('<<ComboboxSelected>>', self.on_combination_mode_change)

        ttk.Label(other_section, text="Supplier sebagai Sheet:").grid(row=5, column=0, sticky='w', pady=2)
        
        supplier_frame = ttk.Frame(other_section)
        supplier_frame.grid(row=5, column=1, columnspan=2, sticky='w', padx=(10, 0), pady=2)
        
        supplier_ya_radio = ttk.Radiobutton(supplier_frame, text="Ya", variable=self.supplier_as_sheet, value="ya", command=self.on_supplier_as_sheet_change)
        supplier_ya_radio.pack(side='left')
        supplier_tidak_radio = ttk.Radiobutton(supplier_frame, text="Tidak", variable=self.supplier_as_sheet, value="tidak", command=self.on_supplier_as_sheet_change)
        supplier_tidak_radio.pack(side='left', padx=(10, 0))
        
        # Info label for supplier as sheet option
        supplier_info_label = ttk.Label(other_section, text="Pilih Ya jika Anda membalikkan Supplier dan Importer", 
                                       font=('TkDefaultFont', 8), foreground='gray')
        supplier_info_label.grid(row=6, column=1, columnspan=2, sticky='w', padx=(10, 0), pady=2)
        
        other_section.columnconfigure(1, weight=1)
        
        # Initialize incoterm mode UI state
        self.on_incoterm_mode_change()
    
    def _create_default_mapping_fields(self):
        """Create input fields for default column mappings"""
        field_descriptions = {
            'date': 'Date/Invoice Date',
            'hs_code': 'HS Code',
            'item_description': 'Item Description',
            'gsm': 'GSM (grams per square meter)',
            'item': 'Item/Product Name',
            'add_on': 'Add On/Additional Info',
            'denier': 'Denier',
            'length': 'Length',
            'lustre': 'Lustre',
            'importer': 'Importer Name',
            'supplier': 'Supplier Name',
            'origin_country': 'Origin Country',
            'unit_price': 'Unit Price',
            'quantity': 'Quantity',
            'incoterms': 'Incoterms (for auto-read mode)'
        }
        
        # Load existing settings
        existing_mappings = self.settings_manager.get_default_mappings()
        
        for i, (field_key, description) in enumerate(field_descriptions.items()):
            # Create label
            ttk.Label(self.settings_scrollable_frame, text=f"{description}:", 
                     font=('TkDefaultFont', 9, 'bold')).grid(
                row=i*2, column=0, sticky='w', pady=(10, 2)
            )
            
            # Create entry for multiple column names (comma-separated)
            var = tk.StringVar()
            
            # Pre-populate with existing settings
            if field_key in existing_mappings:
                var.set(', '.join(existing_mappings[field_key]))
            
            self.default_mapping_vars[field_key] = var
            
            entry = ttk.Entry(self.settings_scrollable_frame, textvariable=var, width=60)
            entry.grid(row=i*2, column=1, sticky='ew', padx=(10, 0), pady=(10, 2))
            
            # Help text
            ttk.Label(self.settings_scrollable_frame, 
                     text="Enter column names separated by commas",
                     font=('TkDefaultFont', 8), foreground='gray').grid(
                row=i*2+1, column=1, sticky='w', padx=(10, 0), pady=(0, 5)
            )
        
        self.settings_scrollable_frame.columnconfigure(1, weight=1)
    
    def save_default_mappings(self):
        """Save default mappings to settings"""
        try:
            mappings = {}
            for field_key, var in self.default_mapping_vars.items():
                value = var.get().strip()
                if value:
                    # Split by comma and clean up
                    column_names = [name.strip() for name in value.split(',') if name.strip()]
                    if column_names:
                        mappings[field_key] = column_names
            
            self.settings_manager.set_all_default_mappings(mappings)
            
            if self.settings_manager.save_settings():
                self.log_message("Default mappings saved successfully!")
                messagebox.showinfo("Success", "Default mappings saved successfully!")
            else:
                messagebox.showerror("Error", "Failed to save settings!")
        except Exception as e:
            self.logger.error(f"Error saving default mappings: {e}")
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
    
    def load_default_mappings(self):
        """Load default mappings from settings"""
        try:
            mappings = self.settings_manager.get_default_mappings()
            
            for field_key, var in self.default_mapping_vars.items():
                if field_key in mappings:
                    var.set(', '.join(mappings[field_key]))
                else:
                    var.set('')
            
            self.log_message("Default mappings loaded successfully!")
            messagebox.showinfo("Success", "Default mappings loaded successfully!")
        except Exception as e:
            self.logger.error(f"Error loading default mappings: {e}")
            messagebox.showerror("Error", f"Failed to load settings: {str(e)}")
    
    def clear_default_mappings(self):
        """Clear all default mappings"""
        if messagebox.askyesno("Confirm", "Are you sure you want to clear all default mappings?"):
            for var in self.default_mapping_vars.values():
                var.set('')
            self.settings_manager.clear_all_mappings()
            self.settings_manager.save_settings()
            self.log_message("Default mappings cleared!")
            messagebox.showinfo("Success", "Default mappings cleared!")
    
    def export_mappings(self):
        """Export mappings to a JSON file"""
        try:
            file_path = filedialog.asksaveasfilename(
                title="Export Mappings",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            if file_path:
                # First save current UI values to settings
                mappings = {}
                for field_key, var in self.default_mapping_vars.items():
                    value = var.get().strip()
                    if value:
                        column_names = [name.strip() for name in value.split(',') if name.strip()]
                        if column_names:
                            mappings[field_key] = column_names
                
                self.settings_manager.set_all_default_mappings(mappings)
                
                if self.settings_manager.export_mappings(file_path):
                    self.log_message(f"Mappings exported to: {file_path}")
                    messagebox.showinfo("Success", f"Mappings exported to:\n{file_path}")
                else:
                    messagebox.showerror("Error", "Failed to export mappings!")
        except Exception as e:
            self.logger.error(f"Error exporting mappings: {e}")
            messagebox.showerror("Error", f"Failed to export mappings: {str(e)}")
    
    def import_mappings(self):
        """Import mappings from a JSON file"""
        try:
            file_path = filedialog.askopenfilename(
                title="Import Mappings",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            if file_path:
                if self.settings_manager.import_mappings(file_path):
                    # Update UI with imported values
                    self.load_default_mappings()
                    self.log_message(f"Mappings imported from: {file_path}")
                    messagebox.showinfo("Success", f"Mappings imported from:\n{file_path}")
                else:
                    messagebox.showerror("Error", "Failed to import mappings!")
        except Exception as e:
            self.logger.error(f"Error importing mappings: {e}")
            messagebox.showerror("Error", f"Failed to import mappings: {str(e)}")
    
    def on_auto_apply_change(self):
        """Handle auto-apply checkbox change"""
        enabled = self.auto_apply_mappings.get()
        self.settings_manager.set_auto_apply_mappings(enabled)
        self.settings_manager.save_settings()
        status = "enabled" if enabled else "disabled"
        self.log_message(f"Auto-apply default mappings {status}")
    
    def apply_default_mappings_auto(self):
        """Auto-apply default mappings when loading a sheet (if enabled)"""
        if not self.auto_apply_mappings.get():
            return
        
        if not self.available_columns:
            return
        
        default_mappings = self.settings_manager.get_default_mappings()
        if not default_mappings:
            return
        
        mapped_count = 0
        visible_mapping_keys = set(self.get_visible_mapping_keys())
        
        for field_key in visible_mapping_keys:
            if field_key not in default_mappings:
                continue
            
            # Skip if already mapped
            if self.column_mappings[field_key].get():
                continue
            
            # Try to find a match
            matched_column = self.settings_manager.find_matching_column(
                field_key, 
                self.available_columns[1:]  # Skip empty option
            )
            
            if matched_column:
                self.column_mappings[field_key].set(matched_column)
                mapped_count += 1
        
        if mapped_count > 0:
            self.log_message(f"Auto-applied {mapped_count} default mappings")
    
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
        self.mapping_scrollable_frame = scrollable_frame
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Field mappings
        self.mapping_widgets = {}
        
        # Update field descriptions based on supplier_as_sheet setting
        self.rebuild_column_mapping_fields()
        
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
            
            # Auto-apply default mappings if enabled
            self.apply_default_mappings_auto()
                    
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
        """Auto-map columns based on default mappings first, then fall back to common patterns"""
        if not self.available_columns:
            messagebox.showinfo("Info", "Please select a sheet first!")
            return
        
        mapped_count = 0
        default_mapped_count = 0
        pattern_mapped_count = 0
        
        visible_mapping_keys = set(self.get_visible_mapping_keys())
        
        # Step 1: Try to match using user-defined default mappings (priority)
        default_mappings = self.settings_manager.get_default_mappings()
        
        for field_key in visible_mapping_keys:
            if field_key not in default_mappings or not default_mappings[field_key]:
                continue
            
            # Try to find a match using default mappings
            matched_column = self.settings_manager.find_matching_column(
                field_key, 
                self.available_columns[1:]  # Skip empty option
            )
            
            if matched_column:
                self.column_mappings[field_key].set(matched_column)
                mapped_count += 1
                default_mapped_count += 1
                self.log_message(f"Default mapping: {field_key} -> '{matched_column}'")
        
        # Step 2: Fall back to common patterns for unmapped fields
        # Common mapping patterns (more comprehensive)
        mapping_patterns = {
            'date': ['date', 'invoice date', 'shipment date', 'tanggal', 'tgl', 'invoice_date', 'ship_date'],
            'hs_code': ['hs code', 'hs_code', 'hscode', 'commodity code', 'hs', 'harmonized', 'tariff'],
            'item_description': ['item description', 'description', 'product description', 'deskripsi', 'desc', 'commodity'],
            'gsm': ['gsm', 'gram', 'weight', 'gross weight', 'net weight', 'berat'],
            'item': ['item', 'product', 'produk', 'barang', 'goods', 'merchandise'],
            'add_on': ['add on', 'addon', 'additional', 'tambahan', 'remark', 'note', 'keterangan'],
            'denier': ['denier', 'den', 'dny'],
            'length': ['length', 'len', 'panjang'],
            'lustre': ['lustre', 'luster', 'kilap'],
            'importer': ['importer', 'buyer', 'pembeli', 'consignee', 'penerima'],
            'supplier': ['supplier', 'seller', 'penjual', 'exporter', 'shipper', 'pengirim'],
            'origin_country': ['origin', 'country', 'negara', 'asal', 'source', 'from'],
            'unit_price': ['unit price', 'price', 'harga', 'value'],
            'quantity': ['quantity', 'qty', 'jumlah', 'amount']
        }
        
        for field_key, patterns in mapping_patterns.items():
            if field_key not in visible_mapping_keys:
                continue
            
            # Skip if already mapped by default mappings
            if self.column_mappings[field_key].get():
                continue
            
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
                pattern_mapped_count += 1
        
        # Log results
        self.log_message(f"Auto-mapped {mapped_count} columns total")
        if default_mapped_count > 0:
            self.log_message(f"  - {default_mapped_count} from user default mappings")
        if pattern_mapped_count > 0:
            self.log_message(f"  - {pattern_mapped_count} from common patterns")
        
        messagebox.showinfo("Auto Mapping", 
                          f"Successfully mapped {mapped_count} columns!\n"
                          f"- {default_mapped_count} from your default mappings\n"
                          f"- {pattern_mapped_count} from common patterns")
    
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
        visible_mapping_keys = self.get_visible_mapping_keys()
        mapped_columns = sum(1 for key in visible_mapping_keys if self.column_mappings[key].get())
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
            visible_mapping_keys = self.get_visible_mapping_keys()
            column_mapping = {key: self.column_mappings[key].get() for key in visible_mapping_keys if self.column_mappings[key].get()}
            combination_mode = self.get_combination_mode_value()
            
            # Read data using JavaScript-style reader
            all_raw_data = self.js_excel_reader.read_and_preprocess_data(
                self.current_file_path.get(),
                self.selected_sheet.get(),
                self.date_format.get(),
                self.number_format.get(),
                column_mapping,
                combination_mode
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
            
            self.logger.info(f"Processing with period_year='{period_year}', global_incoterm='{global_incoterm}', incoterm_mode='{incoterm_mode}', supplier_as_sheet='{supplier_as_sheet_mode}', combination_mode='{combination_mode}', output_filename='{output_filename}'")
            
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
                    supplier_as_sheet_mode,
                    combination_mode
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
