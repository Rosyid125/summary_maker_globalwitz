#!/usr/bin/env python3
"""
Excel Summary Maker - Main Application
GlobalWitz X Volza Programs
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
from pathlib import Path

# Add src directory to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.gui.main_window import MainWindow
from src.utils.logger import setup_logger

class ExcelSummaryMaker:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel Summary Maker - GlobalWitz X Volza")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # Setup logging
        self.logger = setup_logger()
        
        # Initialize main window
        self.main_window = MainWindow(self.root, self.logger)
        
        # Setup window close handler
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def run(self):
        """Start the application"""
        try:
            self.logger.info("Starting Excel Summary Maker application")
            self.root.mainloop()
        except Exception as e:
            self.logger.error(f"Application error: {str(e)}")
            messagebox.showerror("Error", f"Application error: {str(e)}")
    
    def on_closing(self):
        """Handle application closing"""
        try:
            # Add any cleanup here if needed
            self.logger.info("Closing Excel Summary Maker application")
            self.root.destroy()
        except Exception as e:
            self.logger.error(f"Error during application shutdown: {str(e)}")
            self.root.destroy()

def main():
    """Main entry point"""
    try:
        app = ExcelSummaryMaker()
        app.run()
    except Exception as e:
        print(f"Failed to start application: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
