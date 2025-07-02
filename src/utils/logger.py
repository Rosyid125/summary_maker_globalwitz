"""
Logger utility for Excel Summary Maker
"""

import logging
import os
from datetime import datetime
from pathlib import Path

def setup_logger(name="ExcelSummaryMaker", log_level=logging.INFO):
    """
    Setup logger with both console and file handlers
    
    Args:
        name (str): Logger name
        log_level: Logging level
    
    Returns:
        logger: Configured logger instance
    """
    
    # Create logger
    logger = logging.getLogger(name)
    logger.setLevel(log_level)
    
    # Clear existing handlers
    logger.handlers.clear()
    
    # Create formatters
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # File handler
    try:
        # Create logs directory if it doesn't exist
        logs_dir = Path("logs")
        logs_dir.mkdir(exist_ok=True)
        
        # Create log file with timestamp
        log_filename = f"excel_summary_maker_{datetime.now().strftime('%Y%m%d')}.log"
        log_path = logs_dir / log_filename
        
        file_handler = logging.FileHandler(log_path, encoding='utf-8')
        file_handler.setLevel(log_level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        
    except Exception as e:
        logger.warning(f"Could not create file handler: {str(e)}")
    
    return logger
