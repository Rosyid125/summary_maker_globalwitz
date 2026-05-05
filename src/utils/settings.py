"""
Settings module for managing default column mappings configuration.
Stores user-defined default column names that should be automatically mapped
when loading Excel files.
"""

import json
import os
from pathlib import Path
from typing import Dict, List, Optional


class SettingsManager:
    """Manages application settings including default column mappings."""
    
    # Default configuration file path
    CONFIG_FILENAME = "app_settings.json"
    
    # Field keys that can have default mappings
    MAPPING_FIELDS = [
        'date', 'hs_code', 'item_description', 'gsm', 'item', 'add_on',
        'denier', 'length', 'lustre', 'importer', 'supplier', 
        'origin_country', 'unit_price', 'quantity', 'incoterms'
    ]
    
    def __init__(self, app_data_dir: Optional[str] = None):
        """
        Initialize settings manager.
        
        Args:
            app_data_dir: Directory to store settings file. If None, uses current directory.
        """
        if app_data_dir is None:
            # Check if running as frozen executable
            if getattr(__import__('sys'), 'frozen', False):
                app_data_dir = Path(__import__('sys').executable).parent
            else:
                app_data_dir = Path.cwd()
        
        self.config_path = Path(app_data_dir) / self.CONFIG_FILENAME
        self.settings = self._load_settings()
    
    def _load_settings(self) -> Dict:
        """Load settings from JSON file or return defaults."""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    # Ensure default_mappings exists
                    if 'default_mappings' not in loaded:
                        loaded['default_mappings'] = {}
                    return loaded
            except (json.JSONDecodeError, IOError) as e:
                print(f"Error loading settings: {e}. Using defaults.")
        
        # Return default structure
        return {
            'default_mappings': {},
            'auto_apply_mappings': False,  # New setting: auto-apply when loading files
            'version': '1.0'
        }
    
    def save_settings(self) -> bool:
        """Save current settings to JSON file."""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=2, ensure_ascii=False)
            return True
        except IOError as e:
            print(f"Error saving settings: {e}")
            return False
    
    def get_default_mappings(self) -> Dict[str, List[str]]:
        """
        Get all default column mappings.
        
        Returns:
            Dictionary mapping field keys to lists of default column names.
            Example: {'date': ['Date', 'Invoice Date'], 'hs_code': ['HS Code', 'HS_Code']}
        """
        return self.settings.get('default_mappings', {})
    
    def get_default_mapping_for_field(self, field_key: str) -> List[str]:
        """
        Get default column names for a specific field.
        
        Args:
            field_key: The field key (e.g., 'date', 'hs_code')
            
        Returns:
            List of default column names for the field.
        """
        mappings = self.get_default_mappings()
        return mappings.get(field_key, [])
    
    def set_default_mapping(self, field_key: str, column_names: List[str]) -> bool:
        """
        Set default column names for a field.
        
        Args:
            field_key: The field key to set mapping for.
            column_names: List of column names (will be stripped and filtered).
            
        Returns:
            True if successful.
        """
        if field_key not in self.MAPPING_FIELDS:
            return False
        
        # Clean up the column names (strip whitespace, remove empty)
        cleaned_names = [
            name.strip() 
            for name in column_names 
            if name and name.strip()
        ]
        
        if 'default_mappings' not in self.settings:
            self.settings['default_mappings'] = {}
        
        if cleaned_names:
            self.settings['default_mappings'][field_key] = cleaned_names
        else:
            # Remove the key if no valid names provided
            self.settings['default_mappings'].pop(field_key, None)
        
        return True
    
    def set_all_default_mappings(self, mappings: Dict[str, List[str]]) -> bool:
        """
        Set all default mappings at once.
        
        Args:
            mappings: Dictionary mapping field keys to lists of column names.
            
        Returns:
            True if successful.
        """
        self.settings['default_mappings'] = {}
        
        for field_key, column_names in mappings.items():
            self.set_default_mapping(field_key, column_names)
        
        return True
    
    def find_matching_column(self, field_key: str, available_columns: List[str]) -> Optional[str]:
        """
        Find the first matching column from available columns for a field.
        
        Args:
            field_key: The field to find mapping for.
            available_columns: List of column names available in the Excel file.
            
        Returns:
            The matching column name, or None if no match found.
        """
        default_names = self.get_default_mapping_for_field(field_key)
        
        # Convert available columns to lowercase for case-insensitive comparison
        available_lower = {col: col.lower() for col in available_columns}
        
        for default_name in default_names:
            default_lower = default_name.lower()
            
            # Try exact match first
            for col, col_lower in available_lower.items():
                if default_lower == col_lower:
                    return col
            
            # Try substring match
            for col, col_lower in available_lower.items():
                if default_lower in col_lower or col_lower in default_lower:
                    return col
        
        return None
    
    def clear_all_mappings(self) -> bool:
        """Clear all default mappings."""
        self.settings['default_mappings'] = {}
        return True
    
    def get_auto_apply_mappings(self) -> bool:
        """Get whether to auto-apply default mappings when loading files."""
        return self.settings.get('auto_apply_mappings', False)
    
    def set_auto_apply_mappings(self, enabled: bool) -> bool:
        """Set whether to auto-apply default mappings when loading files."""
        self.settings['auto_apply_mappings'] = enabled
        return True
    
    def export_mappings(self, filepath: str) -> bool:
        """Export mappings to a JSON file."""
        try:
            export_data = {
                'default_mappings': self.get_default_mappings(),
                'version': self.settings.get('version', '1.0')
            }
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, indent=2, ensure_ascii=False)
            return True
        except IOError as e:
            print(f"Error exporting mappings: {e}")
            return False
    
    def import_mappings(self, filepath: str) -> bool:
        """Import mappings from a JSON file."""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                import_data = json.load(f)
            
            if 'default_mappings' in import_data:
                self.settings['default_mappings'] = import_data['default_mappings']
                return True
            return False
        except (json.JSONDecodeError, IOError) as e:
            print(f"Error importing mappings: {e}")
            return False


# Singleton instance for global access
_settings_manager = None

def get_settings_manager(app_data_dir: Optional[str] = None) -> SettingsManager:
    """Get or create the global settings manager instance."""
    global _settings_manager
    if _settings_manager is None:
        _settings_manager = SettingsManager(app_data_dir)
    return _settings_manager
