# Excel Summary Maker - GlobalWitz X Volza

A Python application with Tkinter GUI for processing Excel import/export data and generating comprehensive summaries.

## Features

- **Modern GUI Interface**: User-friendly Tkinter-based interface with tabbed navigation
- **Multiple Excel Format Support**: Reads .xlsx, .xls, and .xlsm files
- **Flexible Date Parsing**: Supports multiple date formats including DD/MM/YYYY, MM/DD/YYYY, and month names
- **Smart Number Parsing**: Handles both American (1,234.56) and European (1.234,56) number formats
- **Intelligent Column Mapping**: Auto-detect and map Excel columns to required fields
- **Multi-level Data Aggregation**: Creates monthly summaries and overall statistics
- **Professional Excel Output**: Generates formatted Excel reports with color-coded quarters
- **Comprehensive Logging**: Detailed logging for troubleshooting and audit trails
- **Smart Data Refresh**: Automatically clears cached data when switching between different Excel files (Fixed v1.1)

## Requirements

- Python 3.7 or higher
- Required packages (install via `pip install -r requirements.txt`):
  - pandas==2.0.3
  - openpyxl==3.1.2
  - python-dateutil==2.8.2
  - xlsxwriter==3.1.9

## Installation

1. Clone or download this repository
2. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Starting the Application

Run the main application:
```bash
python main.py
```

### Using the GUI

The application has 4 main tabs:

#### 1. File Selection
- **Browse Files**: Select any Excel file from your system
- **Select from Original Excel**: Quick select from the `original_excel` folder
- **Sheet Selection**: Choose which sheet to process from the Excel file
- **File Information**: View file details and data preview

#### 2. Configuration
- **Date Format**: Choose how dates are formatted in your Excel file
  - Auto Detect (recommended)
  - DD/MM/YYYY (Indonesian standard)
  - MM/DD/YYYY (US/Global standard)
  - DD-MONTH-YYYY (with month names)
- **Number Format**: Select number formatting
  - Auto Detect (recommended)
  - American Format (1,234.56)
  - European Format (1.234,56)
- **Other Settings**:
  - Target Year: Filter data by specific year
  - INCOTERM: Set pricing terms (FOB, CIF, etc.)
  - Output Filename: Name for the generated report

#### 3. Column Mapping
- Map Excel columns to required fields:
  - Date/Invoice Date
  - HS Code
  - Item Description
  - GSM (grams per square meter)
  - Item/Product Name
  - Add On/Additional Info
  - Importer Name
  - Supplier Name
  - Origin Country
  - Unit Price
  - Quantity
- **Auto Map Columns**: Automatically detect and map columns based on common patterns

#### 4. Processing
- **Start Processing**: Begin data processing
- **Progress Tracking**: Monitor processing progress with progress bar
- **Processing Log**: View detailed processing logs in real-time
- **Cancel**: Stop processing if needed

## Output

The application generates Excel files in the `processed_excel` folder with multiple sheets:

### Overall Summary Sheet
- Summary statistics for all importers
- Total records, quantities, and values
- Unique supplier and item counts

### Individual Importer Sheets
Each importer gets a dedicated sheet with:

1. **Overall Summary**: Key statistics for the importer
2. **Monthly Summary**: Data grouped by month with quarterly color coding
   - Q1 (Jan-Mar): Light Red
   - Q2 (Apr-Jun): Light Blue  
   - Q3 (Jul-Sep): Light Green
   - Q4 (Oct-Dec): Light Yellow
3. **Supplier Summary**: Analysis by supplier with totals and averages
4. **Item Summary**: Analysis by item/product with quantities and values

## File Structure

```
excel_summary_maker/
├── main.py                          # Main application entry point
├── requirements.txt                 # Python dependencies
├── README.md                       # This file
├── original_excel/                 # Input Excel files
│   └── US-Import-jan-jun-2025.xlsx
├── processed_excel/               # Output Excel files
├── logs/                          # Application logs
└── src/                           # Source code
    ├── core/                      # Core processing modules
    │   ├── excel_reader.py        # Excel file reading and parsing
    │   ├── data_aggregator.py     # Data aggregation and summarization
    │   └── output_formatter.py    # Excel output generation
    ├── gui/                       # GUI components
    │   └── main_window.py         # Main application window
    └── utils/                     # Utility modules
        ├── helpers.py             # Date/number parsing utilities
        └── logger.py              # Logging configuration
```

## Data Processing Flow

1. **File Reading**: Load Excel file and scan available sheets
2. **Data Parsing**: Parse dates and numbers according to selected formats
3. **Column Mapping**: Map Excel columns to standardized field names
4. **Data Aggregation**: Create multi-level summaries:
   - Monthly summaries by HS Code + Item + GSM + Add On
   - Overall summaries across all months
   - Supplier-wise summaries
   - Item-wise summaries
5. **Output Generation**: Create formatted Excel report with multiple sheets

## Error Handling

- **File Validation**: Checks for valid Excel files and formats
- **Data Validation**: Validates required fields and data types
- **Error Logging**: Comprehensive error logging with timestamps
- **User Feedback**: Clear error messages and processing status updates

## Troubleshooting

### Common Issues

1. **File Not Loading**: 
   - Ensure the Excel file is not open in another application
   - Check file permissions and format

2. **Column Mapping Issues**:
   - Use the "Auto Map Columns" feature
   - Manually verify column mappings match your data

3. **Date Parsing Problems**:
   - Try different date format options
   - Check if dates are stored as text or numbers in Excel

4. **Number Format Issues**:
   - Select the appropriate number format for your locale
   - Ensure numeric columns don't contain text

### Log Files

Check the `logs/` folder for detailed error logs:
- `excel_summary_maker_YYYYMMDD.log`

## Technical Details

### Supported Date Formats
- Excel serial numbers
- DD/MM/YYYY, MM/DD/YYYY formats
- DD-MONTH-YYYY with month names (English/Indonesian)
- ISO format (YYYY-MM-DD)
- YYYYMM format (6-digit)

### Number Format Detection
- Automatic detection of thousands separators and decimal points
- Support for both comma and period as decimal separators
- Currency symbol removal

### Excel Output Features
- Professional formatting with borders and colors
- Quarterly color coding for easy visual analysis
- Auto-adjusted column widths
- Multiple sheets per importer
- Summary statistics and totals

## License

This project is developed for GlobalWitz X Volza Programs.

## Support

For technical support or feature requests, contact the development team.
