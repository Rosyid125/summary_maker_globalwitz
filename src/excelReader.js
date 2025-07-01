// src/excelReader.js
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");
const { MONTH_ORDER, getMonthName, parseNumber, excelSerialNumberToDate } = require("./utils");

const DEFAULT_INPUT_FOLDER = "original_excel";
const DEFAULT_SHEET_NAME = "DATA OLAH";

// function parseDate_DDMMYYYY(dateString) {
//   if (typeof dateString !== "string") return null;
//   const parts = dateString.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
//   if (!parts) return null;

//   const day = parseInt(parts[1], 10);
//   const month = parseInt(parts[2], 10);
//   let year = parseInt(parts[3], 10);

//   if (year < 100) {
//     year += year > 50 ? 1900 : 2000;
//   }

//   if (isNaN(day) || isNaN(month) || isNaN(year) || month < 1 || month > 12 || day < 1 || day > 31) return null;

//   try {
//     const dateObj = new Date(year, month - 1, day);
//     if (dateObj.getFullYear() === year && dateObj.getMonth() === month - 1 && dateObj.getDate() === day) {
//       return dateObj;
//     }
//     return null;
//   } catch (e) {
//     return null;
//   }
// }

function parseDate_DDMMYYYY(dateString) {
  if (typeof dateString !== "string") return null;

  // Format DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY (STANDAR IMPORT INDOENESIA)
  let parts = dateString.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
  if (parts) {
    const day = parseInt(parts[1], 10);
    const month = parseInt(parts[2], 10);
    let year = parseInt(parts[3], 10);

    if (year < 100) {
      year += year > 50 ? 1900 : 2000;
    }

    if (!isNaN(day) && !isNaN(month) && !isNaN(year) && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const dateObj = new Date(year, month - 1, day);
      if (dateObj.getFullYear() === year && dateObj.getMonth() === month - 1 && dateObj.getDate() === day) {
        return dateObj;
      }
    }
  }

  // Format YYYY-MM-DD (KHUSUS UNTUK IMPORT VIETNAM)
  parts = dateString.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (parts) {
    const year = parseInt(parts[1], 10);
    const month = parseInt(parts[2], 10);
    const day = parseInt(parts[3], 10);

    if (!isNaN(year) && !isNaN(month) && !isNaN(day) && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const dateObj = new Date(year, month - 1, day);
      if (dateObj.getFullYear() === year && dateObj.getMonth() === month - 1 && dateObj.getDate() === day) {
        return dateObj;
      }
    }
  }

  return null;
}

function parseDate_MMDDYYYY(dateString) {
  if (typeof dateString !== "string") return null;

  // Format MM/DD/YYYY, MM-DD-YYYY, MM.DD.YYYY (STANDAR IMPORT USA/GLOBAL)
  let parts = dateString.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
  if (parts) {
    const month = parseInt(parts[1], 10);
    const day = parseInt(parts[2], 10);
    let year = parseInt(parts[3], 10);

    if (year < 100) {
      year += year > 50 ? 1900 : 2000;
    }

    if (!isNaN(day) && !isNaN(month) && !isNaN(year) && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const dateObj = new Date(year, month - 1, day);
      if (dateObj.getFullYear() === year && dateObj.getMonth() === month - 1 && dateObj.getDate() === day) {
        return dateObj;
      }
    }
  }

  // Format YYYY-MM-DD (STANDAR ISO)
  parts = dateString.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (parts) {
    const year = parseInt(parts[1], 10);
    const month = parseInt(parts[2], 10);
    const day = parseInt(parts[3], 10);

    if (!isNaN(year) && !isNaN(month) && !isNaN(day) && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const dateObj = new Date(year, month - 1, day);
      if (dateObj.getFullYear() === year && dateObj.getMonth() === month - 1 && dateObj.getDate() === day) {
        return dateObj;
      }
    }
  }

  return null;
}

function parseDate_DDMONTHYYYY(dateString) {
  if (typeof dateString !== "string") return null;

  // Format DD-Month-YYYY atau DD/Month/YYYY atau DD.Month.YYYY
  // Contoh: "01-mei-2025", "25-jan-2025", "15/feb/2024", "30.des.2023"
  const monthNames = {
    'jan': 0, 'januari': 0,
    'feb': 1, 'februari': 1,
    'mar': 2, 'maret': 2,
    'apr': 3, 'april': 3,
    'mei': 4, 'may': 4,
    'jun': 5, 'juni': 5,
    'jul': 6, 'juli': 6,
    'agu': 7, 'agustus': 7, 'aug': 7, 'august': 7,
    'sep': 8, 'september': 8,
    'okt': 9, 'oktober': 9, 'oct': 9, 'october': 9,
    'nov': 10, 'november': 10,
    'des': 11, 'desember': 11, 'dec': 11, 'december': 11
  };

  const parts = dateString.match(/(\d{1,2})[\/\-\.]([a-zA-Z]+)[\/\-\.](\d{2,4})/);
  if (parts) {
    const day = parseInt(parts[1], 10);
    const monthStr = parts[2].toLowerCase();
    let year = parseInt(parts[3], 10);

    if (year < 100) {
      year += year > 50 ? 1900 : 2000;
    }

    const monthIndex = monthNames[monthStr];
    if (monthIndex !== undefined && !isNaN(day) && !isNaN(year) && day >= 1 && day <= 31) {
      const dateObj = new Date(year, monthIndex, day);
      if (dateObj.getFullYear() === year && dateObj.getMonth() === monthIndex && dateObj.getDate() === day) {
        return dateObj;
      }
    }
  }

  return null;
}

function parseDate(dateString, dateFormat = 'DD/MM/YYYY') {
  if (dateFormat === 'MM/DD/YYYY') {
    return parseDate_MMDDYYYY(dateString);
  } else if (dateFormat === 'DD-MONTH-YYYY') {
    return parseDate_DDMONTHYYYY(dateString);
  } else {
    return parseDate_DDMMYYYY(dateString);
  }
}

function readAndPreprocessData(inputFileNameOrPath = "input.xlsx", sheetNameToProcess = DEFAULT_SHEET_NAME, dateFormat = 'DD/MM/YYYY', numberFormat = 'EUROPEAN', columnMapping = {}) {
  let inputFile;
  
  // Check if it's already a full path or just a filename
  if (path.isAbsolute(inputFileNameOrPath) || inputFileNameOrPath.includes(path.sep)) {
    inputFile = inputFileNameOrPath;
  } else {
    inputFile = path.join(DEFAULT_INPUT_FOLDER, inputFileNameOrPath);
  }
  
  if (!fs.existsSync(inputFile)) {
    console.error(`Error: File input "${inputFile}" tidak ditemukan.`);
    return null;
  }

  try {
    const workbook = XLSX.readFile(inputFile);
    if (!workbook.SheetNames.includes(sheetNameToProcess)) {
      console.error(`Error: Sheet "${sheetNameToProcess}" tidak ditemukan di file ${inputFile}`);
      return null;
    }
    const worksheet = workbook.Sheets[sheetNameToProcess];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: null });

    console.log(`Membaca ${jsonData.length} baris dari sheet "${sheetNameToProcess}" dengan format tanggal ${dateFormat} dan format angka ${numberFormat}...`);

    // Helper function untuk mendapatkan nilai dari mapping kolom dengan fallback
    const getColumnValue = (row, mappingKey, defaultColumns) => {
      if (columnMapping[mappingKey] && columnMapping[mappingKey] !== "") {
        return row[columnMapping[mappingKey]];
      }
      // Fallback ke default columns
      for (const col of defaultColumns) {
        if (row[col] !== undefined) {
          return row[col];
        }
      }
      return null;
    };

    return jsonData.map((row, rowIndex) => {
      // Ambil nilai date dengan mapping atau default
      const dateValue = getColumnValue(row, 'date', ["DATE", "CUSTOMS CLEARANCE DATE"]);
      let month = "-";

      if (dateValue !== null && typeof dateValue !== "undefined") {
        let parsedDateObj = null;
        
        // Coba parse sebagai Excel serial number terlebih dahulu (format seperti 45658)
        // Excel menyimpan tanggal sebagai angka serial sejak 1 Januari 1900
        // Contoh: 45658 = 1 Januari 2025, 45689 = 1 Februari 2025
        if (typeof dateValue === "number" || (typeof dateValue === "string" && /^\d+$/.test(dateValue.trim()))) {
          const serialNumber = typeof dateValue === "number" ? dateValue : parseFloat(dateValue.trim());
          if (!isNaN(serialNumber) && serialNumber > 0) {
            parsedDateObj = excelSerialNumberToDate(serialNumber);
          }
        }
        
        // Jika belum berhasil, coba parse sebagai string dengan format yang dipilih user
        if (!parsedDateObj && typeof dateValue === "string") {
          parsedDateObj = parseDate(dateValue.trim(), dateFormat);
          
          // Format YYYYMM (6 digit)
          if (!parsedDateObj && /^\d{6}$/.test(dateValue.trim())) {
            const year = parseInt(dateValue.substring(0, 4));
            const monthNum = parseInt(dateValue.substring(4, 6));
            if (monthNum >= 1 && monthNum <= 12 && year >= 1900 && year <= 2100) {
              parsedDateObj = new Date(year, monthNum - 1, 1);
            }
          }
        }
        
        if (parsedDateObj && !isNaN(parsedDateObj.getTime())) {
          month = getMonthName(parsedDateObj);
        }
      }

      // if (month === '-' && dateValue) {
      //     console.warn(`Baris ${rowIndex + 2}: Tidak dapat memparsing DATE "${dateValue}" menjadi bulan (format diharapkan dd/mm/yyyy).`);
      // }      // REVISI: Normalisasi ITEM, GSM, ADD ON ke "-" jika kosong
      const getItemValue = (val) => (val === null || typeof val === "undefined" || String(val).trim() === "" ? "-" : String(val).trim());

      return {
        month: month,
        hsCode: String(getColumnValue(row, 'hsCode', ["HS CODE"]) || "-").trim(),
        itemDesc: getItemValue(getColumnValue(row, 'itemDesc', ["ITEM DESC", "PRODUCT DESCRIPTION(EN)"])),
        gsm: getItemValue(getColumnValue(row, 'gsm', ["GSM"])),
        item: getItemValue(getColumnValue(row, 'item', ["ITEM"])),
        addOn: getItemValue(getColumnValue(row, 'addOn', ["ADD ON"])),
        importer: String(getColumnValue(row, 'importer', ["IMPORTER", "PURCHASER"]) || "").trim(),
        supplier: String(getColumnValue(row, 'supplier', ["SUPPLIER"]) || "").trim(),
        originCountry: String(getColumnValue(row, 'originCountry', ["ORIGIN COUNTRY"]) || "-").trim(),
        usdQtyUnit: parseNumber(getColumnValue(row, 'unitPrice', ["CIF KG Unit In USD", "USD Qty Unit", "UNIT PRICE(USD)"]), numberFormat),
        qty: parseNumber(getColumnValue(row, 'quantity', ["Net KG Wt", "qty", "BUSINESS QUANTITY (KG)"]), numberFormat),
      };
    });
  } catch (error) {
    console.error(`Error saat membaca file Excel "${inputFile}":`, error);
    return null;
  }
}

// Helper functions untuk membaca informasi struktur Excel
function getExcelInfo(inputFileNameOrPath = "input.xlsx") {
  let inputFile;
  
  // Check if it's already a full path or just a filename
  if (path.isAbsolute(inputFileNameOrPath) || inputFileNameOrPath.includes(path.sep)) {
    inputFile = inputFileNameOrPath;
  } else {
    inputFile = path.join(DEFAULT_INPUT_FOLDER, inputFileNameOrPath);
  }
  
  if (!fs.existsSync(inputFile)) {
    console.error(`Error: File input "${inputFile}" tidak ditemukan.`);
    return null;
  }

  try {
    const workbook = XLSX.readFile(inputFile);
    const sheetNames = workbook.SheetNames;
    
    // Untuk mendapatkan column names, kita baca sheet pertama sebagai default
    let columnNames = [];
    if (sheetNames.length > 0) {
      const firstSheet = workbook.Sheets[sheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false });
      if (jsonData.length > 0) {
        columnNames = jsonData[0].filter(col => col !== null && col !== undefined && col !== "");
      }
    }

    return {
      sheetNames,
      columnNames
    };
  } catch (error) {
    console.error(`Error saat membaca struktur file Excel "${inputFile}":`, error);
    return null;
  }
}

function getSheetColumnNames(inputFileNameOrPath, sheetName) {
  let inputFile;
  
  // Check if it's already a full path or just a filename
  if (path.isAbsolute(inputFileNameOrPath) || inputFileNameOrPath.includes(path.sep)) {
    inputFile = inputFileNameOrPath;
  } else {
    inputFile = path.join(DEFAULT_INPUT_FOLDER, inputFileNameOrPath);
  }
  
  if (!fs.existsSync(inputFile)) {
    return [];
  }

  try {
    const workbook = XLSX.readFile(inputFile);
    if (!workbook.SheetNames.includes(sheetName)) {
      return [];
    }
    
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
    if (jsonData.length > 0) {
      return jsonData[0].filter(col => col !== null && col !== undefined && col !== "");
    }
    return [];
  } catch (error) {
    console.error(`Error saat membaca kolom dari sheet "${sheetName}":`, error);
    return [];
  }
}

/**
 * Scan dan daftar semua file Excel di folder input
 * @param {string} inputFolder - Folder tempat file input berada
 * @returns {Array} Array berisi informasi file Excel yang ditemukan
 */
function scanExcelFiles(inputFolder = DEFAULT_INPUT_FOLDER) {
  try {
    const folderPath = path.resolve(inputFolder);
    
    if (!fs.existsSync(folderPath)) {
      console.log(`Folder ${inputFolder} tidak ditemukan.`);
      return [];
    }

    const files = fs.readdirSync(folderPath);
    const excelFiles = files.filter(file => {
      const ext = path.extname(file).toLowerCase();
      return ext === '.xlsx' || ext === '.xls';
    });

    return excelFiles.map(file => {
      const filePath = path.join(folderPath, file);
      const stats = fs.statSync(filePath);
      
      return {
        name: file,
        path: filePath,
        size: (stats.size / 1024 / 1024).toFixed(2) + ' MB',
        modified: stats.mtime.toLocaleDateString('id-ID')
      };
    });
  } catch (error) {
    console.error(`Error scanning folder ${inputFolder}:`, error.message);
    return [];
  }
}

module.exports = { readAndPreprocessData, getExcelInfo, getSheetColumnNames, scanExcelFiles };
