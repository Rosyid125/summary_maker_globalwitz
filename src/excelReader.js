// src/excelReader.js
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");
const { MONTH_ORDER, getMonthName, parseNumber } = require("./utils");

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

function parseDate(dateString, dateFormat = 'DD/MM/YYYY') {
  if (dateFormat === 'MM/DD/YYYY') {
    return parseDate_MMDDYYYY(dateString);
  } else {
    return parseDate_DDMMYYYY(dateString);
  }
}

function readAndPreprocessData(inputFileName = "input.xlsx", sheetNameToProcess = DEFAULT_SHEET_NAME, dateFormat = 'DD/MM/YYYY', numberFormat = 'EUROPEAN', columnMapping = {}) {
  const inputFile = path.join(DEFAULT_INPUT_FOLDER, inputFileName);
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
        if (typeof dateValue === "string") {
          parsedDateObj = parseDate(dateValue.trim(), dateFormat);
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

module.exports = { readAndPreprocessData };
