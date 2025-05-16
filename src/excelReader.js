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

function readAndPreprocessData(inputFileName = "input.xlsx", sheetNameToProcess = DEFAULT_SHEET_NAME) {
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

    console.log(`Membaca ${jsonData.length} baris dari sheet "${sheetNameToProcess}"...`);

    return jsonData.map((row, rowIndex) => {
      const dateValue = row["DATE"];
      let month = "-";

      if (dateValue !== null && typeof dateValue !== "undefined") {
        let parsedDateObj = null;
        if (typeof dateValue === "string") {
          parsedDateObj = parseDate_DDMMYYYY(dateValue.trim());
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
      // }

      // REVISI: Normalisasi ITEM, GSM, ADD ON ke "-" jika kosong
      const getItemValue = (val) => (val === null || typeof val === "undefined" || String(val).trim() === "" ? "-" : String(val).trim());

      return {
        month: month,
        hsCode: String(row["HS CODE"] || "-").trim(), // HS Code tetap '-' jika kosong
        itemDesc: getItemValue(row["ITEM DESC"]), // Gunakan fungsi helper untuk ITEM DESC juga
        gsm: getItemValue(row["GSM"]),
        item: getItemValue(row["ITEM"]),
        addOn: getItemValue(row["ADD ON"]),
        importer: String(row["IMPORTER"] || "").trim(), // String kosong jika tidak ada, ditangani di index.js
        supplier: String(row["SUPPLIER"] || "").trim(), // String kosong jika tidak ada, ditangani di index.js
        originCountry: String(row["ORIGIN COUNTRY"] || "-").trim(),
        usdQtyUnit: parseNumber(row["CIF KG Unit In USD"] || row["USD Qty Unit"]),
        qty: parseNumber(row["Net KG Wt"] || row["qty"]),
      };
    });
  } catch (error) {
    console.error(`Error saat membaca file Excel "${inputFile}":`, error);
    return null;
  }
}

module.exports = { readAndPreprocessData };
