// src/excelReader.js
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");
const { MONTH_ORDER, excelSerialNumberToDate, getMonthName, parseNumber } = require("./utils"); // Pastikan getMonthName sudah benar di utils.js

const DEFAULT_INPUT_FOLDER = "original_excel";
const DEFAULT_SHEET_NAME = "DATA OLAH";

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
      // Tambahkan rowIndex untuk debugging jika perlu
      const dateValue = row["DATE"];
      let month = "N/A"; // Default jika tidak bisa diparsing

      if (dateValue !== null && typeof dateValue !== "undefined") {
        let parsedDateObj = null;

        // 1. Coba parse sebagai nomor seri Excel
        if (typeof dateValue === "number") {
          if (dateValue > 1 && dateValue < 2958465) {
            // Batas wajar untuk nomor seri Excel
            parsedDateObj = excelSerialNumberToDate(dateValue);
          } else if (dateValue.toString().length === 5 && parseInt(dateValue.toString().substring(0, 2)) >= 1 && parseInt(dateValue.toString().substring(0, 2)) <= 12) {
            // Jika formatnya MMYYY (misal 45755, ini tidak standar, mungkin perlu penyesuaian)
            // Ini asumsi, jika format DATE Anda benar-benar hanya 5 digit seperti contoh, perlu klarifikasi lebih lanjut
            // Untuk saat ini, kita akan coba interpretasi sebagai YYYYMM jika panjangnya 6
          }
        }

        // 2. Jika bukan nomor seri atau parsing gagal, coba sebagai string
        if (!parsedDateObj && typeof dateValue === "string") {
          // Coba format YYYYMM (misalnya '202412')
          if (/^\d{6}$/.test(dateValue)) {
            const year = parseInt(dateValue.substring(0, 4));
            const monthNum = parseInt(dateValue.substring(4, 6));
            if (monthNum >= 1 && monthNum <= 12 && year >= 1900 && year <= 2100) {
              // Validasi dasar
              parsedDateObj = new Date(year, monthNum - 1, 1); // Buat objek Date dummy hanya untuk bulan
            }
          } else {
            // Coba format lain yang bisa di-parse Date.parse()
            const timestamp = Date.parse(dateValue);
            if (!isNaN(timestamp)) {
              parsedDateObj = new Date(timestamp);
            }
          }
        }

        // 3. Jika berupa objek Date langsung dari xlsx (jarang, tapi mungkin)
        else if (dateValue instanceof Date && !isNaN(dateValue)) {
          parsedDateObj = dateValue;
        }

        if (parsedDateObj && !isNaN(parsedDateObj.getTime())) {
          month = getMonthName(parsedDateObj);
        }
      }

      // Debugging jika bulan masih N/A
      // if (month === 'N/A' && dateValue) {
      //     console.warn(`Baris ${rowIndex + 2}: Tidak dapat memparsing DATE "${dateValue}" menjadi bulan.`);
      // }

      return {
        month: month, // Akan 'N/A' jika tidak bisa diparsing
        hsCode: String(row["HS CODE"] || "N/A").trim(),
        itemDesc: String(row["ITEM DESC"] || "N/A").trim(),
        gsm: String(row["GSM"] || "N/A").trim(),
        item: String(row["ITEM"] || "N/A").trim(),
        addOn: String(row["ADD ON"] || "N/A").trim(),
        importer: String(row["IMPORTER"] || "").trim(),
        supplier: String(row["SUPPLIER"] || "").trim(),
        originCountry: String(row["ORIGIN COUNTRY"] || "N/A").trim(),
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
