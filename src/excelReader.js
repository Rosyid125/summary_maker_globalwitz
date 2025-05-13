// src/excelReader.js
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");
const { MONTH_ORDER, getMonthName, parseNumber } = require("./utils"); // excelSerialNumberToDate mungkin tidak lagi utama

const DEFAULT_INPUT_FOLDER = "original_excel";
const DEFAULT_SHEET_NAME = "DATA OLAH";

function parseDate_DDMMYYYY(dateString) {
  if (typeof dateString !== "string") return null;
  const parts = dateString.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
  if (!parts) return null;

  // parts[0] adalah string asli, parts[1] adalah hari, parts[2] adalah bulan, parts[3] adalah tahun
  const day = parseInt(parts[1], 10);
  const month = parseInt(parts[2], 10); // Bulan adalah 1-12
  let year = parseInt(parts[3], 10);

  // Heuristik untuk tahun 2 digit (misalnya '24' menjadi '2024')
  if (year < 100) {
    year += year > 50 ? 1900 : 2000; // Asumsi: >50 adalah abad 20, <=50 adalah abad 21
  }

  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
  if (month < 1 || month > 12 || day < 1 || day > 31) return null; // Validasi dasar

  // JavaScript Date constructor menggunakan bulan 0-11
  try {
    const dateObj = new Date(year, month - 1, day);
    // Periksa apakah tanggal yang dibuat valid (misalnya 31 Feb menjadi 3 Mar)
    if (dateObj.getFullYear() === year && dateObj.getMonth() === month - 1 && dateObj.getDate() === day) {
      return dateObj;
    }
    return null;
  } catch (e) {
    return null;
  }
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
    // Penting: Jika Excel menyimpan tanggal dd/mm/yyyy sebagai string, raw:false aman.
    // Jika Excel menyimpannya sebagai angka (nomor seri tanggal), maka raw:true mungkin diperlukan untuk mendapatkan string aslinya,
    // TAPI jika formatnya konsisten dd/mm/yyyy, `raw:false` seharusnya memberikan string tanggalnya.
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: null });

    console.log(`Membaca ${jsonData.length} baris dari sheet "${sheetNameToProcess}"...`);

    return jsonData.map((row, rowIndex) => {
      const dateValue = row["DATE"];
      let month = "N/A";

      if (dateValue !== null && typeof dateValue !== "undefined") {
        let parsedDateObj = null;

        if (typeof dateValue === "string") {
          // Prioritaskan parsing dd/mm/yyyy
          parsedDateObj = parseDate_DDMMYYYY(dateValue.trim());

          // Jika gagal dan formatnya YYYYMM (sebagai fallback, jika mungkin masih ada)
          if (!parsedDateObj && /^\d{6}$/.test(dateValue.trim())) {
            const year = parseInt(dateValue.substring(0, 4));
            const monthNum = parseInt(dateValue.substring(4, 6));
            if (monthNum >= 1 && monthNum <= 12 && year >= 1900 && year <= 2100) {
              parsedDateObj = new Date(year, monthNum - 1, 1);
            }
          }
          // Fallback lain jika diperlukan, misalnya Date.parse untuk format US
          // else if (!parsedDateObj) {
          // const timestamp = Date.parse(dateValue.trim());
          // if (!isNaN(timestamp)) {
          // parsedDateObj = new Date(timestamp);
          // }
          // }
        }
        // Jika Excel secara internal menyimpan sebagai angka dan `raw:false` mengembalikannya sebagai angka
        else if (typeof dateValue === "number" && dateValue > 1 && dateValue < 2958465) {
          // Jika ternyata tetap ada nomor seri Excel, kita masih bisa menanganinya
          // const { excelSerialNumberToDate } = require('./utils'); // Impor jika perlu
          // parsedDateObj = excelSerialNumberToDate(dateValue);
          console.warn(`Baris ${rowIndex + 2}: Kolom DATE adalah angka (${dateValue}), diharapkan dd/mm/yyyy. Coba di-parse sebagai nomor seri Excel.`);
          // Anda mungkin ingin membuang excelSerialNumberToDate dari utils jika benar-benar tidak diperlukan lagi
        }

        if (parsedDateObj && !isNaN(parsedDateObj.getTime())) {
          month = getMonthName(parsedDateObj);
        }
      }

      if (month === "N/A" && dateValue) {
        console.warn(`Baris ${rowIndex + 2}: Tidak dapat memparsing DATE "${dateValue}" menjadi bulan (format diharapkan dd/mm/yyyy).`);
      }

      return {
        month: month,
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
