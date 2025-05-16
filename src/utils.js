// src/utils.js
const MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];

function excelSerialNumberToDate(serial) {
  if (typeof serial !== "number" || isNaN(serial)) return null;
  // Pastikan serial berada dalam rentang yang wajar untuk tanggal Excel
  if (serial < 1 || serial > 2958465) {
    // 2958465 adalah untuk 31/12/9999
    // console.warn(`Nomor seri Excel di luar rentang wajar: ${serial}`);
    return null;
  }
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  if (isNaN(date_info.getTime())) return null;

  const fractional_day = serial - Math.floor(serial) + 0.0000001;
  let total_seconds = Math.floor(86400 * fractional_day);
  const seconds = total_seconds % 60;
  total_seconds -= seconds;
  const hours = Math.floor(total_seconds / (60 * 60));
  const minutes = Math.floor(total_seconds / 60) % 60;
  // Buat tanggal dengan timezone lokal, bukan UTC
  return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
}

function getMonthName(dateObj) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
    return "N/A"; // Kembalikan 'N/A' jika objek Date tidak valid
  }
  return MONTH_ORDER[dateObj.getMonth()];
}

function parseNumber(value) {
  if (typeof value === "number") {
    return isNaN(value) ? 0 : value;
  }
  if (typeof value === "string") {
    const cleanedValue = value.trim();
    if (cleanedValue === "") return 0; // Handle empty string explicitly

    // Regex untuk mendeteksi format angka:
    // 1. Angka dengan koma sebagai desimal, titik sebagai ribuan (opsional)
    //    Contoh: "1.234,56", "123,45"
    const europeanRegex = /^-?\d{1,3}(\.\d{3})*(,\d+)?$/;
    // 2. Angka dengan titik sebagai desimal, koma sebagai ribuan (opsional)
    //    Contoh: "1,234.56", "123.45"
    const americanRegex = /^-?\d{1,3}(,\d{3})*(\.\d+)?$/;

    let numStr = cleanedValue;

    if (europeanRegex.test(numStr)) {
      // Format Eropa: hapus titik (pemisah ribuan), ganti koma (desimal) dengan titik
      numStr = numStr.replace(/\./g, "").replace(",", ".");
    } else if (americanRegex.test(numStr)) {
      // Format Amerika: hapus koma (pemisah ribuan)
      numStr = numStr.replace(/,/g, "");
    }
    // Jika tidak cocok regex di atas, coba parse langsung
    // Ini akan menangani angka tanpa pemisah ribuan, atau jika formatnya sudah sesuai float

    const num = parseFloat(numStr);
    return isNaN(num) ? 0 : num;
  }
  return 0; // Default jika bukan angka atau string yang bisa diparsing
}

// ... (averageGreaterThanZero tetap sama) ...
function averageGreaterThanZero(arr) {
  const filteredArr = arr.filter((num) => typeof num === "number" && num > 0);
  if (filteredArr.length === 0) {
    return 0;
  }
  return filteredArr.reduce((sum, val) => sum + val, 0) / filteredArr.length;
}

module.exports = {
  MONTH_ORDER,
  excelSerialNumberToDate,
  getMonthName,
  parseNumber, // Pastikan fungsi yang diperbarui diekspor
  averageGreaterThanZero,
};
