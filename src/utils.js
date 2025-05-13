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
    // Hilangkan semua spasi untuk menghindari kesalahan parsing
    const cleanedValue = value.trim();

    // Deteksi apakah ini format desimal Eropa (koma) atau Inggris (titik)
    const commaCount = (cleanedValue.match(/,/g) || []).length;
    const dotCount = (cleanedValue.match(/\./g) || []).length;

    // Jika ada koma dan titik, prioritaskan tanda desimal yang lebih umum (koma untuk Eropa)
    if (commaCount === 1 && dotCount === 0) {
      // Format Eropa, ganti koma dengan titik
      return parseFloat(cleanedValue.replace(",", "."));
    } else if (dotCount === 1 && commaCount === 0) {
      // Format Inggris, langsung parse
      return parseFloat(cleanedValue);
    } else if (commaCount > 1) {
      // Format Eropa dengan ribuan dan desimal
      const normalized = cleanedValue.replace(/\./g, "").replace(",", ".");
      return parseFloat(normalized);
    } else if (dotCount > 1) {
      // Format Inggris dengan ribuan
      const normalized = cleanedValue.replace(/,/g, "");
      return parseFloat(normalized);
    }
  }
  return 0;
}

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
  parseNumber,
  averageGreaterThanZero,
};
