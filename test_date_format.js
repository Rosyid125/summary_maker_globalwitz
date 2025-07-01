// Test file untuk verifikasi format tanggal baru
const { readAndPreprocessData } = require('./src/excelReader');

// Untuk test kita perlu import langsung dari excelReader
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");
const { MONTH_ORDER, getMonthName, parseNumber } = require("./src/utils");

// Copy fungsi parseDate_DDMONTHYYYY untuk test
function parseDate_DDMONTHYYYY(dateString) {
  if (typeof dateString !== "string") return null;

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

console.log('Testing DD-MONTH-YYYY format:');
console.log('01-mei-2025 =>', parseDate_DDMONTHYYYY('01-mei-2025')); // Expected: Date object for May 1, 2025
console.log('25-jan-2025 =>', parseDate_DDMONTHYYYY('25-jan-2025')); // Expected: Date object for Jan 25, 2025
console.log('15/feb/2024 =>', parseDate_DDMONTHYYYY('15/feb/2024')); // Expected: Date object for Feb 15, 2024
console.log('30.des.2023 =>', parseDate_DDMONTHYYYY('30.des.2023')); // Expected: Date object for Dec 30, 2023
console.log('12-agustus-2025 =>', parseDate_DDMONTHYYYY('12-agustus-2025')); // Expected: Date object for Aug 12, 2025
console.log('05-okt-2024 =>', parseDate_DDMONTHYYYY('05-okt-2024')); // Expected: Date object for Oct 5, 2024

console.log('\nTesting edge cases:');
console.log('Invalid month =>', parseDate_DDMONTHYYYY('01-xyz-2025')); // Expected: null
console.log('Invalid format =>', parseDate_DDMONTHYYYY('01/02/2025')); // Expected: null (no month name)
console.log('Empty string =>', parseDate_DDMONTHYYYY('')); // Expected: null

console.log('\nTesting getMonthName with parsed dates:');
const testDate1 = parseDate_DDMONTHYYYY('01-mei-2025');
const testDate2 = parseDate_DDMONTHYYYY('25-des-2024');
if (testDate1) console.log('01-mei-2025 -> Month name:', getMonthName(testDate1)); // Expected: "Mei"
if (testDate2) console.log('25-des-2024 -> Month name:', getMonthName(testDate2)); // Expected: "Des"
