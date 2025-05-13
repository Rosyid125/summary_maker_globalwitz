// src/outputFormatter.js
const { MONTH_ORDER } = require("./utils");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const DEFAULT_OUTPUT_FOLDER = "processed_excel";

function formatOutputForGroup(groupName, summaryLvl1Data, summaryLvl2Data) {
  const outputRowsAsStrings = []; // Akan menyimpan string yang sudah di-join dengan tab

  // Revisi Header
  const headerLine1Parts = ["SUPPLIER", "HS CODE", "ITEM", "GSM", "ADD ON"];
  const headerLine2Parts = ["", "", "", "", ""]; // Kolom awal kosong untuk baris kedua header

  MONTH_ORDER.forEach((month) => {
    headerLine1Parts.push(month, ""); // Nama bulan di baris 1, kolom sebelahnya kosong
    headerLine2Parts.push("PRICE", "QTY"); // PRICE dan QTY di baris 2
  });
  headerLine1Parts.push("RECAP", "", ""); // RECAP di baris 1, 2 kolom sebelahnya kosong
  headerLine2Parts.push("AVG PRICE", "INCOTERM", "TOTAL QTY"); // Detail rekap di baris 2

  outputRowsAsStrings.push(headerLine1Parts.join("\t"));
  outputRowsAsStrings.push(headerLine2Parts.join("\t"));

  const distinctHsCodeGsm = summaryLvl2Data
    .map((item) => ({ hsCode: item.hsCode, gsm: item.gsm }))
    .sort((a, b) => {
      if (a.hsCode < b.hsCode) return -1;
      if (a.hsCode > b.hsCode) return 1;
      if (a.gsm < b.gsm) return -1;
      if (a.gsm > b.gsm) return 1;
      return 0;
    })
    .filter((item, index, self) => index === self.findIndex((t) => t.hsCode === item.hsCode && t.gsm === item.gsm));

  distinctHsCodeGsm.forEach((combo, index) => {
    const rowParts = [];
    // Revisi Kolom Supplier: Hanya diisi untuk baris pertama dari setiap grup supplier
    if (index === 0) {
      rowParts.push(groupName);
    } else {
      rowParts.push(""); // Kosongkan untuk baris berikutnya dalam grup yang sama
    }
    rowParts.push(combo.hsCode);
    rowParts.push("N/A"); // ITEM - default
    rowParts.push(combo.gsm);
    rowParts.push("N/A"); // ADD ON - default

    MONTH_ORDER.forEach((month) => {
      const monthData = summaryLvl1Data.find((d) => d.hsCode === combo.hsCode && d.gsm === combo.gsm && d.month === month);
      if (monthData) {
        rowParts.push(monthData.avgPrice.toFixed(2));
        rowParts.push(Math.round(monthData.totalQty));
      } else {
        rowParts.push("N/A", "N/A");
      }
    });

    const recapData = summaryLvl2Data.find((d) => d.hsCode === combo.hsCode && d.gsm === combo.gsm);
    if (recapData) {
      rowParts.push(recapData.avgOfSummaryPrice.toFixed(2));
      rowParts.push("N/A"); // INCOTERM - default
      rowParts.push(Math.round(recapData.totalOfSummaryQty));
    } else {
      rowParts.push("N/A", "N/A", "N/A");
    }
    outputRowsAsStrings.push(rowParts.join("\t"));
  });
  return outputRowsAsStrings; // Kembalikan array of string (baris TSV)
}

function writeOutputToFile(workbookOutput, outputFileName = "summary_output.xlsx") {
  if (!fs.existsSync(DEFAULT_OUTPUT_FOLDER)) {
    fs.mkdirSync(DEFAULT_OUTPUT_FOLDER, { recursive: true });
  }
  const outputFile = path.join(DEFAULT_OUTPUT_FOLDER, outputFileName);

  if (workbookOutput.SheetNames.length > 0) {
    XLSX.writeFile(workbookOutput, outputFile);
    console.log(`\nProses selesai. Output disimpan di: ${outputFile}`);
  } else {
    console.log("\nTidak ada data yang diproses untuk output Excel.");
  }
}

module.exports = { formatOutputForGroup, writeOutputToFile };
