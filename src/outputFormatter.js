// src/outputFormatter.js
const { MONTH_ORDER } = require("./utils");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const DEFAULT_OUTPUT_FOLDER = "processed_excel";

function formatOutputForGroup(groupName, summaryLvl1Data, summaryLvl2Data) {
  const outputRowsAsStrings = [];

  const headerLine1Parts = ["SUPPLIER", "HS CODE", "ITEM", "GSM", "ADD ON"];
  const headerLine2Parts = ["", "", "", "", ""];

  MONTH_ORDER.forEach((month) => {
    headerLine1Parts.push(month, "");
    headerLine2Parts.push("PRICE", "QTY");
  });
  headerLine1Parts.push("RECAP", "", "");
  headerLine2Parts.push("AVG PRICE", "INCOTERM", "TOTAL QTY");

  outputRowsAsStrings.push(headerLine1Parts.join("\t"));
  outputRowsAsStrings.push(headerLine2Parts.join("\t"));

  // REVISI: Sorting sekarang juga berdasarkan ITEM dan ADD ON
  const distinctCombinations = summaryLvl2Data
    .map((item) => ({
      hsCode: item.hsCode,
      item: item.item, // Tambahkan item
      gsm: item.gsm,
      addOn: item.addOn, // Tambahkan addOn
    }))
    .sort((a, b) => {
      if (a.hsCode < b.hsCode) return -1;
      if (a.hsCode > b.hsCode) return 1;
      if (a.item < b.item) return -1; // Urutkan berdasarkan item
      if (a.item > b.item) return 1;
      if (a.gsm < b.gsm) return -1;
      if (a.gsm > b.gsm) return 1;
      if (a.addOn < b.addOn) return -1; // Urutkan berdasarkan addOn
      if (a.addOn > b.addOn) return 1;
      return 0;
    })
    .filter(
      (item, index, self) =>
        index ===
        self.findIndex(
          (t) =>
            t.hsCode === item.hsCode &&
            t.item === item.item && // Bandingkan item
            t.gsm === item.gsm &&
            t.addOn === item.addOn // Bandingkan addOn
        )
    );

  distinctCombinations.forEach((combo, index) => {
    const rowParts = [];
    rowParts.push(index === 0 ? groupName : "");
    rowParts.push(combo.hsCode);
    // REVISI: Isi kolom ITEM dan ADD ON dari data combo
    rowParts.push(combo.item);
    rowParts.push(combo.gsm);
    rowParts.push(combo.addOn);

    MONTH_ORDER.forEach((month) => {
      // REVISI: Pencarian data bulanan sekarang juga berdasarkan ITEM dan ADD ON
      const monthData = summaryLvl1Data.find((d) => d.hsCode === combo.hsCode && d.item === combo.item && d.gsm === combo.gsm && d.addOn === combo.addOn && d.month === month);
      if (monthData) {
        rowParts.push(monthData.avgPrice.toFixed(2));
        rowParts.push(Math.round(monthData.totalQty));
      } else {
        rowParts.push("N/A", "N/A");
      }
    });

    // REVISI: Pencarian data rekap sekarang juga berdasarkan ITEM dan ADD ON
    const recapData = summaryLvl2Data.find((d) => d.hsCode === combo.hsCode && d.item === combo.item && d.gsm === combo.gsm && d.addOn === combo.addOn);
    if (recapData) {
      rowParts.push(recapData.avgOfSummaryPrice.toFixed(2));
      rowParts.push("N/A");
      rowParts.push(Math.round(recapData.totalOfSummaryQty));
    } else {
      rowParts.push("N/A", "N/A", "N/A");
    }
    outputRowsAsStrings.push(rowParts.join("\t"));
  });
  return outputRowsAsStrings;
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
