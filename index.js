// index.js
const readlineSync = require("readline-sync");
const XLSX = require("xlsx");
const { readAndPreprocessData } = require("./src/excelReader");
const { performAggregation } = require("./src/aggregator");
const { formatOutputForGroup, writeOutputToFile } = require("./src/outputFormatter");

const DEFAULT_INPUT_FILENAME = "input.xlsx";
const DEFAULT_OUTPUT_FILENAME = "summary_output.xlsx";
const DEFAULT_SHEET_NAME = "DATA OLAH";

async function main() {
  console.log("Memulai proses pembuatan summary Excel...");

  const inputFile = readlineSync.question(`Masukkan nama file Excel input (default: ${DEFAULT_INPUT_FILENAME}): `) || DEFAULT_INPUT_FILENAME;
  const sheetName = readlineSync.question(`Masukkan nama sheet yang akan diproses (default: ${DEFAULT_SHEET_NAME}): `) || DEFAULT_SHEET_NAME;

  let allData = readAndPreprocessData(inputFile, sheetName);

  if (!allData || allData.length === 0) {
    console.log("Tidak ada data untuk diproses atau terjadi error saat membaca file.");
    return;
  }

  const workbookOutput = XLSX.utils.book_new();

  // Pisahkan data: yang punya importer valid dan yang tidak
  const dataWithValidImporter = [];
  const dataWithBlankOrNAImporter = [];

  allData.forEach((row) => {
    if (row.importer === "" || row.importer === "N/A" || row.importer === null || typeof row.importer === "undefined") {
      dataWithBlankOrNAImporter.push(row);
    } else {
      dataWithValidImporter.push(row);
    }
  });

  // Proses data dengan IMPORTER kosong/N/A terlebih dahulu
  if (dataWithBlankOrNAImporter.length > 0) {
    console.log(`Ditemukan ${dataWithBlankOrNAImporter.length} baris dengan IMPORTER kosong atau "N/A".`);
    const blankImporterSheetName = readlineSync.question("Masukkan nama sheet untuk data tanpa Importer: ").trim();

    if (blankImporterSheetName) {
      console.log(`\nMemproses data tanpa Importer untuk sheet "${blankImporterSheetName}"...`);
      // Pengelompokan untuk data tanpa importer (berdasarkan Supplier/Origin)
      const groupedBySupplierOrOriginForBlank = {};
      dataWithBlankOrNAImporter.forEach((row) => {
        // Revisi 3: Jika SUPPLIER kosong atau N/A, gunakan ORIGIN COUNTRY
        const groupKey = row.supplier && row.supplier !== "N/A" ? row.supplier : row.originCountry;
        if (!groupedBySupplierOrOriginForBlank[groupKey]) {
          groupedBySupplierOrOriginForBlank[groupKey] = [];
        }
        groupedBySupplierOrOriginForBlank[groupKey].push(row);
      });

      const sheetDataForBlankImporterTSVRows = [];
      const groupKeysBlank = Object.keys(groupedBySupplierOrOriginForBlank).sort();

      groupKeysBlank.forEach((groupName, groupIndex) => {
        console.log(`  - Memproses grup (tanpa importer): ${groupName}`);
        const groupData = groupedBySupplierOrOriginForBlank[groupName];
        const { summaryLvl1, summaryLvl2 } = performAggregation(groupData);

        if (summaryLvl2.length > 0) {
          const formattedGroupOutputTSVRows = formatOutputForGroup(groupName, summaryLvl1, summaryLvl2);
          sheetDataForBlankImporterTSVRows.push(...formattedGroupOutputTSVRows);
          if (groupIndex < groupKeysBlank.length - 1) {
            sheetDataForBlankImporterTSVRows.push("");
          }
        }
      });

      if (sheetDataForBlankImporterTSVRows.length > 0) {
        const sheetAoA = sheetDataForBlankImporterTSVRows.map((rowStr) => rowStr.split("\t"));
        const newSheet = XLSX.utils.aoa_to_sheet(sheetAoA);
        XLSX.utils.book_append_sheet(workbookOutput, newSheet, blankImporterSheetName.substring(0, 30)); // Batasi panjang
        console.log(`    Sheet "${blankImporterSheetName.substring(0, 30)}" untuk data tanpa Importer telah dibuat.`);
      } else {
        console.log(`    Tidak ada data summary yang dihasilkan untuk data tanpa Importer.`);
      }
    } else {
      console.log("Nama sheet untuk data tanpa Importer tidak valid, data ini akan dilewati.");
    }
  }

  // Proses data dengan IMPORTER valid
  const uniqueImporters = [...new Set(dataWithValidImporter.map((row) => row.importer))].sort();

  for (const importer of uniqueImporters) {
    console.log(`\nMemproses untuk IMPORTER: ${importer}...`);
    const importerData = dataWithValidImporter.filter((row) => row.importer === importer);
    if (importerData.length === 0) continue;

    const groupedBySupplierOrOrigin = {};
    importerData.forEach((row) => {
      // Revisi 3: Jika SUPPLIER kosong atau N/A, gunakan ORIGIN COUNTRY
      const groupKey = row.supplier && row.supplier !== "N/A" ? row.supplier : row.originCountry;
      if (!groupedBySupplierOrOrigin[groupKey]) {
        groupedBySupplierOrOrigin[groupKey] = [];
      }
      groupedBySupplierOrOrigin[groupKey].push(row);
    });

    const sheetDataForImporterTSVRows = [];
    const groupKeys = Object.keys(groupedBySupplierOrOrigin).sort();

    groupKeys.forEach((groupName, groupIndex) => {
      console.log(`  - Memproses grup: ${groupName}`);
      const groupData = groupedBySupplierOrOrigin[groupName];
      const { summaryLvl1, summaryLvl2 } = performAggregation(groupData);

      if (summaryLvl2.length > 0) {
        const formattedGroupOutputTSVRows = formatOutputForGroup(groupName, summaryLvl1, summaryLvl2);
        sheetDataForImporterTSVRows.push(...formattedGroupOutputTSVRows);
        if (groupIndex < groupKeys.length - 1) {
          sheetDataForImporterTSVRows.push("");
        }
      }
    });

    if (sheetDataForImporterTSVRows.length > 0) {
      const sheetAoA = sheetDataForImporterTSVRows.map((rowStr) => rowStr.split("\t"));
      const newSheet = XLSX.utils.aoa_to_sheet(sheetAoA);

      let currentSheetName = importer.replace(/[\*\?\:\\\/\[\]]/g, "_");
      currentSheetName = currentSheetName.substring(0, 30);

      let N = 0;
      let finalSheetName = currentSheetName;
      while (workbookOutput.SheetNames.includes(finalSheetName)) {
        N++;
        finalSheetName = `${currentSheetName.substring(0, Math.max(0, 28 - String(N).length))}${N}`; // Pastikan cukup ruang untuk nomor
      }

      XLSX.utils.book_append_sheet(workbookOutput, newSheet, finalSheetName);
      console.log(`    Sheet "${finalSheetName}" telah dibuat.`);
    } else {
      console.log(`    Tidak ada data summary yang dihasilkan untuk IMPORTER: ${importer}.`);
    }
  }

  const outputFile = readlineSync.question(`Masukkan nama file Excel output (default: ${DEFAULT_OUTPUT_FILENAME}): `) || DEFAULT_OUTPUT_FILENAME;
  writeOutputToFile(workbookOutput, outputFile);
}

main().catch((err) => console.error("Terjadi kesalahan tidak terduga:", err));
