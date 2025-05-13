// index.js
const readlineSync = require("readline-sync");
const { MONTH_ORDER } = require("./src/utils");
const { readAndPreprocessData } = require("./src/excelReader");
const { performAggregation } = require("./src/aggregator");
const { prepareGroupBlock, writeOutputToFile } = require("./src/outputFormatter"); // Ganti nama fungsi

const DEFAULT_INPUT_FILENAME = "input.xlsx";
const DEFAULT_OUTPUT_FILENAME = "summary_output.xlsx";
const DEFAULT_SHEET_NAME = "DATA OLAH";

async function main() {
  console.log("Memulai proses pembuatan summary Excel...");

  const inputFile = readlineSync.question(`Masukkan nama file Excel input (default: ${DEFAULT_INPUT_FILENAME}): `) || DEFAULT_INPUT_FILENAME;
  const sheetName = readlineSync.question(`Masukkan nama sheet yang akan diproses (default: ${DEFAULT_SHEET_NAME}): `) || DEFAULT_SHEET_NAME;
  const periodYear = readlineSync.question(`Masukkan tahun periode (misal, 2024): `) || new Date().getFullYear();

  let allData = readAndPreprocessData(inputFile, sheetName);

  if (!allData || allData.length === 0) {
    console.log("Tidak ada data untuk diproses atau terjadi error saat membaca file.");
    return;
  }

  const workbookDataForExcelJS = [];

  const dataWithValidImporter = [];
  const dataWithBlankOrNAImporter = [];

  allData.forEach((row) => {
    if (row.importer === "" || row.importer === "N/A" || row.importer === null || typeof row.importer === "undefined") {
      dataWithBlankOrNAImporter.push(row);
    } else {
      dataWithValidImporter.push(row);
    }
  });

  // Hitung total kolom sekali saja berdasarkan MONTH_ORDER dan kolom tetap
  const totalColumns = 5 + MONTH_ORDER.length * 2 + 3; // 5 awal + (bulan*2) + 3 recap

  function processAndBuildSheet(dataToProcess, sheetBaseName) {
    console.log(`\nMemproses data untuk sheet berbasis "${sheetBaseName}"...`);

    const groupedBySupplierOrOrigin = {};
    dataToProcess.forEach((row) => {
      const groupKey = row.supplier && row.supplier !== "N/A" ? row.supplier : row.originCountry;
      if (!groupedBySupplierOrOrigin[groupKey]) {
        groupedBySupplierOrOrigin[groupKey] = [];
      }
      groupedBySupplierOrOrigin[groupKey].push(row);
    });

    const allRowsForThisSheet = []; // Tidak ada header global di sini lagi
    const supplierGroupsMeta = [];

    const groupKeys = Object.keys(groupedBySupplierOrOrigin).sort();
    groupKeys.forEach((groupName, groupIndex) => {
      console.log(`  - Memproses grup: ${groupName}`);
      const groupData = groupedBySupplierOrOrigin[groupName];
      const { summaryLvl1, summaryLvl2 } = performAggregation(groupData);

      if (summaryLvl2.length > 0) {
        // prepareGroupBlock sekarang membuat blok lengkap (header + data + total)
        const groupBlock = prepareGroupBlock(groupName, summaryLvl1, summaryLvl2);
        allRowsForThisSheet.push(...groupBlock.groupBlockRows);

        supplierGroupsMeta.push({
          name: groupName,
          productRowCount: groupBlock.distinctCombinationsCount,
          headerRowCount: groupBlock.headerRowCount, // Simpan jumlah baris header grup
          hasFollowingGroup: groupIndex < groupKeys.length - 1,
        });

        if (groupIndex < groupKeys.length - 1) {
          allRowsForThisSheet.push([]);
        }
      }
    });

    if (allRowsForThisSheet.length > 0) {
      return {
        name: sheetBaseName,
        rowsForSheet: allRowsForThisSheet, // Berisi semua blok grup
        supplierGroupsMeta: supplierGroupsMeta,
        totalColumns: totalColumns, // Kirim total kolom untuk merge periode
      };
    }
    return null;
  }
  //--------------------------------------------------------------------

  if (dataWithBlankOrNAImporter.length > 0) {
    const blankImporterSheetNameInput = readlineSync.question("Masukkan nama sheet untuk data tanpa Importer: ").trim();
    if (blankImporterSheetNameInput) {
      const sheetName = blankImporterSheetNameInput.substring(0, 30).replace(/[\*\?\:\\\/\[\]]/g, "_");
      const sheetResult = processAndBuildSheet(dataWithBlankOrNAImporter, sheetName);
      if (sheetResult) {
        workbookDataForExcelJS.push(sheetResult);
        console.log(`    Data untuk sheet "${sheetName}" telah disiapkan.`);
      } else {
        console.log(`    Tidak ada data summary yang signifikan untuk data tanpa Importer.`);
      }
    } else {
      console.log("Nama sheet untuk data tanpa Importer tidak valid, data ini akan dilewati.");
    }
  }

  const uniqueImporters = [...new Set(dataWithValidImporter.map((row) => row.importer))].sort();
  let existingSheetNames = workbookDataForExcelJS.map((s) => s.name);

  for (const importer of uniqueImporters) {
    const importerData = dataWithValidImporter.filter((row) => row.importer === importer);
    if (importerData.length === 0) continue;

    let baseSheetName = importer.replace(/[\*\?\:\\\/\[\]]/g, "_").substring(0, 30);
    const sheetResult = processAndBuildSheet(importerData, baseSheetName);

    if (sheetResult) {
      let N = 0;
      let finalSheetName = sheetResult.name;
      while (existingSheetNames.includes(finalSheetName)) {
        N++;
        finalSheetName = `${sheetResult.name.substring(0, Math.max(0, 28 - String(N).length))}${N}`;
      }
      sheetResult.name = finalSheetName;
      existingSheetNames.push(finalSheetName);
      workbookDataForExcelJS.push(sheetResult);
      console.log(`    Data untuk sheet "${finalSheetName}" telah disiapkan.`);
    } else {
      console.log(`    Tidak ada data summary yang signifikan untuk IMPORTER: ${importer}.`);
    }
  }

  const outputFile = readlineSync.question(`Masukkan nama file Excel output (default: ${DEFAULT_OUTPUT_FILENAME}): `) || DEFAULT_OUTPUT_FILENAME;
  await writeOutputToFile(workbookDataForExcelJS, outputFile, periodYear); // Kirim periodYear
}

main().catch((err) => console.error("Terjadi kesalahan tidak terduga:", err));
