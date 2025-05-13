// index.js
const readlineSync = require("readline-sync");
const { MONTH_ORDER } = require("./src/utils");
const { readAndPreprocessData } = require("./src/excelReader");
const { performAggregation } = require("./src/aggregator");
const { prepareGroupBlock, writeOutputToFile } = require("./src/outputFormatter");

const DEFAULT_INPUT_FILENAME = "input.xlsx";
const DEFAULT_OUTPUT_FILENAME = "summary_output.xlsx";
const DEFAULT_SHEET_NAME = "DATA OLAH";

async function main() {
  console.log("Memulai proses pembuatan summary Excel...");

  const inputFile = readlineSync.question(`Masukkan nama file Excel input (default: ${DEFAULT_INPUT_FILENAME}): `) || DEFAULT_INPUT_FILENAME;
  const sheetName = readlineSync.question(`Masukkan nama sheet yang akan diproses (default: ${DEFAULT_SHEET_NAME}): `) || DEFAULT_SHEET_NAME;
  const periodYear = readlineSync.question(`Masukkan tahun periode (misal, 2024): `) || new Date().getFullYear();

  let allRawData = readAndPreprocessData(inputFile, sheetName);

  if (!allRawData || allRawData.length === 0) {
    console.log("Tidak ada data untuk diproses atau terjadi error saat membaca file.");
    return;
  }

  const workbookDataForExcelJS = [];
  const totalColumns = 5 + MONTH_ORDER.length * 2 + 3;

  const dataWithValidImporter = [];
  const dataWithBlankOrNAImporter = [];

  allRawData.forEach((row) => {
    if (row.importer === "" || row.importer === "N/A" || row.importer === null || typeof row.importer === "undefined") {
      dataWithBlankOrNAImporter.push(row);
    } else {
      dataWithValidImporter.push(row);
    }
  });

  function processSheetData(dataToProcessForSheet, sheetBaseName) {
    console.log(`\nMemproses data untuk sheet berbasis "${sheetBaseName}"...`);

    const groupedBySupplierOrOrigin = {};
    dataToProcessForSheet.forEach((row) => {
      const groupKey = row.supplier && row.supplier !== "N/A" && row.supplier.trim() !== "" ? row.supplier : row.originCountry;
      if (!groupedBySupplierOrOrigin[groupKey]) {
        groupedBySupplierOrOrigin[groupKey] = [];
      }
      groupedBySupplierOrOrigin[groupKey].push(row);
    });

    const allRowsForThisSheetContent = [];
    const supplierGroupsMeta = [];
    let sheetOverallMonthlyTotals = Array(12).fill(0);
    const itemSummaryDataForSheet = {};

    const groupKeys = Object.keys(groupedBySupplierOrOrigin).sort();
    groupKeys.forEach((groupName, groupIndex) => {
      console.log(`  - Memproses grup supplier/origin: ${groupName}`);
      const groupData = groupedBySupplierOrOrigin[groupName];
      const { summaryLvl1, summaryLvl2 } = performAggregation(groupData);

      if (summaryLvl2.length > 0) {
        const groupBlock = prepareGroupBlock(groupName, summaryLvl1, summaryLvl2);
        allRowsForThisSheetContent.push(...groupBlock.groupBlockRows);

        supplierGroupsMeta.push({
          name: groupName,
          productRowCount: groupBlock.distinctCombinationsCount,
          headerRowCount: groupBlock.headerRowCount,
          hasFollowingGroup: groupIndex < groupKeys.length - 1,
        });

        summaryLvl1.forEach((lvl1Row) => {
          const monthIndex = MONTH_ORDER.indexOf(lvl1Row.month);
          if (monthIndex !== -1) {
            sheetOverallMonthlyTotals[monthIndex] += lvl1Row.totalQty;
            const itemKey = `${lvl1Row.item}-${lvl1Row.gsm}-${lvl1Row.addOn}`;
            if (!itemSummaryDataForSheet[itemKey]) {
              itemSummaryDataForSheet[itemKey] = {
                item: lvl1Row.item,
                gsm: lvl1Row.gsm,
                addOn: lvl1Row.addOn,
                monthlyQtys: Array(12).fill(0),
                totalQtyRecap: 0,
              };
            }
            itemSummaryDataForSheet[itemKey].monthlyQtys[monthIndex] += lvl1Row.totalQty;
            itemSummaryDataForSheet[itemKey].totalQtyRecap += lvl1Row.totalQty;
          }
        });

        if (groupIndex < groupKeys.length - 1) {
          allRowsForThisSheetContent.push([]);
        }
      }
    });

    if (allRowsForThisSheetContent.length > 0) {
      allRowsForThisSheetContent.push([]);

      const totalAllHeaderMonthRow = ["Month", null, null, null, null];
      MONTH_ORDER.forEach((m) => totalAllHeaderMonthRow.push(m, null));
      totalAllHeaderMonthRow.push("RECAP", null, null);
      allRowsForThisSheetContent.push(totalAllHeaderMonthRow);

      const grandTotalAllSuppliers = sheetOverallMonthlyTotals.reduce((sum, qty) => sum + qty, 0);
      const totalAllMoRow = ["TOTAL ALL SUPPLIER PER MO", null, null, null, null];
      sheetOverallMonthlyTotals.forEach((total) => {
        totalAllMoRow.push(total, null);
      });
      totalAllMoRow.push(grandTotalAllSuppliers, null, null);
      allRowsForThisSheetContent.push(totalAllMoRow);

      const quarterlyTotalsAll = [0, 0, 0, 0];
      sheetOverallMonthlyTotals.forEach((total, i) => {
        if (i < 3) quarterlyTotalsAll[0] += total;
        else if (i < 6) quarterlyTotalsAll[1] += total;
        else if (i < 9) quarterlyTotalsAll[2] += total;
        else quarterlyTotalsAll[3] += total;
      });
      const totalAllQuartalRow = ["TOTAL ALL SUPPLIER PER QUARTAL", null, null, null, null];
      totalAllQuartalRow.push(quarterlyTotalsAll[0], null, null, null, null, null);
      totalAllQuartalRow.push(quarterlyTotalsAll[1], null, null, null, null, null);
      totalAllQuartalRow.push(quarterlyTotalsAll[2], null, null, null, null, null);
      totalAllQuartalRow.push(quarterlyTotalsAll[3], null, null, null, null, null);
      totalAllQuartalRow.push(null, null, null);
      allRowsForThisSheetContent.push(totalAllQuartalRow);

      allRowsForThisSheetContent.push([]);
      // Baris Judul Utama untuk Tabel "TOTAL PER ITEM"
      const itemTableMainTitleRow = ["TOTAL PER ITEM"]; // Hanya satu sel untuk di-merge full
      allRowsForThisSheetContent.push(itemTableMainTitleRow);

      // Baris Header Bulan untuk Tabel "TOTAL PER ITEM"
      const itemTableHeaderMonthRow = ["Month", null, null, null, null];
      MONTH_ORDER.forEach((m) => itemTableHeaderMonthRow.push(m, null));
      itemTableHeaderMonthRow.push("RECAP", null, null);
      allRowsForThisSheetContent.push(itemTableHeaderMonthRow);

      Object.keys(itemSummaryDataForSheet)
        .sort()
        .forEach((itemKey) => {
          const itemData = itemSummaryDataForSheet[itemKey];
          const itemRow = [`${itemData.item} ${itemData.gsm} ${itemData.addOn}`, null, null, null, null];
          itemData.monthlyQtys.forEach((qty) => itemRow.push(qty, null));
          itemRow.push(itemData.totalQtyRecap, null, null);
          allRowsForThisSheetContent.push(itemRow);
        });

      return {
        name: sheetBaseName,
        allRowsForSheetContent: allRowsForThisSheetContent,
        supplierGroupsMeta: supplierGroupsMeta,
        totalColumns: totalColumns,
      };
    }
    return null;
  }

  // ... (sisa kode main() untuk memproses importer kosong dan valid, sama seperti sebelumnya) ...
  if (dataWithBlankOrNAImporter.length > 0) {
    const blankImporterSheetNameInput = readlineSync.question("Masukkan nama sheet untuk data tanpa Importer (default: Data_Tanpa_Importer): ").trim() || "Data_Tanpa_Importer";
    const sheetName = blankImporterSheetNameInput.substring(0, 30).replace(/[\*\?\:\\\/\[\]]/g, "_");
    const sheetResult = processSheetData(dataWithBlankOrNAImporter, sheetName);
    if (sheetResult) {
      workbookDataForExcelJS.push(sheetResult);
      console.log(`    Data untuk sheet "${sheetName}" telah disiapkan.`);
    } else {
      console.log(`    Tidak ada data summary yang signifikan untuk data tanpa Importer untuk dimasukkan ke sheet "${sheetName}".`);
    }
  } else {
    console.log("Tidak ada data dengan Importer kosong atau N/A.");
  }

  const uniqueImporters = [...new Set(dataWithValidImporter.map((row) => row.importer))].sort();
  let existingSheetNames = workbookDataForExcelJS.map((s) => s.name);

  for (const importer of uniqueImporters) {
    const importerData = dataWithValidImporter.filter((row) => row.importer === importer);
    if (importerData.length === 0) continue;

    let baseSheetName = importer.replace(/[\*\?\:\\\/\[\]]/g, "_").substring(0, 30);
    const sheetResult = processSheetData(importerData, baseSheetName);

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
  await writeOutputToFile(workbookDataForExcelJS, outputFile, periodYear);
}

main().catch((err) => console.error("Terjadi kesalahan tidak terduga:", err));
