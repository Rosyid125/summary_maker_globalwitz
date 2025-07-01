// index.js
const readlineSync = require("readline-sync");
const path = require("path");
const { MONTH_ORDER } = require("./src/utils"); // Pastikan utils.js ada jika dibutuhkan
const { readAndPreprocessData, getExcelInfo, getSheetColumnNames, scanExcelFiles } = require("./src/excelReader");
const { performAggregation } = require("./src/aggregator");
const { prepareGroupBlock, writeOutputToFile } = require("./src/outputFormatter");

const DEFAULT_INPUT_FOLDER = "original_excel";
const DEFAULT_INPUT_FILENAME = "input.xlsx";
const DEFAULT_OUTPUT_FILENAME = "summary_output.xlsx";
const DEFAULT_SHEET_NAME = "DATA OLAH";

async function main() {
  console.log("Memulai proses pembuatan summary Excel...");

  // Scan files di folder input
  console.log("\n=== MEMINDAI FILE EXCEL ===");
  const availableFiles = scanExcelFiles(DEFAULT_INPUT_FOLDER);
  
  if (availableFiles.length === 0) {
    console.log(`Tidak ada file Excel (.xlsx/.xls) yang ditemukan di folder "${DEFAULT_INPUT_FOLDER}".`);
    console.log("Silakan letakkan file Excel di folder tersebut dan coba lagi.");
    return;
  }

  // Tampilkan daftar file yang tersedia
  console.log("File Excel yang tersedia:");
  availableFiles.forEach((file, index) => {
    console.log(`${index + 1}. ${file.name} (${file.size}, dimodifikasi: ${file.modified})`);
  });

  // Biarkan user memilih file
  const fileChoice = readlineSync.question(`\nPilih file Excel (1-${availableFiles.length}): `);
  
  if (!fileChoice || isNaN(fileChoice) || fileChoice < 1 || fileChoice > availableFiles.length) {
    console.log("Pilihan tidak valid. Proses dihentikan.");
    return;
  }

  const selectedFile = availableFiles[fileChoice - 1];
  console.log(`File yang dipilih: ${selectedFile.name}\n`);

  // Menggunakan path lengkap untuk file yang dipilih
  const inputFile = selectedFile.path;
  
  // Membaca informasi struktur Excel
  console.log("\nMembaca struktur file Excel...");
  const excelInfo = getExcelInfo(inputFile);
  if (!excelInfo) {
    console.log("Gagal membaca file Excel. Proses dihentikan.");
    return;
  }

  // Pemilihan Sheet
  console.log("\n=== PILIH SHEET ===");
  console.log("Sheet yang tersedia:");
  excelInfo.sheetNames.forEach((sheetName, index) => {
    console.log(`${index + 1}. ${sheetName}`);
  });
  
  const sheetChoice = readlineSync.question(`Pilih sheet (1-${excelInfo.sheetNames.length}, default: cari "${DEFAULT_SHEET_NAME}" atau pilih pertama): `);
  let selectedSheet;
  
  if (sheetChoice && !isNaN(sheetChoice) && sheetChoice >= 1 && sheetChoice <= excelInfo.sheetNames.length) {
    selectedSheet = excelInfo.sheetNames[sheetChoice - 1];
  } else {
    // Cari sheet default atau ambil yang pertama
    selectedSheet = excelInfo.sheetNames.find(sheet => sheet === DEFAULT_SHEET_NAME) || excelInfo.sheetNames[0];
  }
  
  console.log(`Sheet yang dipilih: "${selectedSheet}"\n`);

  // Membaca kolom dari sheet yang dipilih
  const columnNames = getSheetColumnNames(inputFile, selectedSheet);
  if (columnNames.length === 0) {
    console.log("Tidak ada kolom yang ditemukan di sheet tersebut. Proses dihentikan.");
    return;
  }

  // Pilihan format tanggal
  console.log("=== PILIH FORMAT TANGGAL ===");
  console.log("1. DD/MM/YYYY (Indonesia - default)");
  console.log("2. MM/DD/YYYY (USA/Global)");
  console.log("3. DD-MONTH-YYYY (dengan nama bulan, contoh: 01-mei-2025, 25-jan-2025)");
  console.log("\nCATATAN: Format Excel serial number (contoh: 45658) akan otomatis terdeteksi untuk semua pilihan di atas.");
  const dateFormatChoice = readlineSync.question("Masukkan pilihan (1, 2, atau 3, default: 1): ") || "1";
  
  let dateFormat;
  if (dateFormatChoice === "2") {
    dateFormat = "MM/DD/YYYY";
  } else if (dateFormatChoice === "3") {
    dateFormat = "DD-MONTH-YYYY";
  } else {
    dateFormat = "DD/MM/YYYY";
  }
  
  console.log(`Format tanggal yang dipilih: ${dateFormat}\n`);
  
  // Pilihan format angka/desimal
  console.log("=== PILIH FORMAT ANGKA ===");
  console.log("1. Koma sebagai desimal (1.234,56 - European/Indonesia - default)");
  console.log("2. Titik sebagai desimal (1,234.56 - American/Global)");
  const numberFormatChoice = readlineSync.question("Masukkan pilihan (1 atau 2, default: 1): ") || "1";
  const numberFormat = numberFormatChoice === "2" ? "AMERICAN" : "EUROPEAN";
  console.log(`Format angka yang dipilih: ${numberFormat === "EUROPEAN" ? "Koma sebagai desimal (1.234,56)" : "Titik sebagai desimal (1,234.56)"}\n`);
  
  // Helper function untuk pemilihan kolom
  function selectColumn(fieldName, defaultColumns = []) {
    console.log(`\n--- PILIH KOLOM UNTUK ${fieldName.toUpperCase()} ---`);
    console.log("Kolom yang tersedia:");
    columnNames.forEach((colName, index) => {
      const isDefault = defaultColumns.includes(colName);
      console.log(`${index + 1}. ${colName}${isDefault ? " (default)" : ""}`);
    });
    console.log(`${columnNames.length + 1}. Skip/Kosongkan`);
    
    const choice = readlineSync.question(`Pilih kolom untuk ${fieldName} (1-${columnNames.length + 1}${defaultColumns.length > 0 ? ', default: kolom default yang tersedia' : ''}): `);
    
    if (choice && !isNaN(choice)) {
      const choiceNum = parseInt(choice);
      if (choiceNum >= 1 && choiceNum <= columnNames.length) {
        return columnNames[choiceNum - 1];
      } else if (choiceNum === columnNames.length + 1) {
        return ""; // Skip
      }
    }
    
    // Default: cari kolom default yang tersedia
    for (const defaultCol of defaultColumns) {
      if (columnNames.includes(defaultCol)) {
        return defaultCol;
      }
    }
    
    return ""; // Tidak ada default yang cocok
  }

  // Mapping kolom Excel dengan pemilihan berbasis nomor
  console.log("\n=== MAPPING KOLOM EXCEL ===");
  console.log("Silakan pilih kolom yang sesuai untuk setiap field yang dibutuhkan:");
  
  const columnMapping = {
    date: selectColumn("TANGGAL", ["DATE", "CUSTOMS CLEARANCE DATE"]),
    hsCode: selectColumn("HS CODE", ["HS CODE"]),
    itemDesc: selectColumn("DESKRIPSI ITEM", ["ITEM DESC", "PRODUCT DESCRIPTION(EN)"]),
    gsm: selectColumn("GSM", ["GSM"]),
    item: selectColumn("ITEM", ["ITEM"]),
    addOn: selectColumn("ADD ON", ["ADD ON"]),
    importer: selectColumn("IMPORTER", ["IMPORTER", "PURCHASER"]),
    supplier: selectColumn("SUPPLIER", ["SUPPLIER"]),
    originCountry: selectColumn("ORIGIN COUNTRY", ["ORIGIN COUNTRY"]),
    unitPrice: selectColumn("UNIT PRICE USD", ["CIF KG Unit In USD", "USD Qty Unit", "UNIT PRICE(USD)"]),
    quantity: selectColumn("QUANTITY KG", ["Net KG Wt", "qty", "BUSINESS QUANTITY (KG)"])
  };
  
  console.log("\n=== MAPPING SELESAI ===");
  console.log("Mapping kolom yang dipilih:");
  Object.entries(columnMapping).forEach(([key, value]) => {
    console.log(`  ${key}: ${value || "(tidak dipilih/skip)"}`);
  });
  console.log();
  
  const periodYear = readlineSync.question(`Masukkan tahun periode (misal, 2024): `) || new Date().getFullYear();
  // --- TAMBAHAN: Input INCOTERM dari pengguna ---
  const incotermUserInput = readlineSync.question(`Masukkan nilai INCOTERM untuk kolom RECAP (misal, FOB, CIF, EXW, dll.): `).trim();
  const globalIncoterm = incotermUserInput || "N/A"; // Jika kosong, default ke N/A
  // --- AKHIR TAMBAHAN ---

  let allRawData = readAndPreprocessData(inputFile, selectedSheet, dateFormat, numberFormat, columnMapping);

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

  // --- MODIFIKASI: Tambahkan parameter incotermValue ---
  function processSheetData(dataToProcessForSheet, sheetBaseName, incotermValue) {
    console.log(`\nMemproses data untuk sheet berbasis "${sheetBaseName}" dengan INCOTERM: ${incotermValue}...`);

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
        // --- MODIFIKASI: Teruskan incotermValue ke prepareGroupBlock ---
        const groupBlock = prepareGroupBlock(groupName, summaryLvl1, summaryLvl2, incotermValue);
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
            const qtyToAdd = typeof lvl1Row.totalQty === "number" && !isNaN(lvl1Row.totalQty) ? lvl1Row.totalQty : 0;
            sheetOverallMonthlyTotals[monthIndex] += qtyToAdd;

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
            itemSummaryDataForSheet[itemKey].monthlyQtys[monthIndex] += qtyToAdd;
            itemSummaryDataForSheet[itemKey].totalQtyRecap += qtyToAdd;
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
        totalAllMoRow.push(total, null); // DIHAPUS: Math.round(total)
      });
      totalAllMoRow.push(grandTotalAllSuppliers, null, null); // DIHAPUS: Math.round(grandTotalAllSuppliers)
      allRowsForThisSheetContent.push(totalAllMoRow);

      const quarterlyTotalsAll = [0, 0, 0, 0];
      sheetOverallMonthlyTotals.forEach((total, i) => {
        if (i < 3) quarterlyTotalsAll[0] += total;
        else if (i < 6) quarterlyTotalsAll[1] += total;
        else if (i < 9) quarterlyTotalsAll[2] += total;
        else quarterlyTotalsAll[3] += total;
      });
      const totalAllQuartalRow = ["TOTAL ALL SUPPLIER PER QUARTAL", null, null, null, null];
      // DIHAPUS: Math.round() untuk setiap quarterlyTotalsAll
      totalAllQuartalRow.push(quarterlyTotalsAll[0], null, null, null, null, null);
      totalAllQuartalRow.push(quarterlyTotalsAll[1], null, null, null, null, null);
      totalAllQuartalRow.push(quarterlyTotalsAll[2], null, null, null, null, null);
      totalAllQuartalRow.push(quarterlyTotalsAll[3], null, null, null, null, null);
      totalAllQuartalRow.push(null, null, null);
      allRowsForThisSheetContent.push(totalAllQuartalRow);

      allRowsForThisSheetContent.push([]);
      const itemTableMainTitleRow = ["TOTAL PER ITEM"];
      allRowsForThisSheetContent.push(itemTableMainTitleRow);

      const itemTableHeaderMonthRow = ["Month", null, null, null, null];
      MONTH_ORDER.forEach((m) => itemTableHeaderMonthRow.push(m, null));
      itemTableHeaderMonthRow.push("RECAP", null, null);
      allRowsForThisSheetContent.push(itemTableHeaderMonthRow);

      Object.keys(itemSummaryDataForSheet)
        .sort()
        .forEach((itemKey) => {
          const itemData = itemSummaryDataForSheet[itemKey];
          const itemRow = [`${itemData.item} ${itemData.gsm} ${itemData.addOn}`, null, null, null, null];
          itemData.monthlyQtys.forEach((qty) => itemRow.push(qty, null)); // DIHAPUS: Math.round(qty)
          itemRow.push(itemData.totalQtyRecap, null, null); // DIHAPUS: Math.round(itemData.totalQtyRecap)
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
  // --- AKHIR MODIFIKASI ---

  if (dataWithBlankOrNAImporter.length > 0) {
    const blankImporterSheetNameInput = readlineSync.question("Masukkan nama sheet untuk data tanpa Importer (default: Data_Tanpa_Importer): ").trim() || "Data_Tanpa_Importer";
    const sheetNameForBlank = blankImporterSheetNameInput.substring(0, 30).replace(/[\*\?\:\\\/\[\]]/g, "_");
    // --- MODIFIKASI: Teruskan globalIncoterm ---
    const sheetResult = processSheetData(dataWithBlankOrNAImporter, sheetNameForBlank, globalIncoterm);
    if (sheetResult) {
      workbookDataForExcelJS.push(sheetResult);
      console.log(`    Data untuk sheet "${sheetNameForBlank}" telah disiapkan.`);
    } else {
      console.log(`    Tidak ada data summary yang signifikan untuk data tanpa Importer untuk dimasukkan ke sheet "${sheetNameForBlank}".`);
    }
  } else {
    console.log("Tidak ada data dengan Importer kosong atau -.");
  }

  const uniqueImporters = [...new Set(dataWithValidImporter.map((row) => row.importer))].sort();
  let existingSheetNames = workbookDataForExcelJS.map((s) => s.name);

  for (const importer of uniqueImporters) {
    const importerData = dataWithValidImporter.filter((row) => row.importer === importer);
    if (importerData.length === 0) continue;

    let baseSheetName = importer.replace(/[\*\?\:\\\/\[\]]/g, "_").substring(0, 30);
    // --- MODIFIKASI: Teruskan globalIncoterm ---
    const sheetResult = processSheetData(importerData, baseSheetName, globalIncoterm);

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
