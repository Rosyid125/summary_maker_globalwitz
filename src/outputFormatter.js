// src/outputFormatter.js
const { MONTH_ORDER } = require("./utils");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const DEFAULT_OUTPUT_FOLDER = "processed_excel";

/**
 * Menyiapkan SELURUH BLOK DATA (HEADER + produk + total) untuk SATU GRUP supplier/origin.
 * @param {string} groupName Nama grup (Supplier/Origin).
 * @param {Array<Object>} summaryLvl1Data Data summary bulanan untuk grup ini.
 * @param {Array<Object>} summaryLvl2Data Data rekap untuk grup ini.
 * @returns {Object} Objek berisi {
 *    groupBlockRows: Array<Array<any>>, // Termasuk HEADER, data produk & baris total untuk grup ini
 *    overallTotalQtyForGroup: number,
 *    distinctCombinationsCount: number, // Jumlah baris data produk
 *    headerRowCount: 2 // Jumlah baris untuk header grup ini
 * }
 */
function prepareGroupBlock(groupName, summaryLvl1Data, summaryLvl2Data) {
  const groupBlockRows = [];
  const headerRowCount = 2;

  // --- Tambahkan Header untuk grup ini ---
  const headerRow1 = ["SUPPLIER", "HS CODE", "ITEM", "GSM", "ADD ON"];
  const headerRow2 = [null, null, null, null, null];
  MONTH_ORDER.forEach((month) => {
    headerRow1.push(month, null);
    headerRow2.push("PRICE", "QTY");
  });
  headerRow1.push("RECAP", null, null);
  headerRow2.push("AVG PRICE", "INCOTERM", "TOTAL QTY");
  groupBlockRows.push(headerRow1);
  groupBlockRows.push(headerRow2);

  // --- Data Produk ---
  const monthlyTotals = Array(12).fill(0);
  const distinctCombinations = summaryLvl2Data
    .map((item) => ({
      /* ... seperti sebelumnya ... */ hsCode: item.hsCode,
      item: item.item,
      gsm: item.gsm,
      addOn: item.addOn,
    }))
    .sort((a, b) => {
      /* ... logika sorting sama ... */
      if (a.hsCode < b.hsCode) return -1;
      if (a.hsCode > b.hsCode) return 1;
      if (a.item < b.item) return -1;
      if (a.item > b.item) return 1;
      if (a.gsm < b.gsm) return -1;
      if (a.gsm > b.gsm) return 1;
      if (a.addOn < b.addOn) return -1;
      if (a.addOn > b.addOn) return 1;
      return 0;
    })
    .filter((item, index, self) => index === self.findIndex((t) => t.hsCode === item.hsCode && t.item === item.item && t.gsm === item.gsm && t.addOn === item.addOn));

  distinctCombinations.forEach((combo, index) => {
    const dataRow = [];
    dataRow.push(index === 0 ? groupName : null);
    dataRow.push(combo.hsCode);
    dataRow.push(combo.item);
    dataRow.push(combo.gsm);
    dataRow.push(combo.addOn);
    MONTH_ORDER.forEach((month, monthIndex) => {
      /* ... seperti sebelumnya ... */
      const monthData = summaryLvl1Data.find((d) => d.hsCode === combo.hsCode && d.item === combo.item && d.gsm === combo.gsm && d.addOn === combo.addOn && d.month === month);
      if (monthData) {
        dataRow.push(parseFloat(monthData.avgPrice.toFixed(2)));
        const qty = Math.round(monthData.totalQty);
        dataRow.push(qty);
        monthlyTotals[monthIndex] += qty;
      } else {
        dataRow.push("N/A", "N/A");
      }
    });
    const recapData = summaryLvl2Data.find((d) => d.hsCode === combo.hsCode && d.item === combo.item && d.gsm === combo.gsm && d.addOn === combo.addOn);
    if (recapData) {
      /* ... seperti sebelumnya ... */
      dataRow.push(parseFloat(recapData.avgOfSummaryPrice.toFixed(2)));
      dataRow.push("N/A");
      dataRow.push(Math.round(recapData.totalOfSummaryQty));
    } else {
      dataRow.push("N/A", "N/A", "N/A");
    }
    groupBlockRows.push(dataRow);
  });

  const overallTotalQtyForThisGroup = monthlyTotals.reduce((sum, qty) => sum + qty, 0);

  if (distinctCombinations.length > 0) {
    const totalQtyPerMoRow = ["TOTAL QTY PER MO", null, null, null, null];
    monthlyTotals.forEach((total) => {
      totalQtyPerMoRow.push(total, null);
    });
    totalQtyPerMoRow.push(null, null, overallTotalQtyForThisGroup);
    groupBlockRows.push(totalQtyPerMoRow);

    const quarterlyTotals = [0, 0, 0, 0];
    monthlyTotals.forEach((total, index) => {
      /* ... seperti sebelumnya ... */
      if (index < 3) quarterlyTotals[0] += total;
      else if (index < 6) quarterlyTotals[1] += total;
      else if (index < 9) quarterlyTotals[2] += total;
      else quarterlyTotals[3] += total;
    });
    const totalQtyPerQuartalRow = ["TOTAL QTY PER QUARTAL", null, null, null, null];
    totalQtyPerQuartalRow.push(quarterlyTotals[0], null, null, null, null, null);
    totalQtyPerQuartalRow.push(quarterlyTotals[1], null, null, null, null, null);
    totalQtyPerQuartalRow.push(quarterlyTotals[2], null, null, null, null, null);
    totalQtyPerQuartalRow.push(quarterlyTotals[3], null, null, null, null, null);
    totalQtyPerQuartalRow.push(null, null, null);
    groupBlockRows.push(totalQtyPerQuartalRow);
  }

  return {
    groupBlockRows: groupBlockRows,
    overallTotalQtyForGroup: overallTotalQtyForThisGroup,
    distinctCombinationsCount: distinctCombinations.length,
    headerRowCount: headerRowCount,
  };
}

async function writeOutputToFile(workbookData, outputFileName = "summary_output.xlsx", periodYear) {
  if (!fs.existsSync(DEFAULT_OUTPUT_FOLDER)) {
    fs.mkdirSync(DEFAULT_OUTPUT_FOLDER, { recursive: true });
  }
  const outputFile = path.join(DEFAULT_OUTPUT_FOLDER, outputFileName);
  const workbook = new ExcelJS.Workbook();

  for (const sheetInfo of workbookData) {
    // sheetInfo: { name, rowsForSheet, supplierGroupsMeta, totalColumns }
    const worksheet = workbook.addWorksheet(sheetInfo.name);

    // Tambahkan Baris Judul Periode
    const periodTitleRow = worksheet.addRow([`${periodYear} PERIODE`]);
    worksheet.mergeCells(1, 1, 1, sheetInfo.totalColumns); // Merge semua kolom di baris 1
    periodTitleRow.getCell(1).alignment = { vertical: "middle", horizontal: "center" };
    periodTitleRow.getCell(1).font = { bold: true, size: 14 };
    periodTitleRow.height = 20;

    // Tambahkan sisa baris (yang sudah berisi header per grup)
    sheetInfo.rowsForSheet.forEach((rowData, rowIndex) => {
      const rowWithNAOrEmpty = rowData.map((cell) => (cell === null || typeof cell === "undefined" ? "" : cell));
      const addedRowCurrent = worksheet.addRow(rowWithNAOrEmpty); // Baris data dimulai dari baris ke-2 di worksheet

      // Set nilai header secara eksplisit jika itu baris header pertama dari blok
      if (rowData[0] === "SUPPLIER" && rowData[1] === "HS CODE") {
        // Cek jika ini baris header pertama
        let colIdx = 1;
        rowData.forEach((headerText) => {
          if (headerText !== null && typeof headerText !== "undefined") {
            addedRowCurrent.getCell(colIdx).value = headerText;
          }
          colIdx++;
        });
      }
    });

    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      if (rowNumber === 1) return; // Lewati baris judul periode untuk alignment default
      row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        cell.alignment = { vertical: "middle", horizontal: "center" };
      });
    });

    let currentSheetRow = 2; // Mulai setelah baris judul periode
    for (const groupMeta of sheetInfo.supplierGroupsMeta) {
      const headerStartRow = currentSheetRow;
      const dataStartRow = headerStartRow + groupMeta.headerRowCount;
      const productRows = groupMeta.productRowCount;

      // Merge Header Grup
      worksheet.mergeCells(headerStartRow, 1, headerStartRow + 1, 1); // Supplier
      worksheet.mergeCells(headerStartRow, 2, headerStartRow + 1, 2); // HS CODE
      worksheet.mergeCells(headerStartRow, 3, headerStartRow + 1, 3); // ITEM
      worksheet.mergeCells(headerStartRow, 4, headerStartRow + 1, 4); // GSM
      worksheet.mergeCells(headerStartRow, 5, headerStartRow + 1, 5); // ADD ON

      let startHeaderMonthCol = 6;
      for (let i = 0; i < MONTH_ORDER.length; i++) {
        worksheet.mergeCells(headerStartRow, startHeaderMonthCol, headerStartRow, startHeaderMonthCol + 1);
        startHeaderMonthCol += 2;
      }
      worksheet.mergeCells(headerStartRow, startHeaderMonthCol, headerStartRow, startHeaderMonthCol + 2);

      // Merge Kolom Supplier Vertikal untuk data produk
      if (productRows > 0) {
        worksheet.mergeCells(dataStartRow, 1, dataStartRow + productRows - 1, 1);
      }

      if (productRows > 0) {
        const totalQtyPerMoRowIndexForGroup = dataStartRow + productRows;
        const quartalRowIndexForGroup = totalQtyPerMoRowIndexForGroup + 1;

        worksheet.mergeCells(totalQtyPerMoRowIndexForGroup, 1, totalQtyPerMoRowIndexForGroup, 5);
        worksheet.getCell(totalQtyPerMoRowIndexForGroup, 1).font = { bold: true };
        let currentMonthTotalCol = 6;
        for (let i = 0; i < MONTH_ORDER.length; i++) {
          worksheet.mergeCells(totalQtyPerMoRowIndexForGroup, currentMonthTotalCol, totalQtyPerMoRowIndexForGroup, currentMonthTotalCol + 1);
          currentMonthTotalCol += 2;
        }

        worksheet.mergeCells(quartalRowIndexForGroup, 1, quartalRowIndexForGroup, 5);
        worksheet.getCell(quartalRowIndexForGroup, 1).font = { bold: true };
        let currentQuartalCol = 6;
        for (let q = 0; q < 4; q++) {
          worksheet.mergeCells(quartalRowIndexForGroup, currentQuartalCol, quartalRowIndexForGroup, currentQuartalCol + 5);
          currentQuartalCol += 6;
        }

        const recapStartColIndex = 5 + MONTH_ORDER.length * 2 + 1;
        worksheet.mergeCells(totalQtyPerMoRowIndexForGroup, recapStartColIndex, quartalRowIndexForGroup, recapStartColIndex + 2);
        worksheet.getCell(totalQtyPerMoRowIndexForGroup, recapStartColIndex).font = { bold: true };
      }

      // Styling header grup
      [worksheet.getRow(headerStartRow), worksheet.getRow(headerStartRow + 1)].forEach((row) => {
        row.font = { bold: true };
      });

      currentSheetRow += groupMeta.headerRowCount + productRows + (productRows > 0 ? 2 : 0) + (groupMeta.hasFollowingGroup ? 1 : 0);
    }
  }

  if (workbook.worksheets.length > 0) {
    await workbook.xlsx.writeFile(outputFile);
    console.log(`\nProses selesai. Output disimpan di: ${outputFile}`);
  } else {
    console.log("\nTidak ada data yang diproses untuk output Excel.");
  }
}

module.exports = { prepareGroupBlock, writeOutputToFile }; // Ganti nama fungsi
