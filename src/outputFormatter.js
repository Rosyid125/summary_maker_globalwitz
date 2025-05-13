// src/outputFormatter.js
const { MONTH_ORDER } = require("./utils");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const DEFAULT_OUTPUT_FOLDER = "processed_excel";

// prepareGroupBlock tetap sama seperti versi sebelumnya

function prepareGroupBlock(groupName, summaryLvl1Data, summaryLvl2Data) {
  const groupBlockRows = [];
  const headerRowCount = 2;

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

  const monthlyTotals = Array(12).fill(0);
  const distinctCombinations = summaryLvl2Data
    .map((item) => ({
      hsCode: item.hsCode,
      item: item.item,
      gsm: item.gsm,
      addOn: item.addOn,
    }))
    .sort((a, b) => {
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
    header1Length: headerRow1.length, // Diperlukan untuk styling kolom recap
  };
}

// src/outputFormatter.js
// ... (prepareGroupBlock tetap sama)

async function writeOutputToFile(workbookData, outputFileName = "summary_output.xlsx", periodYear) {
  if (!fs.existsSync(DEFAULT_OUTPUT_FOLDER)) {
    fs.mkdirSync(DEFAULT_OUTPUT_FOLDER, { recursive: true });
  }
  const outputFile = path.join(DEFAULT_OUTPUT_FOLDER, outputFileName);
  const workbook = new ExcelJS.Workbook();

  for (const sheetInfo of workbookData) {
    // sheetInfo: { name, allRowsForSheetContent, supplierGroupsMeta, totalColumns }
    const worksheet = workbook.addWorksheet(sheetInfo.name);

    const periodTitleRowCell = worksheet.addRow([`${periodYear} PERIODE`]).getCell(1);
    worksheet.mergeCells(1, 1, 1, sheetInfo.totalColumns);
    periodTitleRowCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF7030A0" } }; // Warna dari colors object
    periodTitleRowCell.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
    periodTitleRowCell.alignment = { vertical: "middle", horizontal: "center" };
    worksheet.getRow(1).height = 20;

    // Tambahkan konten utama sheet (blok supplier, total sheet, total item)
    sheetInfo.allRowsForSheetContent.forEach((rowData, rowIndexInContent) => {
      const actualSheetRowIndex = rowIndexInContent + 2; // +1 karena baris periode, +1 karena baris berikutnya
      const rowWithNAOrEmpty = rowData.map((cell) => (cell === null || typeof cell === "undefined" ? "" : cell));
      const addedRow = worksheet.addRow(rowWithNAOrEmpty);

      // Set nilai header per grup jika itu baris header pertama dari blok
      if (rowData.length > 1 && rowData[0] === "SUPPLIER" && rowData[1] === "HS CODE") {
        let colIdx = 1;
        rowData.forEach((headerText) => {
          if (headerText !== null && typeof headerText !== "undefined") {
            addedRow.getCell(colIdx).value = headerText;
          }
          colIdx++;
        });
      }
    });

    // --- Styling dan Merge ---
    // Default Alignment dan Border
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      if (rowNumber === 1) return;
      row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
    });

    // Iterasi melalui grup supplier untuk merge spesifik grup
    let currentSheetRowForBlocks = 2; // Mulai setelah baris judul periode
    for (const groupMeta of sheetInfo.supplierGroupsMeta) {
      const headerStartRow = currentSheetRowForBlocks;
      const dataStartRow = headerStartRow + groupMeta.headerRowCount;
      const productRows = groupMeta.productRowCount;

      // Styling & Merge Header Grup (seperti sebelumnya)
      // ... (kode styling dan merge header grup dari versi sebelumnya) ...
      const header1 = worksheet.getRow(headerStartRow);
      const header2 = worksheet.getRow(headerStartRow + 1);
      [header1, header2].forEach((r) => (r.font = { bold: true, color: { argb: "FFFFFFFF" } }));
      for (let c = 1; c <= 5; c++) {
        /* ... supplierCols ... */
        header1.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF002060" } };
        header2.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF002060" } };
      }
      worksheet.mergeCells(headerStartRow, 1, headerStartRow + 1, 1); /* ... merge A-E ... */
      worksheet.mergeCells(headerStartRow, 2, headerStartRow + 1, 2);
      worksheet.mergeCells(headerStartRow, 3, headerStartRow + 1, 3);
      worksheet.mergeCells(headerStartRow, 4, headerStartRow + 1, 4);
      worksheet.mergeCells(headerStartRow, 5, headerStartRow + 1, 5);

      let currentColForColor = 6;
      const colors = { q1: "FFFFC000", q2: "FF00B050", q3: "FFFFFF00", q4: "FF00B0F0", recap: "FF002060" };
      // Q1-Q4 & RECAP coloring
      for (let q = 0; q < 4; q++) {
        const qColor = q === 0 ? colors.q1 : q === 1 ? colors.q2 : q === 2 ? colors.q3 : colors.q4;
        for (let i = 0; i < 3; i++) {
          worksheet.mergeCells(headerStartRow, currentColForColor, headerStartRow, currentColForColor + 1);
          header1.getCell(currentColForColor).fill = { type: "pattern", pattern: "solid", fgColor: { argb: qColor } };
          header1.getCell(currentColForColor + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: qColor } };
          header2.getCell(currentColForColor).fill = { type: "pattern", pattern: "solid", fgColor: { argb: qColor } };
          header2.getCell(currentColForColor + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: qColor } };
          if (qColor === colors.q3) {
            // Kuning cerah, teks hitam
            header1.getCell(currentColForColor).font = { bold: true, color: { argb: "FF000000" } };
            header2.getCell(currentColForColor).font = { bold: true, color: { argb: "FF000000" } };
            header2.getCell(currentColForColor + 1).font = { bold: true, color: { argb: "FF000000" } };
          }
          currentColForColor += 2;
        }
      }
      worksheet.mergeCells(headerStartRow, currentColForColor, headerStartRow, currentColForColor + 2);
      for (let i = 0; i < 3; i++) {
        header1.getCell(currentColForColor + i).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.recap } };
        header2.getCell(currentColForColor + i).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.recap } };
      }

      if (productRows > 0) {
        worksheet.mergeCells(dataStartRow, 1, dataStartRow + productRows - 1, 1);

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
      currentSheetRowForBlocks += groupMeta.headerRowCount + productRows + (productRows > 0 ? 2 : 0) + (groupMeta.hasFollowingGroup ? 1 : 0);
    }

    // --- Merge untuk Tabel TOTAL ALL SUPPLIER & TOTAL PER ITEM di akhir sheet ---
    // Kita perlu melacak posisi baris ini dengan lebih baik
    let lastRowProcessed = currentSheetRowForBlocks - 1; // Baris terakhir dari blok supplier terakhir
    if (sheetInfo.supplierGroupsMeta.length > 0 && !sheetInfo.supplierGroupsMeta[sheetInfo.supplierGroupsMeta.length - 1].hasFollowingGroup) {
      // Jika grup terakhir tidak ada pemisah, kurangi 1
      // Tidak perlu jika baris kosong sudah dihitung benar
    }

    // Cari baris "TOTAL ALL SUPPLIER PER MO"
    let totalAllMoRowActualIndex = -1;
    let totalAllQuartalRowActualIndex = -1;
    let itemTableHeader1ActualIndex = -1;

    for (let i = 2; i <= worksheet.rowCount; i++) {
      // Mulai dari baris 3 (setelah periode dan header global)
      if (worksheet.getRow(i).getCell(1).value === "TOTAL ALL SUPPLIER PER MO") {
        totalAllMoRowActualIndex = i;
        totalAllQuartalRowActualIndex = i + 1;
        itemTableHeader1ActualIndex = i + 3; // +1 untuk baris kosong, +1 untuk header item
        break;
      }
    }

    if (totalAllMoRowActualIndex !== -1) {
      // Merge TOTAL ALL SUPPLIER PER MO
      worksheet.mergeCells(totalAllMoRowActualIndex, 1, totalAllMoRowActualIndex, 5); // Label
      worksheet.getCell(totalAllMoRowActualIndex, 1).font = { bold: true };
      let col = 6;
      for (let i = 0; i < MONTH_ORDER.length; i++) {
        worksheet.mergeCells(totalAllMoRowActualIndex, col, totalAllMoRowActualIndex, col + 1);
        col += 2;
      }
      worksheet.mergeCells(totalAllMoRowActualIndex, col, totalAllMoRowActualIndex, col + 2); // Recap
      worksheet.getCell(totalAllMoRowActualIndex, col + 2).font = { bold: true }; // Total Qty Recap

      // Merge TOTAL ALL SUPPLIER PER QUARTAL
      worksheet.mergeCells(totalAllQuartalRowActualIndex, 1, totalAllQuartalRowActualIndex, 5); // Label
      worksheet.getCell(totalAllQuartalRowActualIndex, 1).font = { bold: true };
      col = 6;
      for (let q = 0; q < 4; q++) {
        worksheet.mergeCells(totalAllQuartalRowActualIndex, col, totalAllQuartalRowActualIndex, col + 5);
        col += 6;
      }
    }

    if (itemTableHeader1ActualIndex !== -1) {
      // Merge TOTAL PER ITEM Header
      worksheet.mergeCells(itemTableHeader1ActualIndex, 1, itemTableHeader1ActualIndex, 5); // Label
      worksheet.getCell(itemTableHeader1ActualIndex, 1).font = { bold: true };
      let col = 6;
      for (let i = 0; i < MONTH_ORDER.length; i++) {
        worksheet.mergeCells(itemTableHeader1ActualIndex, col, itemTableHeader1ActualIndex, col + 1);
        col += 2;
      } // Bulan
      worksheet.mergeCells(itemTableHeader1ActualIndex, col, itemTableHeader1ActualIndex, col + 2); // RECAP

      // Header kedua untuk TOTAL PER ITEM (Item-GSM-AddOn, QTY, TOTAL QTY)
      const itemTableHeader2ActualIndex = itemTableHeader1ActualIndex + 1;
      worksheet.mergeCells(itemTableHeader2ActualIndex, 1, itemTableHeader2ActualIndex, 5); // Label
      worksheet.getCell(itemTableHeader2ActualIndex, 1).font = { bold: true };
      col = 6;
      for (let i = 0; i < MONTH_ORDER.length; i++) {
        worksheet.mergeCells(itemTableHeader2ActualIndex, col, itemTableHeader2ActualIndex, col + 1);
        col += 2;
      } // QTY bulanan
      worksheet.mergeCells(itemTableHeader2ActualIndex, col, itemTableHeader2ActualIndex, col + 2); // RECAP TOTAL QTY
      worksheet.getCell(itemTableHeader2ActualIndex, col + 2).font = { bold: true };

      // Merge untuk setiap baris item di tabel TOTAL PER ITEM
      // Kolom pertama (Item-GSM-Addon) merge 5 sel
      let currentItemRow = itemTableHeader2ActualIndex + 1;
      while (currentItemRow <= worksheet.rowCount && worksheet.getRow(currentItemRow).getCell(6).value !== null) {
        // Asumsi kolom F (Jan QTY) akan ada isinya jika ini baris item
        if (
          worksheet.getRow(currentItemRow).getCell(1).value && // Pastikan bukan baris kosong
          worksheet.getRow(currentItemRow).getCell(1).value !== "TOTAL QTY PER MO" && // Bukan baris total lagi
          worksheet.getRow(currentItemRow).getCell(1).value !== "TOTAL QTY PER QUARTAL"
        ) {
          worksheet.mergeCells(currentItemRow, 1, currentItemRow, 5); // Merge A-E untuk nama item
          let itemMonthlyCol = 6;
          for (let i = 0; i < MONTH_ORDER.length; i++) {
            worksheet.mergeCells(currentItemRow, itemMonthlyCol, currentItemRow, itemMonthlyCol + 1); // Merge QTY per bulan
            itemMonthlyCol += 2;
          }
          worksheet.mergeCells(currentItemRow, itemMonthlyCol, currentItemRow, itemMonthlyCol + 2); // Merge RECAP
        }
        currentItemRow++;
      }
    }
  }

  if (workbook.worksheets.length > 0) {
    await workbook.xlsx.writeFile(outputFile);
    console.log(`\nProses selesai. Output disimpan di: ${outputFile}`);
  } else {
    console.log("\nTidak ada data yang diproses untuk output Excel.");
  }
}

module.exports = { prepareGroupBlock, writeOutputToFile };
