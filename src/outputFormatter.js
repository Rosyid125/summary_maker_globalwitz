// src/outputFormatter.js
const { MONTH_ORDER } = require("./utils");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const DEFAULT_OUTPUT_FOLDER = "processed_excel";

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
      totalQtyPerMoRow.push(total, null); // Nilai QTY di sel pertama area merge bulan
    });
    totalQtyPerMoRow.push(overallTotalQtyForThisGroup, null, null); // Nilai RECAP QTY di sel pertama area merge RECAP
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
    totalQtyPerQuartalRow.push(null, null, null); // RECAP kosong di baris ini
    groupBlockRows.push(totalQtyPerQuartalRow);
  }

  return {
    groupBlockRows: groupBlockRows,
    overallTotalQtyForGroup: overallTotalQtyForThisGroup,
    distinctCombinationsCount: distinctCombinations.length,
    headerRowCount: headerRowCount,
    header1Length: headerRow1.length,
  };
}

async function writeOutputToFile(workbookData, outputFileName = "summary_output.xlsx", periodYear) {
  if (!fs.existsSync(DEFAULT_OUTPUT_FOLDER)) {
    fs.mkdirSync(DEFAULT_OUTPUT_FOLDER, { recursive: true });
  }
  const outputFile = path.join(DEFAULT_OUTPUT_FOLDER, outputFileName);
  const workbook = new ExcelJS.Workbook();

  const colors = {
    period: "FF7030A0",
    supplierCols: "FF002060",
    q1: "FFFFC000",
    q2: "FF00B050",
    q3: "FFFFFF00",
    q4: "FF00B0F0",
    recap: "FF002060",
    textWhite: "FFFFFFFF",
    textBlack: "FF000000",
  };

  for (const sheetInfo of workbookData) {
    const worksheet = workbook.addWorksheet(sheetInfo.name);

    const periodTitleRowCell = worksheet.addRow([`${periodYear} PERIODE`]).getCell(1);
    worksheet.mergeCells(1, 1, 1, sheetInfo.totalColumns);
    periodTitleRowCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.period } };
    periodTitleRowCell.font = { bold: true, size: 14, color: { argb: colors.textWhite } };
    periodTitleRowCell.alignment = { vertical: "middle", horizontal: "center" };
    worksheet.getRow(1).height = 20;

    sheetInfo.allRowsForSheetContent.forEach((rowData, rowIndexInContent) => {
      const rowWithNAOrEmpty = rowData.map((cell) => (cell === null || typeof cell === "undefined" ? "" : cell));
      const addedRow = worksheet.addRow(rowWithNAOrEmpty);
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

    let currentSheetRowForBlocks = 2; // Mulai setelah baris judul periode
    for (const groupMeta of sheetInfo.supplierGroupsMeta) {
      const headerStartRow = currentSheetRowForBlocks;
      const dataStartRow = headerStartRow + groupMeta.headerRowCount;
      const productRows = groupMeta.productRowCount;

      const header1 = worksheet.getRow(headerStartRow);
      const header2 = worksheet.getRow(headerStartRow + 1);
      [header1, header2].forEach((r) => (r.font = { bold: true, color: { argb: colors.textWhite } }));
      for (let c = 1; c <= 5; c++) {
        header1.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.supplierCols } };
        header2.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.supplierCols } };
      }
      worksheet.mergeCells(headerStartRow, 1, headerStartRow + 1, 1);
      worksheet.mergeCells(headerStartRow, 2, headerStartRow + 1, 2);
      worksheet.mergeCells(headerStartRow, 3, headerStartRow + 1, 3);
      worksheet.mergeCells(headerStartRow, 4, headerStartRow + 1, 4);
      worksheet.mergeCells(headerStartRow, 5, headerStartRow + 1, 5);

      let currentColForColor = 6;
      for (let q = 0; q < 4; q++) {
        const qColor = q === 0 ? colors.q1 : q === 1 ? colors.q2 : q === 2 ? colors.q3 : colors.q4;
        const textColor = qColor === colors.q3 ? colors.textBlack : colors.textWhite;
        for (let i = 0; i < 3; i++) {
          worksheet.mergeCells(headerStartRow, currentColForColor, headerStartRow, currentColForColor + 1);
          const cellH1 = header1.getCell(currentColForColor);
          const cellH2_price = header2.getCell(currentColForColor);
          const cellH2_qty = header2.getCell(currentColForColor + 1);

          [cellH1, cellH2_price, cellH2_qty].forEach((c) => {
            c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: qColor } };
            c.font = { bold: true, color: { argb: textColor } };
          });
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
    let lastRowProcessed = currentSheetRowForBlocks - 1;
    if (sheetInfo.supplierGroupsMeta.length > 0 && !sheetInfo.supplierGroupsMeta[sheetInfo.supplierGroupsMeta.length - 1].hasFollowingGroup && sheetInfo.supplierGroupsMeta[sheetInfo.supplierGroupsMeta.length - 1].productRowCount > 0) {
      // Jika grup terakhir TIDAK diikuti baris kosong DAN punya produk (berarti ada 2 baris totalnya)
      // Maka currentSheetRowForBlocks sudah benar menunjuk ke baris KOSONG setelah blok terakhir
      // Tidak perlu penyesuaian di sini
    } else if (
      sheetInfo.supplierGroupsMeta.length > 0 &&
      !sheetInfo.supplierGroupsMeta[sheetInfo.supplierGroupsMeta.length - 1].hasFollowingGroup &&
      sheetInfo.supplierGroupsMeta[sheetInfo.supplierGroupsMeta.length - 1].productRowCount === 0
    ) {
      // Jika grup terakhir TIDAK diikuti baris kosong DAN TIDAK punya produk (hanya header)
      // Maka currentSheetRowForBlocks sudah benar
    } else if (sheetInfo.supplierGroupsMeta.length === 0) {
      // Tidak ada blok supplier sama sekali
      lastRowProcessed = 1; // Hanya ada baris periode
    }

    let totalAllHeaderMonthActualIndex = -1;
    let totalAllMoRowActualIndex = -1;
    let totalAllQuartalRowActualIndex = -1;
    let itemTableHeader1ActualIndex = -1;

    // Cari dari baris setelah blok supplier terakhir
    for (let i = lastRowProcessed + 1; i <= worksheet.rowCount; i++) {
      const cellAValue = worksheet.getRow(i).getCell(1).value ? String(worksheet.getRow(i).getCell(1).value) : "";
      if (cellAValue === "Month") {
        totalAllHeaderMonthActualIndex = i;
        totalAllMoRowActualIndex = i + 1;
        totalAllQuartalRowActualIndex = i + 2;
        break;
      }
    }
    for (let i = totalAllQuartalRowActualIndex > 0 ? totalAllQuartalRowActualIndex + 1 : lastRowProcessed + 1; i <= worksheet.rowCount; i++) {
      const cellAValue = worksheet.getRow(i).getCell(1).value ? String(worksheet.getRow(i).getCell(1).value) : "";
      if (cellAValue === "TOTAL PER ITEM") {
        itemTableHeader1ActualIndex = i;
        break;
      }
    }

    if (totalAllHeaderMonthActualIndex !== -1) {
      worksheet.mergeCells(totalAllHeaderMonthActualIndex, 1, totalAllHeaderMonthActualIndex, 5);
      worksheet.getCell(totalAllHeaderMonthActualIndex, 1).font = { bold: true };
      let col = 6;
      for (let i = 0; i < MONTH_ORDER.length; i++) {
        worksheet.mergeCells(totalAllHeaderMonthActualIndex, col, totalAllHeaderMonthActualIndex, col + 1);
        col += 2;
      }
      worksheet.mergeCells(totalAllHeaderMonthActualIndex, col, totalAllHeaderMonthActualIndex, col + 2);
      worksheet.getCell(totalAllHeaderMonthActualIndex, col).font = { bold: true };

      worksheet.mergeCells(totalAllMoRowActualIndex, 1, totalAllMoRowActualIndex, 5);
      worksheet.getCell(totalAllMoRowActualIndex, 1).font = { bold: true };
      col = 6;
      for (let i = 0; i < MONTH_ORDER.length; i++) {
        worksheet.mergeCells(totalAllMoRowActualIndex, col, totalAllMoRowActualIndex, col + 1);
        col += 2;
      }
      worksheet.mergeCells(totalAllMoRowActualIndex, col, totalAllMoRowActualIndex, col + 2);
      worksheet.getCell(totalAllMoRowActualIndex, col).font = { bold: true }; // Ini menunjuk ke AVG PRICE RECAP

      worksheet.mergeCells(totalAllQuartalRowActualIndex, 1, totalAllQuartalRowActualIndex, 5);
      worksheet.getCell(totalAllQuartalRowActualIndex, 1).font = { bold: true };
      col = 6;
      for (let q = 0; q < 4; q++) {
        worksheet.mergeCells(totalAllQuartalRowActualIndex, col, totalAllQuartalRowActualIndex, col + 5);
        col += 6;
      }
    }

    if (itemTableHeader1ActualIndex !== -1) {
      worksheet.mergeCells(itemTableHeader1ActualIndex, 1, itemTableHeader1ActualIndex, 5);
      worksheet.getCell(itemTableHeader1ActualIndex, 1).font = { bold: true };
      let col = 6;
      for (let i = 0; i < MONTH_ORDER.length; i++) {
        worksheet.mergeCells(itemTableHeader1ActualIndex, col, itemTableHeader1ActualIndex, col + 1);
        col += 2;
      }
      worksheet.mergeCells(itemTableHeader1ActualIndex, col, itemTableHeader1ActualIndex, col + 2);
      worksheet.getCell(itemTableHeader1ActualIndex, col).font = { bold: true };

      // Baris header kedua untuk TOTAL PER ITEM (Item-GSM-AddOn...) sudah dihapus di index.js
      // Jadi, baris berikutnya adalah data item
      let currentItemRow = itemTableHeader1ActualIndex + 1;
      while (currentItemRow <= worksheet.rowCount) {
        const firstCellVal = worksheet.getRow(currentItemRow).getCell(1).value;
        if (!firstCellVal || String(firstCellVal).trim() === "") break; // Berhenti jika baris kosong

        worksheet.mergeCells(currentItemRow, 1, currentItemRow, 5);
        let itemMonthlyCol = 6;
        for (let i = 0; i < MONTH_ORDER.length; i++) {
          worksheet.mergeCells(currentItemRow, itemMonthlyCol, currentItemRow, itemMonthlyCol + 1);
          itemMonthlyCol += 2;
        }
        worksheet.mergeCells(currentItemRow, itemMonthlyCol, currentItemRow, itemMonthlyCol + 2);
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
