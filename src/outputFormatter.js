// src/outputFormatter.js
const { MONTH_ORDER } = require("./utils");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const DEFAULT_OUTPUT_FOLDER = "processed_excel";

// Fungsi prepareGroupBlock tetap sama
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
      totalQtyPerMoRow.push(Math.round(total), null);
    });
    totalQtyPerMoRow.push(Math.round(overallTotalQtyForThisGroup), null, null);
    groupBlockRows.push(totalQtyPerMoRow);

    const quarterlyTotals = [0, 0, 0, 0];
    monthlyTotals.forEach((total, index) => {
      if (index < 3) quarterlyTotals[0] += total;
      else if (index < 6) quarterlyTotals[1] += total;
      else if (index < 9) quarterlyTotals[2] += total;
      else quarterlyTotals[3] += total;
    });
    const totalQtyPerQuartalRow = ["TOTAL QTY PER QUARTAL", null, null, null, null];
    totalQtyPerQuartalRow.push(Math.round(quarterlyTotals[0]), null, null, null, null, null);
    totalQtyPerQuartalRow.push(Math.round(quarterlyTotals[1]), null, null, null, null, null);
    totalQtyPerQuartalRow.push(Math.round(quarterlyTotals[2]), null, null, null, null, null);
    totalQtyPerQuartalRow.push(Math.round(quarterlyTotals[3]), null, null, null, null, null);
    totalQtyPerQuartalRow.push(null, null, null);
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
    totalPerItemTitle: "FFFF0000",
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

    let currentSheetRowForBlocks = 2;
    for (const groupMeta of sheetInfo.supplierGroupsMeta) {
      const headerStartRow = currentSheetRowForBlocks;
      const dataStartRow = headerStartRow + groupMeta.headerRowCount;
      const productRows = groupMeta.productRowCount;

      const header1 = worksheet.getRow(headerStartRow);
      const header2 = worksheet.getRow(headerStartRow + 1);
      [header1, header2].forEach((r) => {
        r.font = { bold: true, color: { argb: colors.textWhite } };
        r.eachCell((c) => (c.alignment = { vertical: "middle", horizontal: "center" }));
      });

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
      const quarterBgColors = [colors.q1, colors.q2, colors.q3, colors.q4];
      for (let q = 0; q < 4; q++) {
        const qBgColor = quarterBgColors[q];
        const textColor = qBgColor === colors.q3 ? colors.textBlack : colors.textWhite;
        for (let i = 0; i < 3; i++) {
          worksheet.mergeCells(headerStartRow, currentColForColor, headerStartRow, currentColForColor + 1);
          const cellH1 = header1.getCell(currentColForColor);
          const cellH2_price = header2.getCell(currentColForColor);
          const cellH2_qty = header2.getCell(currentColForColor + 1);

          [cellH1, cellH2_price, cellH2_qty].forEach((c) => {
            c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: qBgColor } };
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
        worksheet.getCell(dataStartRow, 1).alignment = { vertical: "middle", horizontal: "center" };

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

    let totalAllHeaderMonthActualIndex = -1;
    let totalAllMoRowActualIndex = -1;
    let totalAllQuartalRowActualIndex = -1;
    let itemTableMainTitleActualIndex = -1;
    let itemTableHeaderMonthActualIndex = -1;

    let searchStartRow = sheetInfo.supplierGroupsMeta.length > 0 ? currentSheetRowForBlocks : 2;

    for (let i = searchStartRow; i <= worksheet.rowCount; i++) {
      const cellAValue = worksheet.getRow(i).getCell(1).value ? String(worksheet.getRow(i).getCell(1).value).trim() : "";
      if (cellAValue === "Month" && totalAllHeaderMonthActualIndex === -1) {
        const nextRowCellA = worksheet.getRow(i + 1).getCell(1).value ? String(worksheet.getRow(i + 1).getCell(1).value).trim() : "";
        if (nextRowCellA === "TOTAL ALL SUPPLIER PER MO") {
          totalAllHeaderMonthActualIndex = i;
          totalAllMoRowActualIndex = i + 1;
          totalAllQuartalRowActualIndex = i + 2;
        }
      } else if (cellAValue === "TOTAL PER ITEM" && itemTableMainTitleActualIndex === -1) {
        itemTableMainTitleActualIndex = i;
        itemTableHeaderMonthActualIndex = i + 1;
      }
    }

    // REVISI: Definisi recapStartColTAS di sini agar bisa diakses
    const recapStartColTAS = 5 + MONTH_ORDER.length * 2 + 1;

    if (totalAllHeaderMonthActualIndex !== -1) {
      const thmRow = worksheet.getRow(totalAllHeaderMonthActualIndex);
      thmRow.font = { bold: true, color: { argb: colors.textWhite } };
      worksheet.mergeCells(totalAllHeaderMonthActualIndex, 1, totalAllHeaderMonthActualIndex, 5);
      thmRow.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.supplierCols } };
      let col = 6;
      const quarterBgColorsTAS = [colors.q1, colors.q2, colors.q3, colors.q4];
      for (let q = 0; q < 4; q++) {
        const qBgColor = quarterBgColorsTAS[q];
        const textColor = qBgColor === colors.q3 ? colors.textBlack : colors.textWhite;
        for (let i = 0; i < 3; i++) {
          worksheet.mergeCells(totalAllHeaderMonthActualIndex, col, totalAllHeaderMonthActualIndex, col + 1);
          thmRow.getCell(col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: qBgColor } };
          thmRow.getCell(col).font = { bold: true, color: { argb: textColor } };
          thmRow.getCell(col + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: qBgColor } };
          col += 2;
        }
      }
      worksheet.mergeCells(totalAllHeaderMonthActualIndex, recapStartColTAS, totalAllHeaderMonthActualIndex, recapStartColTAS + 2);
      for (let i = 0; i < 3; i++) {
        thmRow.getCell(recapStartColTAS + i).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.recap } };
      }
      thmRow.getCell(recapStartColTAS).font = { bold: true, color: { argb: colors.textWhite } };

      const tamRow = worksheet.getRow(totalAllMoRowActualIndex);
      worksheet.mergeCells(totalAllMoRowActualIndex, 1, totalAllMoRowActualIndex, 5);
      col = 6;
      for (let i = 0; i < MONTH_ORDER.length; i++) {
        worksheet.mergeCells(totalAllMoRowActualIndex, col, totalAllMoRowActualIndex, col + 1);
        col += 2;
      }
      // Pewarnaan baris "TOTAL ALL SUPPLIER PER MO" sampai RECAP
      for (let c = 1; c <= recapStartColTAS + 2; c++) {
        tamRow.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.period } };
      }
      tamRow.font = { bold: true, color: { argb: colors.textWhite } };

      const taqRow = worksheet.getRow(totalAllQuartalRowActualIndex);
      worksheet.mergeCells(totalAllQuartalRowActualIndex, 1, totalAllQuartalRowActualIndex, 5);
      col = 6;
      for (let q = 0; q < 4; q++) {
        worksheet.mergeCells(totalAllQuartalRowActualIndex, col, totalAllQuartalRowActualIndex, col + 5);
        col += 6;
      }
      // Pewarnaan baris "TOTAL ALL SUPPLIER PER QUARTAL" sampai RECAP
      for (let c = 1; c <= recapStartColTAS + 2; c++) {
        taqRow.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.period } };
      }
      taqRow.font = { bold: true, color: { argb: colors.textWhite } };

      worksheet.mergeCells(totalAllMoRowActualIndex, recapStartColTAS, totalAllQuartalRowActualIndex, recapStartColTAS + 2);
      const mergedRecapTASCell = worksheet.getCell(totalAllMoRowActualIndex, recapStartColTAS);
      mergedRecapTASCell.font = { bold: true, color: { argb: colors.textWhite } };
      // Warna latar sudah diatur per baris, jadi tidak perlu override fill di sini KECUALI jika ingin warna recap biru
      mergedRecapTASCell.alignment = { vertical: "middle", horizontal: "center" };
    }

    if (itemTableMainTitleActualIndex !== -1) {
      const itmtRow = worksheet.getRow(itemTableMainTitleActualIndex);
      worksheet.mergeCells(itemTableMainTitleActualIndex, 1, itemTableMainTitleActualIndex, sheetInfo.totalColumns);
      itmtRow.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.totalPerItemTitle } };
      itmtRow.getCell(1).font = { bold: true, size: 12, color: { argb: colors.textWhite } };
      itmtRow.height = 18;

      const itemHeaderMonthRowActual = worksheet.getRow(itemTableHeaderMonthActualIndex);
      itemHeaderMonthRowActual.font = { bold: true, color: { argb: colors.textWhite } };
      worksheet.mergeCells(itemTableHeaderMonthActualIndex, 1, itemTableHeaderMonthActualIndex, 5);
      itemHeaderMonthRowActual.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.supplierCols } };
      let col = 6;
      const quarterBgColorsTPI = [colors.q1, colors.q2, colors.q3, colors.q4];
      for (let q = 0; q < 4; q++) {
        const qBgColor = quarterBgColorsTPI[q];
        const textColor = qBgColor === colors.q3 ? colors.textBlack : colors.textWhite;
        for (let i = 0; i < 3; i++) {
          worksheet.mergeCells(itemTableHeaderMonthActualIndex, col, itemTableHeaderMonthActualIndex, col + 1);
          itemHeaderMonthRowActual.getCell(col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: qBgColor } };
          itemHeaderMonthRowActual.getCell(col).font = { bold: true, color: { argb: textColor } };
          itemHeaderMonthRowActual.getCell(col + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: qBgColor } };
          col += 2;
        }
      }
      const recapHeaderItemColStart = col;
      worksheet.mergeCells(itemTableHeaderMonthActualIndex, recapHeaderItemColStart, itemTableHeaderMonthActualIndex, recapHeaderItemColStart + 2);
      for (let i = 0; i < 3; i++) {
        itemHeaderMonthRowActual.getCell(recapHeaderItemColStart + i).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.recap } };
      }
      itemHeaderMonthRowActual.getCell(recapHeaderItemColStart).font = { bold: true, color: { argb: colors.textWhite } };

      let currentItemRow = itemTableHeaderMonthActualIndex + 1;
      while (currentItemRow <= worksheet.rowCount) {
        const firstCell = worksheet.getRow(currentItemRow).getCell(1);
        if (
          !firstCell.value ||
          String(firstCell.value).trim() === "" ||
          String(firstCell.value).startsWith("TOTAL ALL SUPPLIER") ||
          String(firstCell.value).startsWith("Month") ||
          String(firstCell.value).startsWith("TOTAL QTY PER MO") ||
          String(firstCell.value).startsWith("TOTAL QTY PER QUARTAL") ||
          String(firstCell.value).startsWith("TOTAL PER ITEM")
        ) {
          if (String(firstCell.value).trim() !== "" && !String(firstCell.value).startsWith("TOTAL PER ITEM")) {
            break;
          } else if (String(firstCell.value).trim() === "" && currentItemRow > itemTableHeaderMonthActualIndex + 1) {
            break;
          }
        }
        if (String(firstCell.value).trim() !== "" && !String(firstCell.value).startsWith("TOTAL PER ITEM")) {
          worksheet.mergeCells(currentItemRow, 1, currentItemRow, 5);
          let itemMonthlyCol = 6;
          for (let i = 0; i < MONTH_ORDER.length; i++) {
            worksheet.mergeCells(currentItemRow, itemMonthlyCol, currentItemRow, itemMonthlyCol + 1);
            itemMonthlyCol += 2;
          }
          worksheet.mergeCells(currentItemRow, itemMonthlyCol, currentItemRow, itemMonthlyCol + 2);
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
