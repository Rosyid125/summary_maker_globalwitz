// src/aggregator.js
const { averageGreaterThanZero } = require("./utils");

function performAggregation(data) {
  // 'data' di sini adalah 'groupData' dari index.js
  const monthlySummary = {};
  data.forEach((row) => {
    // Pastikan 'row.month' digunakan di sini
    if (!row.month || row.month === "N/A" || !row.hsCode || !row.gsm) {
      // console.warn("Baris dilewati karena bulan, HS Code, atau GSM tidak valid:", row);
      return;
    }
    const key = `${row.month}-${row.hsCode}-${row.gsm}`; // Menggunakan row.month
    if (!monthlySummary[key]) {
      monthlySummary[key] = {
        month: row.month, // Menyimpan month yang benar
        hsCode: row.hsCode,
        gsm: row.gsm,
        usdQtyUnits: [],
        totalQty: 0,
      };
    }
    monthlySummary[key].usdQtyUnits.push(row.usdQtyUnit);
    monthlySummary[key].totalQty += row.qty;
  });

  const summaryLvl1Data = Object.values(monthlySummary).map((group) => ({
    month: group.month, // 'month' di sini berasal dari kunci monthlySummary
    hsCode: group.hsCode,
    gsm: group.gsm,
    avgPrice: averageGreaterThanZero(group.usdQtyUnits),
    totalQty: group.totalQty,
  }));

  // ... (sisa kode aggregator sama, summaryLvl2 menggunakan hasil summaryLvl1) ...
  const recapSummary = {};
  summaryLvl1Data.forEach((row) => {
    // row di sini adalah item dari summaryLvl1Data, yang sudah punya 'month'
    const key = `${row.hsCode}-${row.gsm}`;
    if (!recapSummary[key]) {
      recapSummary[key] = {
        hsCode: row.hsCode,
        gsm: row.gsm,
        avgPrices: [],
        totalQty: 0,
      };
    }
    recapSummary[key].avgPrices.push(row.avgPrice);
    recapSummary[key].totalQty += row.totalQty;
  });

  const summaryLvl2Data = Object.values(recapSummary).map((group) => ({
    hsCode: group.hsCode,
    gsm: group.gsm,
    avgOfSummaryPrice: averageGreaterThanZero(group.avgPrices),
    totalOfSummaryQty: group.totalQty,
  }));

  return { summaryLvl1: summaryLvl1Data, summaryLvl2: summaryLvl2Data };
}

module.exports = { performAggregation };
