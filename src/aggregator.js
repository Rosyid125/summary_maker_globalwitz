// src/aggregator.js
const { averageGreaterThanZero } = require("./utils");

function performAggregation(data) {
  const monthlySummary = {};
  data.forEach((row) => {
    // Pastikan month, hsCode, gsm, item, dan addOn valid sebelum membuat kunci
    if (
      !row.month ||
      row.month === "N/A" ||
      !row.hsCode ||
      row.hsCode === "N/A" ||
      !row.gsm ||
      row.gsm === "N/A" ||
      !row.item ||
      row.item === "N/A" || // Tambahkan pengecekan untuk item
      !row.addOn // addOn bisa string kosong, jadi cek keberadaannya saja (atau bisa juga row.addOn === 'N/A')
    ) {
      // console.warn("Baris dilewati karena data kunci tidak valid:", row);
      return;
    }
    // REVISI: Kunci sekarang mencakup ITEM dan ADD ON
    const key = `${row.month}-${row.hsCode}-${row.item}-${row.gsm}-${row.addOn}`;
    if (!monthlySummary[key]) {
      monthlySummary[key] = {
        month: row.month,
        hsCode: row.hsCode,
        item: row.item, // Simpan item
        gsm: row.gsm,
        addOn: row.addOn, // Simpan addOn
        usdQtyUnits: [],
        totalQty: 0,
      };
    }
    monthlySummary[key].usdQtyUnits.push(row.usdQtyUnit);
    monthlySummary[key].totalQty += row.qty;
  });

  const summaryLvl1Data = Object.values(monthlySummary).map((group) => ({
    month: group.month,
    hsCode: group.hsCode,
    item: group.item, // Sertakan item
    gsm: group.gsm,
    addOn: group.addOn, // Sertakan addOn
    avgPrice: averageGreaterThanZero(group.usdQtyUnits),
    totalQty: group.totalQty,
  }));

  const recapSummary = {};
  summaryLvl1Data.forEach((row) => {
    // REVISI: Kunci rekap juga mencakup ITEM dan ADD ON
    const key = `${row.hsCode}-${row.item}-${row.gsm}-${row.addOn}`;
    if (!recapSummary[key]) {
      recapSummary[key] = {
        hsCode: row.hsCode,
        item: row.item, // Simpan item
        gsm: row.gsm,
        addOn: row.addOn, // Simpan addOn
        avgPrices: [],
        totalQty: 0,
      };
    }
    recapSummary[key].avgPrices.push(row.avgPrice);
    recapSummary[key].totalQty += row.totalQty;
  });

  const summaryLvl2Data = Object.values(recapSummary).map((group) => ({
    hsCode: group.hsCode,
    item: group.item, // Sertakan item
    gsm: group.gsm,
    addOn: group.addOn, // Sertakan addOn
    avgOfSummaryPrice: averageGreaterThanZero(group.avgPrices),
    totalOfSummaryQty: group.totalQty,
  }));

  return { summaryLvl1: summaryLvl1Data, summaryLvl2: summaryLvl2Data };
}

module.exports = { performAggregation };
