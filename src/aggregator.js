// src/aggregator.js
const { averageGreaterThanZero } = require("./utils");
const logger = require("./logger");

function performAggregation(data) {
  const monthlySummary = {};
  // console.log(`--- Memulai performAggregation untuk ${data.length} baris ---`);

  data.forEach((row, index) => {
    // Kolom yang WAJIB ada: month, hsCode.
    // Kolom gsm, item, addOn bisa '-' atau string kosong dan itu dianggap nilai valid untuk pengelompokan.
    if (!row.month || row.month === "-" || !row.hsCode || row.hsCode === "-") {
      // console.warn(`Aggregator: Baris ${index} dilewati karena bulan atau HS Code tidak valid:`, row);
      return;
    }

    // Jika GSM, ITEM, atau ADD ON tidak ada (null/undefined), kita anggap sebagai string kosong agar tetap bisa dikelompokkan.
    // Jika nilainya adalah string "-", itu akan diperlakukan sebagai nilai "-" yang unik.
    const gsmValue = row.gsm || ""; // Jika null/undefined, jadikan string kosong
    const itemValue = row.item || ""; // Jika null/undefined, jadikan string kosong
    const addOnValue = row.addOn || ""; // Jika null/undefined, jadikan string kosong

    const key = `${row.month}-${row.hsCode}-${itemValue}-${gsmValue}-${addOnValue}`;
    // console.log(`Aggregator: Membuat kunci: ${key} untuk baris:`, row);

    // // --- TAMBAHKAN LOG INI ---
    // if (row.hsCode === "56031200" && row.month === "Feb") {
    //   logger.debug(`Input untuk ${key}: importer = ${row.importer}, usdQtyUnit = ${row.usdQtyUnit}, qty = ${row.qty}`);
    // }

    if (!monthlySummary[key]) {
      monthlySummary[key] = {
        month: row.month,
        hsCode: row.hsCode,
        item: itemValue,
        gsm: gsmValue,
        addOn: addOnValue,
        usdQtyUnits: [],
        totalQty: 0,
      };
    }
    monthlySummary[key].usdQtyUnits.push(row.usdQtyUnit);
    monthlySummary[key].totalQty += row.qty;
  });

  if (Object.keys(monthlySummary).length === 0 && data.length > 0) {
    // console.warn("Aggregator: monthlySummary kosong meskipun ada data input. Periksa validitas bulan/HS Code pada semua baris input.");
  }

  const summaryLvl1Data = Object.values(monthlySummary).map((group) => ({
    month: group.month,
    hsCode: group.hsCode,
    item: group.item,
    gsm: group.gsm,
    addOn: group.addOn,
    avgPrice: averageGreaterThanZero(group.usdQtyUnits),
    totalQty: group.totalQty,
  }));

  if (summaryLvl1Data.length === 0 && data.length > 0) {
    // console.warn("Aggregator: summaryLvl1Data kosong. Tidak ada grup bulanan yang valid terbentuk.");
  }

  const recapSummary = {};
  summaryLvl1Data.forEach((row) => {
    const key = `${row.hsCode}-${row.item}-${row.gsm}-${row.addOn}`;
    if (!recapSummary[key]) {
      recapSummary[key] = {
        hsCode: row.hsCode,
        item: row.item,
        gsm: row.gsm,
        addOn: row.addOn,
        avgPrices: [],
        totalQty: 0,
      };
    }
    recapSummary[key].avgPrices.push(row.avgPrice);
    recapSummary[key].totalQty += row.totalQty;
  });

  const summaryLvl2Data = Object.values(recapSummary).map((group) => ({
    hsCode: group.hsCode,
    item: group.item,
    gsm: group.gsm,
    addOn: group.addOn,
    avgOfSummaryPrice: averageGreaterThanZero(group.avgPrices),
    totalOfSummaryQty: group.totalQty,
  }));

  if (summaryLvl2Data.length === 0 && summaryLvl1Data.length > 0) {
    // console.warn("Aggregator: summaryLvl2Data kosong meskipun summaryLvl1Data ada. Ini aneh.");
  }
  // console.log(`--- Selesai performAggregation, summaryLvl1: ${summaryLvl1Data.length}, summaryLvl2: ${summaryLvl2Data.length} ---`);

  return { summaryLvl1: summaryLvl1Data, summaryLvl2: summaryLvl2Data };
}

module.exports = { performAggregation };
