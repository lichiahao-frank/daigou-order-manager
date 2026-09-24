// 本機測試：用人工整理的「預期辨識結果」跑完整流程，產出 Excel
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const Core = require('../public/core.js');

(async () => {
  const src = path.join(__dirname, '../../吉伊卡哇登記_最新.xlsx');
  const existing = Core.readWorkbook(XLSX, fs.readFileSync(src));
  console.log('舊資料：商品', existing.products.length, '訂單', existing.orders.length, '其他分頁', existing.extras.map((e) => e.name));
  const results = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixture-results.json'), 'utf8'));
  const batch = '0924樂園';
  const { newProducts, newOrders } = Core.buildMatches(results, existing, batch);
  console.log('新商品', newProducts.map((p) => [p.name, p.jpy, p.price, p.include, p.warn]));
  console.log('新訂單', newOrders.map((o) => [o.customer, o.item, o.qty, o.price, o.include, o.warn]));
  const data = {
    products: [...existing.products, ...newProducts.filter((p) => p.include)],
    orders: [...existing.orders, ...newOrders.filter((o) => o.include)],
    extras: existing.extras,
  };
  const wb = Core.buildWorkbook(ExcelJS, data, { rate: 0.2, newBatches: [batch], generatedAt: '2026-09-24 12:00' });
  const out = path.join(__dirname, Core.outputFileName('吉伊卡哇登記_最新.xlsx', '20260924'));
  await wb.xlsx.writeFile(out);
  console.log('輸出', out);
})().catch((e) => { console.error(e); process.exit(1); });
