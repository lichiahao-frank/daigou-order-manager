/*
 * 代購訂單管理：核心邏輯（瀏覽器與 Node 共用）
 * - readWorkbook：用 SheetJS 讀舊的登記 Excel
 * - buildMatches：把 AI 辨識結果對到商品、檢查金額與重複
 * - buildWorkbook：用 ExcelJS 產生新的 Excel（Dashboard / Product / Orders + 保留其他分頁）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.OrderCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const ORDER_HEADERS = ['批次', '商品品項', '售價', '客人姓名', '數量', '小計', '境內運分擔方式', '交易方式', '是否匯款', '訂單來源', '狀態'];
  const PRODUCT_HEADERS = ['批次', '商品名稱', '日幣成本', '成本', '售價', '庫存量', '日文品名', '未出貨需求量'];
  const MAIN_SHEETS = ['Dashboard', 'Product', 'Orders'];
  // Excel 保留字，不能當分頁名稱
  const RESERVED_SHEET_NAMES = { history: 'History_出貨歷史' };

  const OPTIONS = {
    shipping: ['分擔', '不分擔', '自吃'],
    payMethod: ['匯款', '貨付', '無卡', '貨到付款'],
    payStatus: ['未匯款', '已匯款', '未交證明', '已交證明'],
    source: ['社群', 'FB', 'IG', '脆'],
    status: ['未出貨', '已出貨', '蹲不限購', '取消', '跑單'],
  };

  // ---------- 小工具 ----------
  function blank(v) {
    return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  }
  function num(v) {
    if (typeof v === 'number') return v;
    if (blank(v)) return null;
    const n = Number(String(v).replace(/[,$＄¥￥\s]/g, ''));
    return Number.isFinite(n) ? n : null;
  }
  function toHalfWidth(s) {
    return String(s).replace(/[！-～]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).replace(/　/g, ' ');
  }
  // 比對用的正規化：全形轉半形、去空白、統一括號與常見異體字
  function norm(s) {
    if (blank(s)) return '';
    return toHalfWidth(s)
      .replace(/\s+/g, '')
      .replace(/[（]/g, '(')
      .replace(/[）]/g, ')')
      .replace(/灵/g, '靈')
      .replace(/獅子/g, '獅薩')
      .toLowerCase();
  }
  function levenshtein(a, b) {
    const m = a.length, n = b.length;
    if (!m) return n;
    if (!n) return m;
    let prev = Array.from({ length: n + 1 }, (_, j) => j);
    for (let i = 1; i <= m; i++) {
      const cur = [i];
      for (let j = 1; j <= n; j++) {
        cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      }
      prev = cur;
    }
    return prev[n];
  }

  // 數字欄：是數字就轉數字，不是（例如庫存寫「S+1」）就原樣保留，避免舊資料遺失
  function numOrRaw(v) {
    if (blank(v)) return null;
    const n = num(v);
    return n === null ? v : n;
  }

  // ---------- 讀舊檔 ----------
  function sheetRows(XLSX, wb, name) {
    const ws = wb.Sheets[name];
    if (!ws) return null;
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  }
  function findSheetName(wb, wanted) {
    const w = wanted.toLowerCase();
    return wb.SheetNames.find((n) => n.toLowerCase() === w) ||
      wb.SheetNames.find((n) => n.toLowerCase().replace(/s$/, '') === w.replace(/s$/, ''));
  }

  function readWorkbook(XLSX, data) {
    const wb = XLSX.read(data, { type: data instanceof ArrayBuffer ? 'array' : 'buffer', cellDates: false, cellNF: true });
    const productName = findSheetName(wb, 'Product');
    const orderName = findSheetName(wb, 'Orders');
    const products = [];
    const orders = [];

    if (productName) {
      const rows = sheetRows(XLSX, wb, productName);
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i] || [];
        if (blank(r[0]) && blank(r[1])) continue;
        products.push({
          batch: blank(r[0]) ? '' : String(r[0]).trim(),
          name: blank(r[1]) ? '' : String(r[1]).trim(),
          jpy: numOrRaw(r[2]),
          cost: numOrRaw(r[3]),
          price: numOrRaw(r[4]),
          stock: numOrRaw(r[5]),
          jpName: blank(r[6]) ? '' : String(r[6]),
          isNew: false,
        });
      }
    }
    if (orderName) {
      const rows = sheetRows(XLSX, wb, orderName);
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i] || [];
        // 舊檔有預填預設值的空白範本列：批次、品項、客人都空就不算訂單
        if (blank(r[0]) && blank(r[1]) && blank(r[3])) continue;
        orders.push({
          batch: blank(r[0]) ? '' : String(r[0]).trim(),
          item: blank(r[1]) ? '' : String(r[1]).trim(),
          price: numOrRaw(r[2]),
          customer: blank(r[3]) ? '' : String(r[3]),
          qty: numOrRaw(r[4]),
          subtotal: num(r[5]),
          shipping: r[6] ?? '',
          payMethod: r[7] ?? '',
          payStatus: r[8] ?? '',
          source: r[9] ?? '',
          status: r[10] ?? '',
          isNew: false,
        });
      }
    }
    // 其他分頁照原樣保留（只保留數值，不保留 Google 專用公式）
    const extras = [];
    for (const name of wb.SheetNames) {
      if (name === productName || name === orderName || name.toLowerCase() === 'dashboard') continue;
      const rows = sheetRows(XLSX, wb, name) || [];
      let last = rows.length;
      while (last > 0 && (rows[last - 1] || []).every(blank)) last--;
      // 保留數字格式（例如日期），不然日期會變成 46206.93 這種序號
      const formats = {};
      const ws = wb.Sheets[name];
      for (const addr of Object.keys(ws)) {
        if (addr[0] === '!') continue;
        const c = ws[addr];
        if (c.t === 'n' && c.z && c.z !== 'General') formats[addr] = c.z;
      }
      extras.push({ name, rows: rows.slice(0, last), formats });
    }
    return { products, orders, extras, hasProduct: !!productName, hasOrders: !!orderName };
  }

  function batchList(products, orders) {
    const seen = new Map();
    for (const r of [...products, ...orders]) {
      if (!blank(r.batch) && !seen.has(r.batch)) seen.set(r.batch, true);
    }
    return [...seen.keys()];
  }

  // ---------- 解析 OCR 結果 ----------
  // lines: OCR 回傳的 [[{text, box:{x,y,width,height}, confidence}], ...]，每個內層陣列是一行
  // 回傳跟 buildMatches 相容的 {kind, products, orders, notes}
  const KANA_RE = /[぀-ヿ]/;
  const HAN_RE = /[一-鿿]/;
  const PRICE_RE = /[¥￥]\s*([\d][\d,]{2,})/;
  const CARD_LABEL_RE = /販.{0,2}開始|販売|予約|受付|在庫|SOLD|NEW|税込/i;
  const TIME_RE = /(\d+\s*(秒|分鐘|分|小時|天|週|周|個月|年)前|剛剛|昨天|今天|前天|(上午|下午|晚上|凌晨|中午|早上)\s*\d{1,2}[:：]\d{2}|^\d{1,2}[:：]\d{2}$|^\d{1,2}\/\d{1,2}(\s+\d{1,2}[:：]\d{2})?$)/;
  const TOTAL_RE = /金[額额]|總[計计額额]|合計|總共|小計/;
  const SHIP_WORDS = [['不分擔', /不分[擔担]/], ['分擔', /分[擔担]/], ['自吃', /自吃/]];
  const ITEM_PREFIX_RE = /^(商品|品項|品名|購買|訂購|購買商品|我要)\s*[:：]\s*/;
  const ITEM_RE = /([^\s:：,，、;；+＋xX×*]{1,24}?)\s*[+＋xX×*]\s*(\d{1,3})(?![\d,])/g;
  const ITEM_UNIT_RE = /([^\s:：,，、;；]{1,24}?)\s*(\d{1,3})\s*(個|隻|件|盒|入|組|份|抽)/g;

  const lineText = (line) => line.map((w) => w.text).join(' ').replace(/\s+/g, ' ').trim();

  function parseShipping(t) {
    for (const [v, re] of SHIP_WORDS) if (re.test(t)) return v;
    return '';
  }
  function parseItems(t) {
    const s = toHalfWidth(t).replace(ITEM_PREFIX_RE, '').trim();
    const items = [];
    for (const re of [ITEM_RE, ITEM_UNIT_RE]) {
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(s))) items.push({ item: m[1].replace(ITEM_PREFIX_RE, '').trim(), qty: Number(m[2]) });
      if (items.length) break;
    }
    return items.filter((it) => it.item && !TOTAL_RE.test(it.item));
  }
  function classifyLine(t) {
    if (!t) return 'empty';
    if (TIME_RE.test(t) && t.length <= 24 && !/[+＋]\s*\d/.test(t)) return 'time';
    if (TOTAL_RE.test(t) && /\d/.test(t)) return 'total';
    if (parseShipping(t) && (/運|郵資/.test(t) || t.replace(/[\s:：]/g, '').length <= 4)) return 'ship';
    if (parseItems(t).length) return 'item';
    return 'other';
  }
  // 名字那一行常混進頭像的雜訊（例如「oo 客人B」），取最後一個像名字的詞
  function pickName(line) {
    const words = line.map((w) => String(w.text || '').trim()).filter((w) => w && /[\p{L}\p{N}]/u.test(w) && !/^(o+|c|x|\(?\d\)?)$/i.test(w));
    return words.length ? words[words.length - 1] : '';
  }

  function parseComments(lines) {
    const orders = [];
    let block = [];
    const flush = () => {
      const tagged = block.map((line) => ({ line, t: lineText(line), k: classifyLine(lineText(line)) }));
      block = [];
      const first = tagged.findIndex((x) => x.k === 'item');
      if (first < 0) return;
      const nameRow = first > 0 && tagged[first - 1].k === 'other' ? tagged[first - 1] : null;
      const o = { customer: nameRow ? pickName(nameRow.line) : '', items: [], stated_total: 0, shipping: '', raw_text: '' };
      const raw = [];
      for (const x of tagged.slice(first)) {
        if (x.k === 'item') o.items.push(...parseItems(x.t));
        else if (x.k === 'total') o.stated_total = Number((toHalfWidth(x.t).match(/(\d[\d,]*)(?!.*\d)/) || [0, '0'])[1].replace(/,/g, '')) || 0;
        else if (x.k === 'ship') o.shipping = parseShipping(x.t);
        else continue;
        raw.push(x.t);
      }
      o.raw_text = raw.join(' ');
      orders.push(o);
    };
    for (const line of lines) {
      if (classifyLine(lineText(line)) === 'time') flush();
      else block.push(line);
    }
    flush();
    return orders;
  }

  // 型錄：依 x 座標分欄，每欄由上往下「日文品名…¥價格」組成一個商品
  function parseCatalog(lines) {
    const words = [];
    for (const line of lines) for (const w of line) if (w && w.box && String(w.text || '').trim()) words.push(w);
    const W = Math.max(1, ...words.map((w) => w.box.x + w.box.width));
    const anchors = words.filter((w) => KANA_RE.test(w.text) || PRICE_RE.test(w.text)).map((w) => w.box.x + w.box.width / 2).sort((a, b) => a - b);
    if (!anchors.length) return [];
    const cols = [[anchors[0]]];
    for (let i = 1; i < anchors.length; i++) {
      if (anchors[i] - anchors[i - 1] > W * 0.15) cols.push([]);
      cols[cols.length - 1].push(anchors[i]);
    }
    const centers = cols.map((c) => c.reduce((s, v) => s + v, 0) / c.length);
    const colOf = (w) => {
      const cx = w.box.x + w.box.width / 2;
      let best = 0;
      centers.forEach((c, i) => { if (Math.abs(cx - c) < Math.abs(cx - centers[best])) best = i; });
      return best;
    };
    const products = [];
    centers.forEach((_, ci) => {
      const colWords = words.filter((w) => colOf(w) === ci).sort((a, b) => a.box.y - b.box.y);
      let name = [];
      let zh = '';
      let last = null; // 剛結束的商品與它價格的位置，手寫字常跟價格在同一高度
      const flush = (jpy, priceWord) => {
        if (!name.length && !jpy) return;
        const p = { name_zh: zh, guessed: !!zh, name_jp: name.join('').replace(/\s+/g, ''), jpy: jpy || 0, price_twd: 0 };
        products.push(p);
        last = priceWord ? { p, y: priceWord.box.y, h: priceWord.box.height } : null;
        name = [];
        zh = '';
      };
      for (const w of colWords) {
        const t = String(w.text).trim();
        const price = t.match(PRICE_RE);
        if (price) { flush(Number(price[1].replace(/,/g, '')), w); continue; }
        if (!name.length && last && !last.p.name_zh && HAN_RE.test(t) && !KANA_RE.test(t) && !CARD_LABEL_RE.test(t) && Math.abs(w.box.y - last.y) < last.h * 2.5) {
          const guess = t.replace(/[\d$＄#＃¥￥,.\s()（）]/g, '');
          if (guess) { last.p.name_zh = guess; last.p.guessed = true; }
          continue;
        }
        if (/^販.{0,2}開始前?$|^販売/.test(t)) { if (name.length) flush(0); continue; }
        if (KANA_RE.test(t)) { name.push(t); continue; }
        // 手寫中文：有漢字、沒有假名，拿掉數字與符號當作中文名稱的參考
        if (HAN_RE.test(t) && !CARD_LABEL_RE.test(t) && name.length) {
          const guess = t.replace(/[\d$＄#＃¥￥,.\s()（）]/g, '');
          if (guess && !zh) zh = guess;
        }
      }
      flush(0);
    });
    return products.filter((p) => p.name_jp.length >= 4);
  }

  function parseOcr(lines) {
    const texts = lines.map(lineText);
    const priceCount = texts.filter((t) => PRICE_RE.test(t)).length;
    const kanaCount = texts.filter((t) => KANA_RE.test(t)).length;
    const orders = parseComments(lines);
    const products = priceCount >= 1 && kanaCount >= 2 ? parseCatalog(lines) : [];
    const kind = products.length && orders.length ? 'mixed' : products.length ? 'catalog' : orders.length ? 'orders' : 'other';
    const notes = [];
    if (products.length) notes.push('型錄的中文名稱與台幣售價請在核對表確認（手寫字辨識不可靠）');
    if (kind === 'other') notes.push('沒有找到「品名+數量」格式的留言或 ¥ 價格');
    return { kind, products, orders, notes: notes.join('；'), text: texts.join('\n') };
  }

  // ---------- 比對辨識結果 ----------
  // 找最像的商品：同批次優先，完全相同 > 包含 > 編輯距離 1
  function matchProduct(itemName, candidates) {
    const q = norm(itemName);
    if (!q) return null;
    let hit = candidates.find((p) => norm(p.name) === q);
    if (hit) return { product: hit, how: 'exact' };
    const contains = candidates.filter((p) => {
      const n = norm(p.name);
      return n && Math.min(n.length, q.length) >= 2 && (n.includes(q) || q.includes(n));
    });
    if (contains.length === 1) return { product: contains[0], how: 'fuzzy' };
    const near = candidates
      .map((p) => ({ p, d: levenshtein(norm(p.name), q) }))
      .filter((x) => x.d <= 1)
      .sort((a, b) => a.d - b.d);
    if (near.length === 1 || (near.length > 1 && near[0].d < near[1].d)) return { product: near[0].p, how: 'fuzzy' };
    return null;
  }

  /**
   * results: 每張截圖的 AI 辨識結果 [{fileName, kind, products:[...], orders:[...]}]
   * existing: readWorkbook 的結果
   * batch: 這次的批次名稱
   * 回傳 { newProducts, newOrders } 供畫面核對
   */
  function buildMatches(results, existing, batch) {
    const newProducts = [];
    const batchNorm = norm(batch);
    const existingInBatch = existing.products.filter((p) => norm(p.batch) === batchNorm);

    for (const res of results) {
      for (const p of res.products || []) {
        const name = String(p.name_zh || '').trim();
        if (!name && !p.name_jp) continue;
        const dup = name && (existingInBatch.find((x) => norm(x.name) === norm(name)) || newProducts.find((x) => norm(x.name) === norm(name)));
        if (dup && dup.isNew) continue; // 同一次上傳重複出現的商品只收一次
        const warns = [];
        if (dup) warns.push('此批次已有同名商品，預設不重複新增');
        if (!name) warns.push('請填中文名稱');
        else if (p.guessed) warns.push('中文名稱是從手寫字猜的，請確認');
        if (!num(p.price_twd)) warns.push('請填台幣售價');
        if (!num(p.jpy)) warns.push('沒讀到日幣價格');
        newProducts.push({
          include: !dup,
          batch,
          name,
          jpName: p.name_jp || '',
          guessed: !!p.guessed,
          jpy: num(p.jpy) || null,
          price: num(p.price_twd) || null,
          isNew: true,
          warn: warns.join('；'),
          fileName: res.fileName,
        });
      }
    }

    // 手寫字猜的名稱不可靠，確認前不拿來配對訂單
    const candidates = [...existingInBatch, ...newProducts.filter((p) => p.include && p.name && !p.guessed)];
    const newOrders = [];
    for (const res of results) {
      for (const o of res.orders || []) {
        const items = (o.items || []).filter((it) => !blank(it.item));
        const lines = [];
        for (const it of items) {
          const m = matchProduct(it.item, candidates);
          const qty = num(it.qty) || 1;
          lines.push({
            include: true,
            batch,
            item: m ? m.product.name : String(it.item).trim(),
            rawItem: String(it.item).trim(),
            matched: m ? m.how : 'none',
            price: m ? m.product.price : null,
            customer: String(o.customer || '').trim(),
            qty,
            shipping: OPTIONS.shipping.includes(o.shipping) ? o.shipping : (o.shipping ? '分擔' : ''),
            payMethod: '',
            payStatus: '未匯款',
            source: '社群',
            status: '未出貨',
            statedTotal: num(o.stated_total),
            rawText: o.raw_text || '',
            fileName: res.fileName,
            isNew: true,
            warn: '',
          });
        }
        // 一則留言的總金額要跟所有品項加總比對
        const sum = lines.reduce((s, l) => s + (l.price || 0) * l.qty, 0);
        for (const l of lines) {
          const warns = [];
          if (l.matched === 'none') warns.push(`找不到「${l.rawItem}」這個商品，請選正確品項`);
          else if (l.matched === 'fuzzy' && l.rawItem !== l.item) warns.push(`「${l.rawItem}」自動對到「${l.item}」，請確認`);
          if (!l.customer) warns.push('沒辨識到客人名字');
          if (l.statedTotal && lines.every((x) => x.price) && Math.abs(sum - l.statedTotal) > 0.5) {
            warns.push(`留言寫總金額 ${l.statedTotal}，但依售價算出 ${sum}`);
          }
          const dup = existing.orders.find((x) => norm(x.batch) === norm(l.batch) && norm(x.item) === norm(l.item) && norm(x.customer) === norm(l.customer) && x.status !== '取消');
          if (dup) {
            warns.push('舊資料已有同客人同品項的訂單，可能重複（預設不勾選）');
            l.include = false;
          }
          l.warn = warns.join('；');
        }
        newOrders.push(...lines);
      }
    }
    return { newProducts, newOrders };
  }

  // ---------- 產生 Excel ----------
  const COLORS = {
    header: 'FF7A4B94',
    headerFont: 'FFFFFFFF',
    title: 'FF4A2A5E',
    band: 'FFF6F0FA',
    newRow: 'FFFFF7D6',
    kpiFill: 'FFFBF3FF',
    border: 'FFD9CCE3',
    pending: 'FFFFE6CC',
    unpaid: 'FFC0392B',
    muted: 'FF8A7F92',
  };
  const thin = { style: 'thin', color: { argb: COLORS.border } };
  const box = { top: thin, left: thin, bottom: thin, right: thin };

  function styleHeader(row) {
    row.eachCell((c) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.header } };
      c.font = { bold: true, color: { argb: COLORS.headerFont } };
      c.alignment = { vertical: 'middle', horizontal: 'center' };
      c.border = box;
    });
    row.height = 22;
  }

  function safeSheetName(name, used) {
    let n = String(name).replace(/[\\/?*[\]:]/g, '_').slice(0, 31);
    if (RESERVED_SHEET_NAMES[n.toLowerCase()]) n = RESERVED_SHEET_NAMES[n.toLowerCase()];
    let base = n, i = 2;
    while (used.has(n.toLowerCase())) n = `${base.slice(0, 28)}_${i++}`;
    used.add(n.toLowerCase());
    return n;
  }

  /**
   * data: { products, orders, extras }（products/orders 已含新資料，新資料 isNew=true）
   * opts: { rate, newBatches, generatedAt }
   */
  function buildWorkbook(ExcelJS, data, opts) {
    const rate = opts.rate || 0.2;
    const wb = new ExcelJS.Workbook();
    wb.creator = '代購訂單管理';
    wb.calcProperties.fullCalcOnLoad = true;
    const used = new Set(MAIN_SHEETS.map((s) => s.toLowerCase()));

    const dash = wb.addWorksheet('Dashboard', { views: [{ showGridLines: false }], properties: { tabColor: { argb: COLORS.header } } });
    const prod = wb.addWorksheet('Product', { views: [{ state: 'frozen', ySplit: 1 }] });
    const ord = wb.addWorksheet('Orders', { views: [{ state: 'frozen', ySplit: 1 }] });

    const products = data.products;
    const orders = data.orders;
    const batches = batchList(products, orders);
    const ORDER_LAST = Math.max(orders.length + 1 + 300, 2000); // 預留空白列給手動登記
    const PROD_LAST = Math.max(products.length + 1 + 200, 1000);
    const oRange = (col) => `Orders!$${col}$2:$${col}$${ORDER_LAST}`;
    const pRange = (col) => `Product!$${col}$2:$${col}$${PROD_LAST}`;

    // 商品查價表：用來判斷舊訂單的售價能不能改成自動查價公式
    const priceKey = new Map();
    const dupKey = new Set();
    for (const p of products) {
      const k = norm(p.batch) + '|' + norm(p.name);
      if (priceKey.has(k)) dupKey.add(k);
      priceKey.set(k, p.price || 0);
    }

    // ----- Product -----
    prod.columns = [
      { width: 14 }, { width: 24 }, { width: 11 }, { width: 10 }, { width: 10 }, { width: 9 }, { width: 40 }, { width: 13 },
    ];
    prod.addRow(PRODUCT_HEADERS);
    styleHeader(prod.getRow(1));
    products.forEach((p, i) => {
      const r = i + 2;
      const row = prod.getRow(r);
      row.getCell(1).value = p.batch || null;
      row.getCell(2).value = p.name || null;
      row.getCell(3).value = p.jpy ?? null;
      // 新商品：成本用公式（日幣 × 匯率）；舊商品保留原值
      row.getCell(4).value = p.isNew && p.jpy != null ? { formula: `ROUND(C${r}*${rate},0)`, result: Math.round(p.jpy * rate) } : (p.cost ?? null);
      row.getCell(5).value = p.price ?? null;
      row.getCell(6).value = p.stock ?? null;
      row.getCell(7).value = p.jpName || null;
      row.getCell(8).value = { formula: `IF(B${r}="","",SUMIFS(${oRange('E')},${oRange('A')},A${r},${oRange('B')},B${r},${oRange('K')},"未出貨"))` };
      row.getCell(3).numFmt = '¥#,##0';
      row.getCell(4).numFmt = '#,##0';
      row.getCell(5).numFmt = '#,##0';
      if (p.isNew) for (let c = 1; c <= 8; c++) row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.newRow } };
    });
    prod.autoFilter = { from: 'A1', to: `H${products.length + 1}` };

    // ----- Orders -----
    ord.columns = [
      { width: 14 }, { width: 22 }, { width: 9 }, { width: 18 }, { width: 7 }, { width: 10 },
      { width: 14 }, { width: 10 }, { width: 11 }, { width: 10 }, { width: 10 },
    ];
    ord.addRow(ORDER_HEADERS);
    styleHeader(ord.getRow(1));
    const priceFormula = (r) => `IF(AND(A${r}<>"",B${r}<>""),SUMIFS(${pRange('E')},${pRange('A')},A${r},${pRange('B')},B${r}),"")`;
    const subtotalFormula = (r) => `IF(AND(C${r}<>"",E${r}<>""),C${r}*E${r},"")`;
    for (let i = 0; i < ORDER_LAST - 1; i++) {
      const r = i + 2;
      const o = orders[i];
      const row = ord.getRow(r);
      if (!o) {
        row.getCell(3).value = { formula: priceFormula(r) };
        row.getCell(6).value = { formula: subtotalFormula(r) };
        continue;
      }
      const k = norm(o.batch) + '|' + norm(o.item);
      const lookup = priceKey.has(k) && !dupKey.has(k) ? priceKey.get(k) : null;
      const price = o.price ?? null;
      row.getCell(1).value = o.batch || null;
      row.getCell(2).value = o.item || null;
      // 售價：如果跟 Product 查到的一樣就用查價公式（改商品售價會自動連動），不一樣就保留原本的數字
      const priceIsNum = typeof price === 'number';
      if (lookup !== null && (price === null || (priceIsNum && Math.abs(price - lookup) < 0.01))) {
        row.getCell(3).value = { formula: priceFormula(r), result: lookup };
      } else {
        row.getCell(3).value = price;
      }
      row.getCell(4).value = o.customer || null;
      row.getCell(5).value = o.qty ?? null;
      const effPrice = lookup !== null && (price === null || (priceIsNum && Math.abs(price - lookup) < 0.01)) ? lookup : price;
      row.getCell(6).value = { formula: subtotalFormula(r), result: typeof effPrice === 'number' && typeof o.qty === 'number' ? effPrice * o.qty : '' };
      row.getCell(7).value = o.shipping || null;
      row.getCell(8).value = o.payMethod || null;
      row.getCell(9).value = o.payStatus || null;
      row.getCell(10).value = o.source || null;
      row.getCell(11).value = o.status || null;
      if (o.isNew) for (let c = 1; c <= 11; c++) row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.newRow } };
    }
    ord.getColumn(3).numFmt = '#,##0';
    ord.getColumn(6).numFmt = '#,##0';
    ord.autoFilter = { from: 'A1', to: `K${orders.length + 1}` };

    // 下拉選單（批次、品項清單放在 Dashboard 右側隱藏欄）
    const listCol = 'Z';
    batches.forEach((b, i) => { dash.getCell(`${listCol}${i + 1}`).value = b; });
    dash.getColumn(listCol).hidden = true;
    const batchListRef = `Dashboard!$${listCol}$1:$${listCol}$${Math.max(batches.length, 1)}`;
    const dv = (list) => ({ type: 'list', allowBlank: true, formulae: [`"${list.join(',')}"`], showErrorMessage: false });
    const oEnd = ORDER_LAST;
    ord.dataValidations.add(`A2:A${oEnd}`, { type: 'list', allowBlank: true, formulae: [batchListRef], showErrorMessage: false });
    ord.dataValidations.add(`G2:G${oEnd}`, dv(OPTIONS.shipping));
    ord.dataValidations.add(`H2:H${oEnd}`, dv(OPTIONS.payMethod));
    ord.dataValidations.add(`I2:I${oEnd}`, dv(OPTIONS.payStatus));
    ord.dataValidations.add(`J2:J${oEnd}`, dv(OPTIONS.source));
    ord.dataValidations.add(`K2:K${oEnd}`, dv(OPTIONS.status));
    ord.addConditionalFormatting({
      ref: `K2:K${oEnd}`,
      rules: [{ type: 'cellIs', operator: 'equal', formulae: ['"未出貨"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COLORS.pending } } } }],
    });
    ord.addConditionalFormatting({
      ref: `I2:I${oEnd}`,
      rules: [
        { type: 'cellIs', operator: 'equal', formulae: ['"未匯款"'], style: { font: { color: { argb: COLORS.unpaid }, bold: true } } },
        { type: 'cellIs', operator: 'equal', formulae: ['"未交證明"'], style: { font: { color: { argb: COLORS.unpaid }, bold: true } } },
      ],
    });

    // ----- Dashboard -----
    buildDashboard(dash, { batches, products, orders, oRange, pRange, batchListRef, opts });

    // ----- 其他分頁：照原樣保留數值 -----
    for (const ex of data.extras || []) {
      const ws = wb.addWorksheet(safeSheetName(ex.name, used));
      ex.rows.forEach((r) => ws.addRow((r || []).map((v) => (v === undefined ? null : v))));
      for (const [addr, z] of Object.entries(ex.formats || {})) ws.getCell(addr).numFmt = z;
      if (ex.rows.length) {
        ws.getRow(1).font = { bold: true };
        const cols = Math.max(...ex.rows.map((r) => (r || []).length));
        for (let c = 1; c <= cols; c++) ws.getColumn(c).width = 14;
      }
    }
    return wb;
  }

  function buildDashboard(ws, ctx) {
    const { batches, orders, products, oRange, batchListRef, opts } = ctx;
    ws.columns = [
      { width: 2 }, { width: 26 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 },
    ];
    const title = ws.getCell('B2');
    title.value = '📦 代購訂單管理 Dashboard';
    title.font = { size: 18, bold: true, color: { argb: COLORS.title } };
    ws.getCell('B3').value = `產生時間：${opts.generatedAt || ''}　｜　淡黃底色 = 本次從截圖新增的資料`;
    ws.getCell('B3').font = { size: 10, color: { argb: COLORS.muted } };

    const section = (cell, text) => {
      const c = ws.getCell(cell);
      c.value = text;
      c.font = { size: 13, bold: true, color: { argb: COLORS.title } };
    };

    // KPI 卡片
    section('B5', '整體概況（所有批次）');
    const kpis = [
      ['未出貨金額', `SUMIFS(${oRange('F')},${oRange('K')},"未出貨")`, '#,##0'],
      ['未出貨件數', `SUMIFS(${oRange('E')},${oRange('K')},"未出貨")`, '#,##0'],
      ['待收款金額（未出貨中）', `SUMIFS(${oRange('F')},${oRange('K')},"未出貨",${oRange('I')},"未匯款")+SUMIFS(${oRange('F')},${oRange('K')},"未出貨",${oRange('I')},"未交證明")`, '#,##0'],
      ['已出貨累計金額', `SUMIFS(${oRange('F')},${oRange('K')},"已出貨")`, '#,##0'],
    ];
    kpis.forEach(([label, f, fmt], i) => {
      const lc = ws.getCell(6, 2 + i);
      const vc = ws.getCell(7, 2 + i);
      lc.value = label;
      lc.font = { size: 10, color: { argb: COLORS.muted } };
      vc.value = { formula: f };
      vc.numFmt = fmt;
      vc.font = { size: 16, bold: true, color: { argb: COLORS.title } };
      for (const c of [lc, vc]) {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.kpiFill } };
        c.border = box;
        c.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    });
    ws.getRow(7).height = 30;

    // 批次查詢（可在 C10 下拉切換）
    section('B9', '🔍 批次查詢（在 C10 下拉選批次）');
    const latest = (opts.newBatches && opts.newBatches[0]) || batches[batches.length - 1] || '';
    const sel = ws.getCell('C10');
    ws.getCell('B10').value = '選擇批次：';
    sel.value = latest;
    sel.dataValidation = { type: 'list', allowBlank: true, formulae: [batchListRef] };
    sel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.newRow } };
    sel.border = box;
    sel.font = { bold: true };
    const q = [
      ['預期總銷售額（已交證明）：', `IF($C$10<>"",SUMIFS(${oRange('F')},${oRange('A')},$C$10,${oRange('K')},"未出貨",${oRange('I')},"已交證明"),0)`],
      ['預期總銷售額：', `IF($C$10<>"",SUMIFS(${oRange('F')},${oRange('A')},$C$10,${oRange('K')},"未出貨"),0)`],
      ['需進貨總件數（未出貨）：', `IF($C$10<>"",SUMIFS(${oRange('E')},${oRange('A')},$C$10,${oRange('K')},"未出貨"),0)`],
      ['訂單筆數（不含取消）：', `IF($C$10<>"",COUNTIFS(${oRange('A')},$C$10)-COUNTIFS(${oRange('A')},$C$10,${oRange('K')},"取消"),0)`],
    ];
    q.forEach(([label, f], i) => {
      ws.getCell(11 + i, 2).value = label;
      const c = ws.getCell(11 + i, 3);
      c.value = { formula: f };
      c.numFmt = '#,##0';
      c.font = { bold: true };
    });

    // 客人查詢
    section('E9', '🔍 客人查詢（在 F10 輸入名字）');
    ws.getCell('E10').value = '客人姓名：';
    const cust = ws.getCell('F10');
    const lastNew = orders.filter((o) => o.isNew).slice(-1)[0];
    cust.value = lastNew ? lastNew.customer : '';
    cust.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.newRow } };
    cust.border = box;
    cust.font = { bold: true };
    const cq = [
      ['未出貨件數：', `IF($F$10<>"",SUMIFS(${oRange('E')},${oRange('D')},$F$10,${oRange('K')},"未出貨"),0)`],
      ['未出貨金額：', `IF($F$10<>"",SUMIFS(${oRange('F')},${oRange('D')},$F$10,${oRange('K')},"未出貨"),0)`],
      ['其中待收款：', `IF($F$10<>"",SUMIFS(${oRange('F')},${oRange('D')},$F$10,${oRange('K')},"未出貨",${oRange('I')},"未匯款")+SUMIFS(${oRange('F')},${oRange('D')},$F$10,${oRange('K')},"未出貨",${oRange('I')},"未交證明"),0)`],
    ];
    cq.forEach(([label, f], i) => {
      ws.getCell(11 + i, 5).value = label;
      const c = ws.getCell(11 + i, 6);
      c.value = { formula: f };
      c.numFmt = '#,##0';
      c.font = { bold: true };
    });

    let r = 17;
    // 本次新增批次的品項進貨統計
    const newBatches = opts.newBatches || [];
    for (const b of newBatches) {
      section(`B${r}`, `🛒 ${b} 品項進貨統計（未出貨）`);
      r++;
      const head = ws.getRow(r);
      ['品項', '日幣單價', '台幣售價', '需進貨量', '日幣小計', '台幣營收'].forEach((h, i) => { head.getCell(2 + i).value = h; });
      styleHeader(head);
      head.getCell(1).fill = undefined;
      r++;
      const start = r;
      for (const p of products.filter((x) => x.batch === b)) {
        const row = ws.getRow(r);
        row.getCell(2).value = p.name;
        row.getCell(3).value = p.jpy ?? null;
        row.getCell(4).value = p.price ?? null;
        row.getCell(5).value = { formula: `SUMIFS(${oRange('E')},${oRange('A')},"${b.replace(/"/g, '""')}",${oRange('B')},B${r},${oRange('K')},"未出貨")` };
        row.getCell(6).value = { formula: `C${r}*E${r}` };
        row.getCell(7).value = { formula: `D${r}*E${r}` };
        row.getCell(3).numFmt = '¥#,##0';
        row.getCell(6).numFmt = '¥#,##0';
        for (const c of [4, 7]) row.getCell(c).numFmt = '#,##0';
        for (let c = 2; c <= 7; c++) row.getCell(c).border = box;
        r++;
      }
      const tot = ws.getRow(r);
      tot.getCell(2).value = '合計';
      for (const c of ['E', 'F', 'G']) {
        const cell = tot.getCell(c);
        cell.value = { formula: `SUM(${c}${start}:${c}${r - 1})` };
        cell.numFmt = c === 'F' ? '¥#,##0' : '#,##0';
      }
      for (let c = 2; c <= 7; c++) {
        tot.getCell(c).font = { bold: true };
        tot.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.band } };
        tot.getCell(c).border = box;
      }
      r += 2;
    }

    // 批次總覽
    section(`B${r}`, '📊 批次總覽');
    r++;
    const head = ws.getRow(r);
    ['批次', '訂單筆數', '未出貨件數', '未出貨金額', '待收款金額', '已出貨金額', '取消筆數'].forEach((h, i) => { head.getCell(2 + i).value = h; });
    styleHeader(head);
    head.getCell(1).fill = undefined;
    r++;
    // 排序：本次新批次 → 日期開頭（如 0821恐龍）由新到舊 → 其他
    const dated = batches.filter((b) => /^\d{4}/.test(b) && !newBatches.includes(b)).sort((a, b) => b.slice(0, 4).localeCompare(a.slice(0, 4)));
    const others = batches.filter((b) => !/^\d{4}/.test(b) && !newBatches.includes(b));
    const ordered = [...newBatches.filter((b) => batches.includes(b)), ...dated, ...others];
    for (const b of ordered) {
      const row = ws.getRow(r);
      const ref = `$B${r}`;
      row.getCell(2).value = b;
      row.getCell(3).value = { formula: `COUNTIFS(${oRange('A')},${ref})` };
      row.getCell(4).value = { formula: `SUMIFS(${oRange('E')},${oRange('A')},${ref},${oRange('K')},"未出貨")` };
      row.getCell(5).value = { formula: `SUMIFS(${oRange('F')},${oRange('A')},${ref},${oRange('K')},"未出貨")` };
      row.getCell(6).value = { formula: `SUMIFS(${oRange('F')},${oRange('A')},${ref},${oRange('K')},"未出貨",${oRange('I')},"未匯款")+SUMIFS(${oRange('F')},${oRange('A')},${ref},${oRange('K')},"未出貨",${oRange('I')},"未交證明")` };
      row.getCell(7).value = { formula: `SUMIFS(${oRange('F')},${oRange('A')},${ref},${oRange('K')},"已出貨")` };
      row.getCell(8).value = { formula: `COUNTIFS(${oRange('A')},${ref},${oRange('K')},"取消")` };
      for (let c = 3; c <= 8; c++) row.getCell(c).numFmt = '#,##0';
      for (let c = 2; c <= 8; c++) {
        row.getCell(c).border = box;
        if (newBatches.includes(b)) row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.newRow } };
      }
      r++;
    }
  }

  // 檔名：原檔名去掉「_最新」與舊的日期版本號，再加上 _YYYYMMDD_V#
  function outputFileName(inputName, today) {
    const base0 = (inputName || '訂單管理').replace(/\.xlsx?$/i, '');
    const m = base0.match(/^(.*)_(\d{8})_V(\d+)$/i);
    let base = m ? m[1] : base0.replace(/_最新$/, '');
    let v = 1;
    if (m && m[2] === today) v = Number(m[3]) + 1;
    return `${base}_${today}_V${v}.xlsx`;
  }

  return { readWorkbook, parseOcr, buildMatches, buildWorkbook, batchList, matchProduct, outputFileName, norm, OPTIONS };
});
