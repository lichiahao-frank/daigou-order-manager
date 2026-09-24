/* 代購訂單管理：網頁互動 */
(function () {
  'use strict';
  const Core = window.OrderCore;
  const $ = (id) => document.getElementById(id);
  const MAX_SIDE = 2400; // 圖片長邊上限，太大會先縮小再送出
  const CONCURRENCY = 1; // 同一個辨識引擎一次處理一張

  // ---------- 簡易密碼門檻 ----------
  // 注意：靜態網站只能在瀏覽器端檢查，只能擋一般使用者，擋不住懂技術的人
  const GATE_HASH = '812de6e718f869feb16b45c6bbcfdb1269fe6f6fffdc2420166482e3cd0aa647';
  const GATE_KEY = 'order-app-gate';
  async function sha256(text) {
    const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
    return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');
  }
  function openApp() {
    $('gate').classList.add('hidden');
    $('appRoot').classList.remove('hidden');
  }
  try { if (localStorage.getItem(GATE_KEY) === GATE_HASH) openApp(); } catch { /* 無法記住就每次輸入 */ }
  $('gateForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const hash = await sha256($('gateInput').value.trim());
    if (hash !== GATE_HASH) {
      $('gateMsg').textContent = '密碼錯誤';
      $('gateInput').select();
      return;
    }
    try { localStorage.setItem(GATE_KEY, hash); } catch { /* 忽略 */ }
    openApp();
  });

  const state = {
    existing: null, // readWorkbook 結果
    fileName: '',
    shots: [], // {id, file, url, status, error, result}
    newProducts: [],
    newOrders: [],
  };

  // ---------- OCR 引擎（背景執行緒） ----------
  let worker = null;
  let msgId = 0;
  const pending = new Map();
  function getWorker() {
    if (!worker) {
      worker = new Worker('vendor/ocr-worker.js', { type: 'module' });
      worker.onmessage = (e) => {
        const p = pending.get(e.data.id);
        if (!p) return;
        pending.delete(e.data.id);
        e.data.ok ? p.resolve(e.data) : p.reject(new Error(e.data.error));
      };
      worker.onerror = (e) => {
        for (const p of pending.values()) p.reject(new Error('辨識引擎載入失敗：' + (e.message || '請檢查網路後重新整理')));
        pending.clear();
        worker = null;
        engineReady = null;
      };
    }
    return worker;
  }
  function callWorker(msg, transfer) {
    return new Promise((resolve, reject) => {
      const id = ++msgId;
      pending.set(id, { resolve, reject });
      getWorker().postMessage({ id, ...msg }, transfer || []);
    });
  }
  let engineReady = null;
  function ensureEngine() {
    if (!engineReady) {
      setEngine('辨識模型載入中（第一次約需下載 45～60 MB）…');
      engineReady = callWorker({ type: 'init' })
        .then(() => setEngine('✅ 辨識模型已就緒'))
        .catch((e) => { engineReady = null; setEngine('⚠ 模型載入失敗'); throw e; });
    }
    return engineReady;
  }
  function setEngine(t) { $('engineStatus').textContent = t; }

  // ---------- 通用：拖放區 ----------
  function bindDrop(zone, input, onFiles) {
    zone.onclick = () => input.click();
    input.onchange = () => { onFiles([...input.files]); input.value = ''; };
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('over'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('over'));
    zone.addEventListener('drop', (e) => {
      e.preventDefault();
      zone.classList.remove('over');
      onFiles([...e.dataTransfer.files]);
    });
  }
  const fmt = (n) => (n == null || n === '' || Number.isNaN(n) ? '' : Number(n).toLocaleString('zh-TW'));
  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  // ---------- 步驟 1：Excel ----------
  bindDrop($('dropExcel'), $('fileExcel'), async (files) => {
    const f = files.find((x) => /\.xlsx$/i.test(x.name));
    if (!f) return showExcelMsg('請選 .xlsx 檔', 'bad');
    try {
      const buf = await f.arrayBuffer();
      state.existing = Core.readWorkbook(XLSX, buf);
      state.fileName = f.name;
      if (!state.existing.hasOrders && !state.existing.hasProduct) {
        showExcelMsg('這個檔案裡找不到 Product 或 Orders 分頁，請確認是登記檔', 'bad');
      }
      renderExcelInfo();
      refreshMatches();
    } catch (e) {
      showExcelMsg('讀取 Excel 失敗：' + e.message, 'bad');
    }
  });
  $('btnFresh').onclick = () => {
    state.existing = { products: [], orders: [], extras: [], hasProduct: false, hasOrders: false };
    state.fileName = '訂單管理.xlsx';
    renderExcelInfo();
    refreshMatches();
  };
  function showExcelMsg(text, kind) {
    $('excelInfo').innerHTML = `<div class="msg ${kind}">${esc(text)}</div>`;
  }
  function renderExcelInfo() {
    const ex = state.existing;
    const batches = Core.batchList(ex.products, ex.orders);
    $('excelInfo').innerHTML = `<div class="stats">
      <span class="pill">📄 ${esc(state.fileName)}</span>
      <span class="pill">商品 ${fmt(ex.products.length)} 筆</span>
      <span class="pill">訂單 ${fmt(ex.orders.length)} 筆</span>
      <span class="pill">批次 ${batches.length} 個</span>
      ${ex.extras.length ? `<span class="pill">其他分頁照原樣保留：${ex.extras.map((e) => esc(e.name)).join('、')}</span>` : ''}
    </div>`;
    $('batchOptions').innerHTML = [...batches].reverse().map((b) => `<option value="${esc(b)}">`).join('');
    updateButtons();
  }

  // ---------- 步驟 2：截圖 ----------
  let seq = 0;
  function addShots(files) {
    for (const file of files) {
      if (!file.type.startsWith('image/')) continue;
      state.shots.push({ id: ++seq, file, url: URL.createObjectURL(file), status: 'wait', error: '', result: null });
    }
    renderThumbs();
    if (state.shots.length) ensureEngine().catch(() => {}); // 提早開始下載模型
  }
  bindDrop($('dropShots'), $('fileShots'), addShots);
  document.addEventListener('paste', (e) => {
    const files = [...(e.clipboardData?.files || [])].filter((f) => f.type.startsWith('image/'));
    if (files.length) addShots(files);
  });
  const STATUS = { wait: ['等待辨識', 'wait'], busy: ['辨識中…', 'busy'], ok: ['完成', 'ok'], bad: ['失敗', 'bad'] };
  function renderThumbs() {
    $('thumbs').innerHTML = state.shots.map((s) => {
      const [label, cls] = STATUS[s.status];
      const kind = s.result ? { catalog: '型錄', orders: '留言', mixed: '型錄＋留言', other: '無關圖片' }[s.result.kind] || '' : '';
      const detail = s.result ? `${s.result.products.length} 商品／${s.result.orders.length} 留言` : '';
      return `<div class="thumb">
        <img src="${s.url}" alt="${esc(s.file.name)}">
        <button class="x" data-del="${s.id}" title="移除">×</button>
        <div class="meta"><span class="tag ${cls}">${label}${kind ? '・' + kind : ''}</span>
        ${s.status === 'bad' ? `<button class="link" data-retry="${s.id}">重試</button>` : `<span>${detail}</span>`}</div>
        ${s.error ? `<div class="meta warn">${esc(s.error)}</div>` : ''}
        ${s.result && s.result.notes ? `<div class="meta raw">${esc(s.result.notes)}</div>` : ''}
        ${s.result && s.result.text ? `<details class="ocrtext meta"><summary>看辨識出的文字</summary><pre>${esc(s.result.text)}</pre></details>` : ''}
      </div>`;
    }).join('');
    updateButtons();
  }
  $('thumbs').addEventListener('click', (e) => {
    const del = e.target.dataset.del, retry = e.target.dataset.retry;
    if (del) {
      state.shots = state.shots.filter((s) => String(s.id) !== del);
      renderThumbs();
      refreshMatches();
    }
    if (retry) {
      const s = state.shots.find((x) => String(x.id) === retry);
      s.status = 'wait'; s.error = '';
      runParse();
    }
  });

  function updateButtons() {
    const pending = state.shots.some((s) => s.status === 'wait');
    const busy = state.shots.some((s) => s.status === 'busy');
    $('btnParse').disabled = !state.existing || !pending || busy;
    $('btnParse').textContent = busy ? '辨識中…' : '開始辨識';
    const why = !state.existing ? '請先完成步驟 1' : !state.shots.length ? '' : !pending && !busy ? '全部截圖都處理完了' : '';
    if (!busy) $('parseStatus').textContent = why;
  }

  // 太大的圖先縮小，OCR 比較快
  async function toImageBuffer(file) {
    const bmp = await createImageBitmap(file);
    const scale = Math.min(1, MAX_SIDE / Math.max(bmp.width, bmp.height));
    if (scale === 1) return file.arrayBuffer();
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(bmp.width * scale);
    canvas.height = Math.round(bmp.height * scale);
    canvas.getContext('2d').drawImage(bmp, 0, 0, canvas.width, canvas.height);
    const blob = await new Promise((r) => canvas.toBlob(r, 'image/png'));
    return blob.arrayBuffer();
  }

  async function parseOne(shot) {
    shot.status = 'busy';
    renderThumbs();
    try {
      await ensureEngine();
      const buffer = await toImageBuffer(shot.file);
      const res = await callWorker({ type: 'recognize', buffer }, [buffer]);
      shot.result = Core.parseOcr(res.lines);
      shot.status = 'ok';
    } catch (e) {
      shot.status = 'bad';
      shot.error = e.message;
    }
    renderThumbs();
  }

  async function runParse() {
    if (!$('batch').value.trim()) {
      $('parseStatus').textContent = '請先填批次名稱';
      $('batch').focus();
      return;
    }
    const queue = state.shots.filter((s) => s.status === 'wait');
    $('parseStatus').textContent = '辨識中…';
    const workers = Array.from({ length: CONCURRENCY }, async () => {
      while (queue.length) await parseOne(queue.shift());
    });
    await Promise.all(workers);
    const failed = state.shots.filter((s) => s.status === 'bad').length;
    $('parseStatus').textContent = failed ? `有 ${failed} 張失敗，可以按「重試」` : '辨識完成，請到下方核對';
    refreshMatches();
  }
  $('btnParse').onclick = runParse;
  $('batch').addEventListener('change', refreshMatches);
  $('rate').addEventListener('input', renderProducts);

  // ---------- 步驟 3：核對 ----------
  function refreshMatches() {
    const done = state.shots.filter((s) => s.status === 'ok' && s.result);
    if (!state.existing || !done.length) {
      if (!state.newProducts.length && !state.newOrders.length) {
        $('cardReview').classList.add('hidden');
        $('cardExport').classList.add('hidden');
      }
      return;
    }
    const batch = $('batch').value.trim();
    const results = done.map((s) => ({ fileName: s.file.name, ...s.result }));
    const m = Core.buildMatches(results, state.existing, batch);
    // 保留使用者手動新增的列
    state.newProducts = [...m.newProducts, ...state.newProducts.filter((p) => p.manual)];
    state.newOrders = [...m.newOrders, ...state.newOrders.filter((o) => o.manual)];
    for (const o of state.newOrders) o.batch = batch;
    for (const p of state.newProducts) p.batch = batch;
    $('cardReview').classList.remove('hidden');
    $('cardExport').classList.remove('hidden');
    rematchOrders();
    renderProducts();
    renderOrders();
  }

  function batchProducts() {
    const batch = $('batch').value.trim();
    const list = (state.existing?.products || []).filter((p) => p.batch === batch);
    return [...list, ...state.newProducts.filter((p) => p.include && p.name)];
  }
  // 商品表改過後，重新配對「沒被手動改過品項」的訂單
  function rematchOrders() {
    const cands = batchProducts().filter((p) => !p.guessed);
    for (const o of state.newOrders) {
      if (o.itemEdited || o.manual) { o.price = priceFor(o.item); continue; }
      const m = Core.matchProduct(o.rawItem || o.item, cands);
      o.item = m ? m.product.name : (o.rawItem || o.item);
      o.price = m ? m.product.price : null;
      o.warn = (o.warn || '').split('；').filter((w) => w && !/找不到|自動對到|商品表還沒有/.test(w))
        .concat(!m ? [`商品表還沒有「${o.item}」`] : (m.how === 'fuzzy' && o.rawItem !== o.item ? [`「${o.rawItem}」自動對到「${o.item}」，請確認`] : []))
        .join('；');
    }
  }
  // 留言中出現、但這個批次商品表沒有的品項
  function missingItems() {
    const have = new Set(batchProducts().map((p) => Core.norm(p.name)));
    const out = [];
    for (const o of state.newOrders) {
      if (!o.include || !o.item || have.has(Core.norm(o.item)) || out.includes(o.item)) continue;
      out.push(o.item);
    }
    return out;
  }

  function priceFor(item) {
    const hit = batchProducts().filter((p) => Core.norm(p.name) === Core.norm(item));
    return hit.length ? hit[hit.length - 1].price : null;
  }

  function renderProducts() {
    const rate = Number($('rate').value) || 0.2;
    $('tbProducts').innerHTML = state.newProducts.map((p, i) => `
      <tr class="${p.include ? '' : 'off'}">
        <td><input type="checkbox" data-p="${i}" data-k="include" ${p.include ? 'checked' : ''}></td>
        <td class="w-l"><input value="${esc(p.name)}" data-p="${i}" data-k="name">${p.warn ? `<div class="warn">⚠ ${esc(p.warn)}</div>` : ''}</td>
        <td class="w-s"><input type="number" value="${p.jpy ?? ''}" data-p="${i}" data-k="jpy"></td>
        <td class="w-s"><input type="number" value="${p.price ?? ''}" data-p="${i}" data-k="price"></td>
        <td class="num">${p.jpy ? fmt(Math.round(p.jpy * rate)) : ''}</td>
        <td><input value="${esc(p.jpName)}" data-p="${i}" data-k="jpName"></td>
      </tr>`).join('') || '<tr><td colspan="6" class="hint" style="padding:12px">這次沒有新商品（訂單會對到此批次已有的商品）</td></tr>';
    $('itemOptions').innerHTML = batchProducts().map((p) => `<option value="${esc(p.name)}">`).join('');
    const miss = missingItems();
    $('missingItems').innerHTML = miss.length
      ? `<div class="msg info">留言裡有這些品項，但商品表還沒有：${miss.map((m) => `<button data-add-item="${esc(m)}">＋ ${esc(m)}</button>`).join(' ')}　<span class="hint">點一下加入商品表，再補上售價；或把上面猜錯名稱的商品改成正確名稱。</span></div>`
      : '';
    updateSummary();
  }

  const sel = (opts, v, attrs) => `<select ${attrs}>${['', ...opts].map((o) => `<option ${o === v ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
  function renderOrders() {
    const O = Core.OPTIONS;
    $('tbOrders').innerHTML = state.newOrders.map((o, i) => {
      const a = (k) => `data-o="${i}" data-k="${k}"`;
      const sub = o.price != null && o.qty ? o.price * o.qty : null;
      return `<tr class="${o.include ? '' : 'off'}">
        <td><input type="checkbox" ${a('include')} ${o.include ? 'checked' : ''}></td>
        <td class="w-m"><input value="${esc(o.customer)}" ${a('customer')}></td>
        <td class="w-l"><input value="${esc(o.item)}" list="itemOptions" ${a('item')}>
          ${o.warn ? `<div class="warn">⚠ ${esc(o.warn)}</div>` : ''}
          ${o.rawText ? `<div class="raw">原文：${esc(o.rawText)}</div>` : ''}</td>
        <td class="w-xs"><input type="number" min="1" value="${o.qty ?? ''}" ${a('qty')}></td>
        <td class="num">${o.price == null ? '<span class="warn">查無售價</span>' : fmt(o.price)}</td>
        <td class="num">${fmt(sub)}</td>
        <td class="w-s">${sel(O.shipping, o.shipping, a('shipping'))}</td>
        <td class="w-s">${sel(O.payStatus, o.payStatus, a('payStatus'))}</td>
        <td class="w-s">${sel(O.source, o.source, a('source'))}</td>
        <td class="w-s">${sel(O.status, o.status, a('status'))}</td>
      </tr>`;
    }).join('') || '<tr><td colspan="10" class="hint" style="padding:12px">這次沒有辨識到訂單</td></tr>';
    updateSummary();
  }

  function onEdit(e) {
    const t = e.target;
    const k = t.dataset.k;
    if (!k) return;
    const isP = t.dataset.p !== undefined;
    const row = isP ? state.newProducts[t.dataset.p] : state.newOrders[t.dataset.o];
    let v = t.type === 'checkbox' ? t.checked : t.value;
    if (['jpy', 'price', 'qty'].includes(k)) v = v === '' ? null : Number(v);
    row[k] = v;
    if (isP) {
      // 商品改了 → 名稱視為已確認，訂單重新配對、售價重算
      if (k === 'name') { row.guessed = false; row.warn = (row.warn || '').split('；').filter((w) => !/手寫|請填中文/.test(w)).join('；'); }
      if (k === 'jpy' && v) row.warn = (row.warn || '').split('；').filter((w) => !/日幣/.test(w)).join('；');
      if (k === 'price' && v) row.warn = (row.warn || '').split('；').filter((w) => !/台幣售價/.test(w)).join('；');
      rematchOrders();
      if (e.type === 'change') { renderProducts(); renderOrders(); } else updateSummary();
    } else {
      if (k === 'item') { row.itemEdited = true; row.price = priceFor(v); }
      if (e.type === 'change' || k === 'include') renderOrders();
      else updateSummary();
    }
  }
  $('cardReview').addEventListener('change', onEdit);

  $('missingItems').addEventListener('click', (e) => {
    const name = e.target.dataset.addItem;
    if (!name) return;
    state.newProducts.push({ include: true, manual: true, batch: $('batch').value.trim(), name, jpName: '', jpy: null, price: null, isNew: true, warn: '請填台幣售價' });
    rematchOrders();
    renderProducts();
    renderOrders();
  });
  $('addProduct').onclick = () => {
    state.newProducts.push({ include: true, manual: true, batch: $('batch').value.trim(), name: '', jpName: '', jpy: null, price: null, isNew: true, warn: '' });
    renderProducts();
  };
  $('addOrder').onclick = () => {
    state.newOrders.push({ include: true, manual: true, batch: $('batch').value.trim(), item: '', customer: '', qty: 1, price: null, shipping: '分擔', payMethod: '', payStatus: '未匯款', source: '社群', status: '未出貨', isNew: true, warn: '' });
    renderOrders();
  };

  // ---------- 步驟 4：匯出 ----------
  function picked() {
    return {
      products: state.newProducts.filter((p) => p.include && p.name),
      orders: state.newOrders.filter((o) => o.include && o.item),
    };
  }
  function updateSummary() {
    const { products, orders } = picked();
    const total = orders.reduce((s, o) => s + (o.price || 0) * (o.qty || 0), 0);
    const missing = orders.filter((o) => o.price == null).length;
    $('exportSummary').innerHTML = `將新增 <b>${products.length}</b> 個商品、<b>${orders.length}</b> 筆訂單，合計 <b>$${fmt(total)}</b>` +
      (missing ? `　<span class="warn">⚠ 有 ${missing} 筆訂單查無售價（品項名稱對不到商品）</span>` : '');
  }

  function today() {
    const d = new Date();
    return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
  }

  $('btnExport').onclick = async () => {
    const batch = $('batch').value.trim();
    if (!batch) return ($('exportMsg').innerHTML = '<div class="msg bad">請先填批次名稱</div>');
    const { products, orders } = picked();
    const btn = $('btnExport');
    btn.disabled = true;
    btn.textContent = '產生中…';
    try {
      const ex = state.existing;
      const data = {
        products: [...ex.products, ...products.map((p) => ({ ...p, batch, cost: null, stock: null }))],
        orders: [...ex.orders, ...orders.map((o) => ({ ...o, batch }))],
        extras: ex.extras,
      };
      const now = new Date();
      const wb = Core.buildWorkbook(ExcelJS, data, {
        rate: Number($('rate').value) || 0.2,
        newBatches: [batch],
        generatedAt: now.toLocaleString('zh-TW', { hour12: false }),
      });
      const buf = await wb.xlsx.writeBuffer();
      const name = Core.outputFileName(state.fileName, today());
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      a.download = name;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 5000);
      $('exportMsg').innerHTML = `<div class="msg ok">已下載 <b>${esc(name)}</b>。下次更新時，請上傳這個新檔案。</div>`;
    } catch (e) {
      $('exportMsg').innerHTML = `<div class="msg bad">產生 Excel 失敗：${esc(e.message)}</div>`;
    } finally {
      btn.disabled = false;
      btn.textContent = '下載 Excel';
    }
  };
})();
