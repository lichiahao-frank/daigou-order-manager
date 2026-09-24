// 留言解析規則測試：模擬 OCR 輸出各種寫法
const C = require('../public/core.js');
const L = (...rows) => rows.map((r, i) => (Array.isArray(r) ? r : [r]).map((t) => ({ text: t, box: { x: 10, y: i * 40, width: 100, height: 30 }, confidence: 1 })));
const cases = [
  ['基本格式（LINE 實際截圖）', L('留言13', '客人A', '幽靈兔+2', '總金額：800', '運費：分擔', '6天前的晚上8:03', ['oo', '客人B'], '商品：幽靈兔+1', '總金額：400', '境內運：分擔', '6天前的晚上8:03'),
    [['客人A', '幽靈兔', 2, 800, '分擔'], ['客人B', '幽靈兔', 1, 400, '分擔']]],
  ['x 與全形＋、一行多品項', L('小明', '幽靈兔x2、幽靈吉＋1', '總計 $1,200', '不分擔', '昨天 下午3:20'),
    [['小明', '幽靈兔', 2, 1200, '不分擔'], ['小明', '幽靈吉', 1, 1200, '不分擔']]],
  ['「個」當單位', L('阿花', '南瓜S娃 1個', '合計630', '自吃', '2小時前'),
    [['阿花', '南瓜S娃', 1, 630, '自吃']]],
  ['賣家閒聊不算訂單', L('賣家', '今天開放下單喔', '3分鐘前', '路人', '好可愛', '1分鐘前'), []],
  ['截圖從留言中間開始（沒有名字）', L('總金額：400', '境內運：分擔', '6天前的晚上8:04', '客人C', '幽靈八+1', '6天前的晚上8:05'),
    [['客人C', '幽靈八', 1, 0, '']]],
];
let fail = 0;
for (const [name, lines, expect] of cases) {
  const r = C.parseOcr(lines);
  const got = r.orders.flatMap((o) => o.items.map((it) => [o.customer, it.item, it.qty, o.stated_total, o.shipping]));
  const ok = JSON.stringify(got) === JSON.stringify(expect);
  if (!ok) fail++;
  console.log(ok ? '✅' : '❌', name, ok ? '' : `\n   預期 ${JSON.stringify(expect)}\n   實際 ${JSON.stringify(got)}`);
}
process.exit(fail ? 1 : 0);
