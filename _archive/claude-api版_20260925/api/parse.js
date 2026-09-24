// 截圖辨識 API：收一張截圖，請 Claude 讀出商品型錄或客人留言訂單，回傳 JSON
const Anthropic = require('@anthropic-ai/sdk');

const MODEL = 'claude-opus-5';
const MAX_IMAGE_BYTES = 4 * 1024 * 1024;
const MEDIA_TYPES = ['image/jpeg', 'image/png', 'image/webp', 'image/gif'];

const SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['kind', 'products', 'orders', 'notes'],
  properties: {
    kind: { type: 'string', enum: ['catalog', 'orders', 'mixed', 'other'] },
    products: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['name_zh', 'name_jp', 'jpy', 'price_twd'],
        properties: {
          name_zh: { type: 'string' },
          name_jp: { type: 'string' },
          jpy: { type: 'integer' },
          price_twd: { type: 'integer' },
        },
      },
    },
    orders: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['customer', 'items', 'stated_total', 'shipping', 'raw_text'],
        properties: {
          customer: { type: 'string' },
          items: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              required: ['item', 'qty'],
              properties: { item: { type: 'string' }, qty: { type: 'integer' } },
            },
          },
          stated_total: { type: 'integer' },
          shipping: { type: 'string', enum: ['分擔', '不分擔', '自吃', ''] },
          raw_text: { type: 'string' },
        },
      },
    },
    notes: { type: 'string' },
  },
};

const SYSTEM = `你是日本代購賣家的訂單登記助理，負責把手機截圖轉成結構化資料。截圖有兩種：

1. 商品型錄（kind=catalog）：日本官網商品卡片，有日文品名與日幣含稅價（¥），賣家通常用紅筆手寫「中文簡稱 $台幣售價」，例如「幽靈兔 $400」。
   - name_zh：手寫的中文簡稱，照寫的字抄，不要自行改名；若沒有手寫，用日文品名翻成簡短中文。
   - name_jp：卡片上的完整日文品名。
   - jpy：日幣含稅價（數字，不含逗號）。price_twd：手寫台幣售價；沒寫就填 0。

2. 客人留言（kind=orders）：LINE 社群、FB、IG 等的留言或聊天。每一則留言是一筆訂單。
   - customer：留言者顯示名稱（頭像旁邊的名字），不是留言內容。
   - items：每個品項與數量。「幽靈兔 +2」「幽靈兔x2」代表數量 2；只寫品名沒寫數量就是 1。品名依客人寫的字抄，但如果明顯是下方「已知商品」清單裡某個商品的錯字或簡寫，改用清單上的名稱。
   - stated_total：留言寫的總金額（數字），沒寫填 0。
   - shipping：運費或境內運的分擔方式，只能是「分擔」「不分擔」「自吃」，沒寫填空字串。
   - raw_text：留言原文（換行改成空格）。
   - 忽略賣家自己的留言、系統訊息、與訂單無關的聊天。

同一張圖兩種都有就用 kind=mixed；都不是就用 other。看不清楚的字在 notes 說明，不要亂猜數字。`;

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: '只接受 POST' });

  const password = process.env.APP_PASSWORD;
  if (!password) return res.status(500).json({ error: '伺服器沒有設定 APP_PASSWORD，請先到 Vercel 設定' });
  if (req.headers['x-app-password'] !== password) return res.status(401).json({ error: '密碼錯誤' });
  if (!process.env.ANTHROPIC_API_KEY) return res.status(500).json({ error: '伺服器沒有設定 ANTHROPIC_API_KEY' });

  const { image, mediaType, knownItems } = req.body || {};
  if (typeof image !== 'string' || !image) return res.status(400).json({ error: '沒有收到圖片' });
  if (!MEDIA_TYPES.includes(mediaType)) return res.status(400).json({ error: `不支援的圖片格式：${mediaType}` });
  if (image.length * 0.75 > MAX_IMAGE_BYTES) return res.status(413).json({ error: '圖片太大（上限 4MB）' });

  const known = Array.isArray(knownItems) ? knownItems.filter((s) => typeof s === 'string').slice(0, 300) : [];
  const client = new Anthropic();

  try {
    const response = await client.beta.messages.create({
      model: MODEL,
      max_tokens: 16000,
      betas: ['server-side-fallback-2026-07-01'],
      fallbacks: 'default',
      thinking: { type: 'adaptive' },
      system: SYSTEM,
      output_config: { format: { type: 'json_schema', schema: SCHEMA } },
      messages: [
        {
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: mediaType, data: image } },
            {
              type: 'text',
              text: known.length
                ? `已知商品（這個批次目前的品名，比對留言時優先使用）：${known.join('、')}\n\n請辨識這張截圖。`
                : '請辨識這張截圖。',
            },
          ],
        },
      ],
    });

    if (response.stop_reason === 'refusal') {
      return res.status(422).json({ error: '這張圖被模型拒絕處理，請改成手動輸入' });
    }
    if (response.stop_reason === 'max_tokens') {
      return res.status(422).json({ error: '這張圖內容太多，請裁成兩張再試' });
    }
    const text = response.content.filter((b) => b.type === 'text').map((b) => b.text).join('');
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch {
      return res.status(502).json({ error: '辨識結果格式錯誤，請重試一次' });
    }
    return res.status(200).json({ result: parsed, usage: response.usage });
  } catch (err) {
    if (err instanceof Anthropic.AuthenticationError) return res.status(500).json({ error: 'API 金鑰無效，請檢查 ANTHROPIC_API_KEY' });
    if (err instanceof Anthropic.RateLimitError) return res.status(429).json({ error: '請求太頻繁，請稍等一下再試' });
    if (err instanceof Anthropic.BadRequestError) return res.status(400).json({ error: `API 拒絕這個請求：${err.message}` });
    if (err instanceof Anthropic.APIError) return res.status(502).json({ error: `Claude API 錯誤（${err.status}）：${err.message}` });
    return res.status(500).json({ error: `伺服器錯誤：${err.message}` });
  }
};

