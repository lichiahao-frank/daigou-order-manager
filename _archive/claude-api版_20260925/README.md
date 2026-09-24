# 代購訂單管理（截圖 → Excel）

## 平常怎麼用

1. 打開網站，上傳目前最新的登記 Excel（第一次用 `吉伊卡哇登記_最新.xlsx`）。
2. 填批次名稱（例：`0924樂園`），把截圖拖進去：
   - 商品型錄截圖：要有手寫的「中文簡稱 $台幣售價」
   - 客人留言截圖：LINE 社群／FB／IG 留言都可以
3. 按「開始辨識」，每張大約 10～40 秒。
4. 核對表格：有橘色 ⚠ 的列要看一下，取消勾選的列不會寫入。
5. 按「下載 Excel」，會拿到 `吉伊卡哇登記_YYYYMMDD_V#.xlsx`。**下次要上傳這個新檔**。

## 產出的 Excel

| 分頁 | 內容 |
|---|---|
| Dashboard | 整體 KPI、批次查詢（C10 下拉）、客人查詢（F10 輸入）、本次批次進貨統計、批次總覽 |
| Product | 原本全部商品＋新商品（淡黃底色），多了「日文品名」和「未出貨需求量」欄 |
| Orders | 原本全部訂單＋新訂單（淡黃底色），售價會自動對 Product 查價，下拉選單都保留 |
| 其他 | 車次、賣貨便、ShippingList 照原樣保留數值；History 因為是 Excel 保留字，改名為 `History_出貨歷史` |

## 部署到 Vercel（只需要做一次）

需要兩個環境變數，**不要寫進程式碼**：

- `ANTHROPIC_API_KEY`：Anthropic API 金鑰（到 console.anthropic.com 申請）
- `APP_PASSWORD`：自己設一組網站存取密碼，避免別人用你的 API 額度

```bash
cd "app"
vercel link
vercel env add ANTHROPIC_API_KEY production
vercel env add APP_PASSWORD production
vercel deploy --prod
```

## 本機測試（開發用）

```bash
npm install
npm test          # 用測試資料產生 Excel 到 test/
npm run dev       # 模擬模式（不花 API 費用），密碼 test123，開 http://localhost:5173
```
