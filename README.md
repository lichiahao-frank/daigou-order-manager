# 代購訂單管理（截圖 → Excel，免費版）

**網址：https://lichiahao-frank.github.io/daigou-order-manager/**

所有處理都在使用者自己的瀏覽器裡完成：Excel 和截圖都不會上傳，不需要伺服器、不需要 API 金鑰。

## 怎麼用

1. 打開網頁，上傳目前最新的登記 Excel（沒有舊檔可以按「從空白開始建立新檔」）。
2. 填批次名稱（例：`0924樂園`），把截圖拖進去，按「開始辨識」。
   - 第一次使用會下載辨識模型，大約 45～60 MB，之後瀏覽器會記住。
   - 客人留言截圖：目前支援 LINE 社群留言，格式像「幽靈兔+2／總金額：800／運費：分擔」。
   - 商品型錄截圖：只能可靠讀出日文品名和 ¥ 價格。中文名稱和台幣售價請在核對表自己填。
3. 核對表格：
   - 有橘色 ⚠ 的列要看一下。
   - 如果出現「留言裡有這些品項，但商品表還沒有」，點一下就能加入商品表，再補上售價。
   - 商品名稱改好之後，訂單會自動重新配對、重算金額。
4. 按「下載 Excel」，會拿到 `原檔名_YYYYMMDD_V#.xlsx`。**下次要上傳這個新檔**。

## 產出的 Excel

| 分頁 | 內容 |
|---|---|
| Dashboard | 整體 KPI、批次查詢（C10 下拉）、客人查詢（F10 輸入）、本次批次進貨統計、批次總覽 |
| Product | 原本全部商品＋新商品（淡黃底色），多了「日文品名」和「未出貨需求量」欄 |
| Orders | 原本全部訂單＋新訂單（淡黃底色），售價會自動對 Product 查價，並保留下拉選單 |
| 其他 | 車次、賣貨便、ShippingList 照原樣保留數值與日期格式；History 是 Excel 保留字，所以改名為 `History_出貨歷史` |

## 已知限制

- 手寫字（型錄上的紅筆中文名與台幣價）辨識不可靠，需要手動填。
- 留言解析規則目前只用 LINE 社群的截圖測過；FB、IG 的留言排版可能拆不準，核對表可以手動修。
- 辨識模型與執行引擎來自 Hugging Face 與 jsDelivr 這兩個免費服務；如果它們連不上，辨識就不能用，但 Excel 功能不受影響。

## 開發者用

```bash
npm install
npm run build   # 重新打包 OCR 背景執行緒（src/ocr-worker.js → public/vendor/ocr-worker.js）
npm test        # Excel 產生測試 + 留言解析規則測試
npm run dev     # 本機開 http://localhost:5173
```

網站本身是 `public/` 資料夾裡的靜態檔案。推送到 `main` 後，GitHub Actions（`.github/workflows/pages.yml`）會自動發布到 GitHub Pages。
依 GitHub Pages 條款，本站不可用於收費的商業服務（SaaS）或交易。
舊的 Claude API 版封存在 `_archive/claude-api版_20260925/`。
