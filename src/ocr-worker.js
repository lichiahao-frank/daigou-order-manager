// OCR 背景執行緒：在使用者的瀏覽器裡跑 PaddleOCR（PP-OCRv6 small，支援繁中與日文）
// 模型第一次會從 Hugging Face 下載（約 30 MB），執行引擎從 jsDelivr 載入，之後由瀏覽器快取
import { PaddleOcrService, V6_SMALL_MODEL } from 'ppu-paddle-ocr/web';

let service = null;
let ready = null;

function init() {
  if (!ready) {
    service = new PaddleOcrService({ model: { ...V6_SMALL_MODEL } });
    ready = service.initialize().catch((err) => {
      ready = null; // 下載失敗時允許重試
      throw err;
    });
  }
  return ready;
}

self.onmessage = async (event) => {
  const { id, type, buffer } = event.data;
  try {
    await init();
    if (type === 'init') return self.postMessage({ id, ok: true });
    const result = await service.recognize(buffer);
    const lines = (result.lines || []).map((line) =>
      line.map((w) => ({ text: w.text, box: w.box, confidence: w.confidence })),
    );
    self.postMessage({ id, ok: true, lines });
  } catch (err) {
    self.postMessage({ id, ok: false, error: String((err && err.message) || err) });
  }
};
