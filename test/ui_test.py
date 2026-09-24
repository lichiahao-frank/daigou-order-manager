# 用無頭瀏覽器把網頁流程完整點一次：上傳 Excel → 丟截圖 → 辨識 → 核對 → 下載
import sys, pathlib
from playwright.sync_api import sync_playwright
ROOT = pathlib.Path(__file__).resolve().parents[2]
OUT = pathlib.Path(__file__).parent
url = sys.argv[1] if len(sys.argv) > 1 else 'http://localhost:5173'
with sync_playwright() as p:
    b = p.chromium.launch(channel='chrome')
    pg = b.new_page(viewport={'width': 1280, 'height': 900})
    logs = []
    pg.on('console', lambda m: logs.append(f'{m.type}: {m.text}'))
    pg.on('pageerror', lambda e: logs.append(f'PAGEERROR: {e}'))
    pg.goto(url)
    pg.set_input_files('#fileExcel', str(ROOT / '吉伊卡哇登記_最新.xlsx'))
    pg.wait_for_selector('#excelInfo .pill')
    print('Excel:', pg.inner_text('#excelInfo').replace('\n', ' | '))
    pg.fill('#batch', '0924樂園')
    pg.set_input_files('#fileShots', [str(ROOT / 'S__27074572_0.jpg'), str(ROOT / 'S__27074571_0.jpg')])
    pg.click('#btnParse')
    pg.wait_for_function("() => ![...document.querySelectorAll('.thumb .tag')].some(t => /等待|辨識中/.test(t.textContent))", timeout=300000)
    print('Status:', pg.inner_text('#parseStatus'), '|', pg.inner_text('#engineStatus'))
    print('Thumbs:', [t.inner_text().replace('\n', ' ') for t in pg.query_selector_all('.thumb')])
    pg.wait_for_selector('#cardReview:not(.hidden)', timeout=5000)
    print('Missing before edit:', pg.inner_text('#missingItems'))
    print('Orders before edit:', pg.eval_on_selector_all('#tbOrders tr', 'rs => rs.map(r => r.querySelector("[data-k=item]").value + " / " + (r.querySelector(".warn")||{}).innerText)'))
    # 模擬賣家修正商品表（依型錄順序：南瓜、ハチワレ、ちいかわ、うさぎ）
    fixes = [('南瓜S娃', None, '630'), ('幽靈八', '1540', '400'), ('幽靈吉', None, '400'), ('幽靈兔', '1540', '400')]
    for i, (name, jpy, price) in enumerate(fixes):
        for k, v in (('name', name), ('jpy', jpy), ('price', price)):
            if v is None: continue
            loc = pg.locator(f'#tbProducts [data-p="{i}"][data-k="{k}"]')
            loc.fill(v); loc.dispatch_event('change')
    print('Missing after edit:', repr(pg.inner_text('#missingItems')))
    print('Products rows:', pg.eval_on_selector_all('#tbProducts tr', 'rs => rs.map(r => [...r.querySelectorAll("input")].map(i => i.type==="checkbox"? i.checked : i.value).concat(r.innerText.trim()))'))
    print('Orders rows:', pg.eval_on_selector_all('#tbOrders tr', 'rs => rs.map(r => r.innerText.replace(/\\s+/g," ").trim() + " | " + [...r.querySelectorAll("input,select")].map(i => i.type==="checkbox"? i.checked : i.value).join(","))'))
    print('Summary:', pg.inner_text('#exportSummary'))
    pg.screenshot(path=str(OUT / 'ui_review.png'), full_page=True)
    with pg.expect_download() as dl:
        pg.click('#btnExport')
    d = dl.value
    target = OUT / ('ui_' + d.suggested_filename)
    d.save_as(str(target))
    print('Downloaded:', target.name)
    print('Msg:', pg.inner_text('#exportMsg'))
    print('Console:', logs)
    b.close()
