import glob
import io
import re
import socket

import pdfplumber
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
PDF_PATH = "/editor_upload_image/files/2025%20Fulfillment%20Ratio%20(Chinese).pdf"
URLS = [
    "https://www.wli.com.hk/SitemapDividendPhilosophy.jhtml",
    "https://wli.com.hk/SitemapDividendPhilosophy.jhtml",
    "http://www.wli.com.hk/SitemapDividendPhilosophy.jhtml",
    "https://www.wli.com.hk/tc/dividend",
    "https://wli.com.hk/tc/dividend",
    "https://www.wli.com.hk" + PDF_PATH,
    "https://wli.com.hk" + PDF_PATH,
]


def fetch(url):
    try:
        r = requests.get(url, headers=UA, timeout=(15, 60))
    except requests.RequestException as e:
        print(f"✗ {url}")
        print(f"   {type(e).__name__}")
        return None
    mark = "✓" if r.status_code == 200 else "✗"
    print(f"{mark} {url}")
    print(f"   狀態 {r.status_code}｜{len(r.content)} bytes｜{r.headers.get('content-type')}"
          f"｜server {r.headers.get('server')}｜最終網址 {r.url}")
    return r.content if r.status_code == 200 else None


def dump(name, content):
    print()
    print(f"=== PDF：{name}（{len(content) // 1024} KB） ===")
    if not content.startswith(b"%PDF"):
        print("不是 PDF")
        return
    with pdfplumber.open(io.BytesIO(content)) as doc:
        print("文件資料：", doc.metadata)
        print(f"頁數：{len(doc.pages)}")
        total = 0
        for i, p in enumerate(doc.pages[:30], 1):
            text = p.extract_text() or ""
            lines = [ln for ln in text.split("\n") if ln.strip()]
            tables = p.extract_tables()
            pcts = len(re.findall(r"\d+(?:\.\d+)?\s*[%％]", text))
            total += pcts
            print()
            print(f"----- 第 {i} 頁｜{p.width:.0f}x{p.height:.0f}｜文字 {len(lines)} 行"
                  f"｜表格 {len(tables)} 個｜百分比 {pcts} 個 -----")
            for ln in lines[:70]:
                print("T|", ln[:170])
            if len(lines) > 70:
                print(f"T| …（餘下 {len(lines) - 70} 行略）")
            for k, tb in enumerate(tables, 1):
                cols = max((len(row) for row in tb), default=0)
                print(f"[表格 {k}] {len(tb)} 行 x {cols} 欄")
                for row in tb[:60]:
                    print("R|", "｜".join("" if c is None else str(c).replace("\n", "⏎") for c in row)[:240])
        if len(doc.pages) > 30:
            print(f"…（只列出前 30 頁，共 {len(doc.pages)} 頁）")
        print()
        print(f"已列出頁面的百分比合計：{total} 個")


if __name__ == "__main__":
    print("=== A. 連線測試 ===")
    for host in ("www.wli.com.hk", "wli.com.hk"):
        try:
            ips = sorted({a[4][0] for a in socket.getaddrinfo(host, 443)})
            print(f"DNS {host} → {', '.join(ips)}")
        except OSError as e:
            print(f"DNS {host} 失敗：{e}")
    online = None
    for url in URLS:
        content = fetch(url)
        if content is None:
            continue
        if content.startswith(b"%PDF"):
            online = online or (url, content)
        else:
            page = content.decode("utf-8", errors="replace")
            links = sorted(set(re.findall(r"[^\"'\s<>]+\.pdf", page, flags=re.I)))
            print(f"   立橋 {page.count('立橋')} 次｜實現率 {page.count('實現率')} 次"
                  f"｜現金價值 {page.count('現金價值')} 次｜PDF 連結：{links[:10]}")

    print()
    print("=== B. PDF 內容 ===")
    local = sorted(glob.glob("*.pdf") + glob.glob("pdf/*.pdf"))
    if online:
        dump("網上 " + online[0], online[1])
    elif local:
        print("網上讀不到，改讀 repo 內的 PDF")
        for path in local:
            with open(path, "rb") as fh:
                dump("repo " + path, fh.read())
    else:
        print("網上讀不到，repo 內也沒有 PDF：請先把官方 PDF 上傳到 repo 根目錄")
    print()
    print("探查完成（只讀，沒有上傳）")
