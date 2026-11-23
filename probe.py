import html as htmllib
import io
import re
from urllib.parse import urljoin, urlparse

import pdfplumber
import requests

START = "https://www.wli.com.hk/SitemapDividendPhilosophy.jhtml"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
WORDS = ("立橋", "分紅", "紅利", "實現率", "現金價值", "派息", "澳門")
KEY = re.compile(r"實現率|現金價值|紅利|分紅|fulfil|ratio|dividend|bonus", re.I)
S = re.S | re.I
MAX_PDF, MAX_SUB = 6, 5


def clean(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", " ", s, flags=S)
    return re.sub(r"\s+", " ", htmllib.unescape(re.sub(r"<[^>]+>", " ", s))).strip()


def get(url):
    for label, headers in (("模擬 Chrome", UA), ("普通連線", {})):
        try:
            r = requests.get(url, headers=headers, timeout=60)
            print(f"[{label}] 狀態 {r.status_code}｜{len(r.content)} bytes｜"
                  f"{r.headers.get('content-type')}｜last-modified {r.headers.get('last-modified')}"
                  f"｜server {r.headers.get('server')}｜最終網址 {r.url}")
            if r.status_code == 200:
                return r
        except requests.RequestException as e:
            print(f"[{label}] {type(e).__name__}: {e}")
    return None


def text_of(r):
    for enc in ("utf-8", "big5hkscs"):
        try:
            return r.content.decode(enc)
        except UnicodeDecodeError:
            pass
    return r.content.decode("utf-8", errors="replace")


def scan(url):
    """列印頁面概況，回傳（PDF 連結, 相關子頁）。"""
    print()
    print(f"=== 頁面：{url} ===")
    r = get(url)
    if r is None:
        return [], []
    page = text_of(r)
    t = re.search(r"<title[^>]*>(.*?)</title>", page, flags=S)
    print("標題：", clean(t.group(1)) if t else "（無）")
    print("關鍵字：", "｜".join(f"{w} {page.count(w)}" for w in WORDS))
    body = clean(page)
    print(f"純文字 {len(body)} 字；開頭：{body[:300]}")

    tables = re.findall(r"<table\b.*?</table>", page, flags=S)
    print(f"--- HTML 表格 {len(tables)} 個 ---")
    for k, tb in enumerate(tables[:8], 1):
        trs = re.findall(r"<tr\b.*?</tr>", tb, flags=S)
        print(f"[表格 {k}] {len(trs)} 行")
        for tr in trs[:25]:
            cells = [clean(c) for c in re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=S)]
            print("H|", "｜".join(cells)[:220])

    for src in re.findall(r'<iframe\b[^>]*src=["\']([^"\']+)["\']', page, flags=S):
        print("iframe：", urljoin(url, src))
    opts = [clean(o) for o in re.findall(r"<option\b[^>]*>(.*?)</option>", page, flags=S)]
    if opts:
        print(f"下拉選項 {len(opts)} 個：", "｜".join(opts[:40]))

    pdfs, subs = [], []
    print("--- 相關連結 ---")
    for href, inner in re.findall(r'<a\b[^>]*href\s*=\s*["\']([^"\']+)["\'][^>]*>(.*?)</a>', page, flags=S):
        text, target = clean(inner), href.strip()
        if target.lower().startswith("javascript"):
            m = re.search(r"['\"]([^'\"]+)['\"]", target)
            target = m.group(1) if m else ""
        if not target or target.startswith("#"):
            continue
        full = urljoin(url, target)
        if ".pdf" in full.lower():
            print(f"PDF｜{text[:60]} → {full}")
            pdfs.append((text, full))
        elif KEY.search(text) or KEY.search(target):
            print(f"頁面｜{text[:60]} → {full}")
            if urlparse(full).netloc == urlparse(url).netloc:
                subs.append(full)
    for raw in re.findall(r"[\w./%:\-~]+\.pdf", page, flags=re.I):
        full = urljoin(url, raw)
        if all(full != u for _, u in pdfs):
            print(f"PDF（非連結）→ {full}")
            pdfs.append(("", full))
    return pdfs, subs


def dump_pdf(text, url):
    print()
    print(f"=== PDF：{text[:60]}｜{url} ===")
    r = get(url)
    if r is None or not r.content.startswith(b"%PDF"):
        print("不是 PDF 或讀取失敗")
        return
    doc = pdfplumber.open(io.BytesIO(r.content))
    print("文件資料：", doc.metadata)
    print(f"頁數：{len(doc.pages)}")
    total_pct = 0
    for i, p in enumerate(doc.pages[:20], 1):
        body = p.extract_text() or ""
        lines = [ln for ln in body.split("\n") if ln.strip()]
        tables = p.extract_tables()
        pcts = len(re.findall(r"\d+(?:\.\d+)?\s*[%％]", body))
        total_pct += pcts
        print()
        print(f"----- 第 {i} 頁｜{p.width:.0f}x{p.height:.0f}｜文字 {len(lines)} 行｜表格 {len(tables)} 個"
              f"｜百分比 {pcts} 個 -----")
        limit = 45 if i <= 4 else 10
        for ln in lines[:limit]:
            print("T|", ln[:160])
        if len(lines) > limit:
            print(f"T| …（餘下 {len(lines) - limit} 行略）")
        for k, tb in enumerate(tables, 1):
            cols = max((len(row) for row in tb), default=0)
            print(f"[表格 {k}] {len(tb)} 行 x {cols} 欄")
            for row in tb[:25 if i <= 4 else 8]:
                print("R|", "｜".join("" if c is None else str(c).replace("\n", "⏎") for c in row)[:220])
    if len(doc.pages) > 20:
        print(f"…（只列出前 20 頁，共 {len(doc.pages)} 頁）")
    print(f"已列出頁面的百分比合計：{total_pct} 個")


if __name__ == "__main__":
    pdfs, subs = scan(START)
    seen = {START}
    for sub in subs:
        if sub in seen or len(seen) > MAX_SUB:
            continue
        seen.add(sub)
        more, _ = scan(sub)
        pdfs += more
    unique = []
    for text, u in pdfs:
        if all(u != x for _, x in unique):
            unique.append((text, u))
    unique.sort(key=lambda x: 0 if KEY.search(x[0] + x[1]) else 1)
    print()
    print(f"共找到 PDF {len(unique)} 份，讀取前 {MAX_PDF} 份（與分紅相關的優先）")
    for text, u in unique[:MAX_PDF]:
        dump_pdf(text, u)
    print()
    print("探查完成（只讀，沒有上傳）")
