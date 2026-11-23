import datetime
import glob
import hashlib
import html as htmllib
import io
import json
import math
import os
import re
import sys
import unicodedata
from urllib.parse import urljoin

import pdfplumber
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties

PAGE = "https://www.wli.com.hk/SitemapDividendPhilosophy.jhtml"
NA = ("不適用", "N/A", "尚未推出", "產品下架")
EXPECT_2025 = (19, 16)        # 探查所得：資料行、數值個數
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
S = re.S | re.I
PCT = re.compile(r"^(\d+(?:\.\d+)?)[%％]$")
TITLE = re.compile(r"^(\d{4})年度分紅實現率$")
CAT = re.compile(r"^產品類[別别]\s*[：:]\s*(.+)$")
NUM = re.compile(r"^\d+\+?$")
ISSUE = re.compile(r"^[(（](\d{4})[)）]$")
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
API = "https://api.notion.com/v1"
VERSION = "2025-09-03"
AUTO = "自動生成"
FORMAT = "v1"            # 改了 Excel 格式時改這個值，下次運行就會重新上傳
HINTS = {
    401: "NOTION_TOKEN 無效：請重新貼上「AIA 自動更新」的 Internal Integration Secret",
    404: "找不到資料庫：請確認資料庫已連接「AIA 自動更新」，以及 NOTION_DATABASE_ID 正確",
}
N_FIXED = 4              # 產品名稱、產品種類、貨幣、紅利類型
HEAD_ROW = 3
FIRST = 4
ROW_H = 24
SEP_H = 8
YEAR_W = 13
LOW_LIMIT = 0.8
HIGH_LIMIT = 1.0
NAVY_HEX = "1F4E79"
NAME_COLOR = "FF1F2933"
CUR_COLOR = "FF7030A0"
SEP_COLOR = "FFE9ECEF"


class Stop(Exception):
    pass


# ---------- 取得官方 PDF ----------
def clean(s):
    return re.sub(r"\s+", " ", htmllib.unescape(re.sub(r"<[^>]+>", " ", s))).strip()


def squash(s):
    return re.sub(r"\s+", "", s or "")


def tidy(s):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", s or "")).strip()


def fetch_online():
    """試連官網；GitHub Actions 目前連不上，所以失敗時只記錄，不停止。"""
    try:
        r = requests.get(PAGE, headers=UA, timeout=(15, 60))
        if r.status_code != 200:
            print(f"官網回應 HTTP {r.status_code}，改用 repo 內的 PDF")
            return None
        page = r.content.decode("utf-8", errors="replace")
        for href, inner in re.findall(r'<a\b[^>]*href\s*=\s*["\']([^"\']+\.pdf)["\'][^>]*>(.*?)</a>',
                                      page, flags=S):
            if "分紅實現率" not in clean(inner):
                continue
            url = urljoin(PAGE, htmllib.unescape(href).strip().replace(" ", "%20"))
            p = requests.get(url, headers=UA, timeout=(15, 120))
            if p.status_code == 200 and p.content.startswith(b"%PDF"):
                return f"官網：{url}", p.content
        print("官網頁面找不到分紅實現率 PDF，改用 repo 內的 PDF")
    except requests.RequestException as e:
        print(f"官網連不上（{type(e).__name__}），改用 repo 內的 PDF")
    return None


def load_sources():
    out = []
    online = fetch_online()
    if online:
        out.append(online)
    for path in sorted(glob.glob("*.pdf") + glob.glob("pdf/*.pdf")):
        with open(path, "rb") as fh:
            out.append((f"repo：{path}", fh.read()))
    if not out:
        raise Stop(f"官網連不上，repo 內也沒有 PDF：請從 {PAGE} 下載「分紅保險產品之分紅實現率」放入 repo 根目錄")
    return out


# ---------- 讀取 PDF ----------
def values_of(row):
    return [squash(c) for c in row[2:] if squash(c)]


def is_data(row):
    vals = values_of(row)
    return (len(row) > 2 and bool(squash(row[0])) and bool(squash(row[1])) and bool(vals)
            and all(PCT.match(v) or v in NA for v in vals))


def is_data_line(line):
    return any(t in line for t in NA) or bool(re.search(r"\d\s*[%％]", line))


def products_of(lines, where):
    """由頁面文字找出每個產品的（名稱, 類別, 「產品類別」所在行）。"""
    out = []
    for i, ln in enumerate(lines):
        m = CAT.match(ln)
        if not m:
            continue
        parts, j = [], i - 1
        while j >= 0 and not TITLE.match(squash(lines[j])) and not is_data_line(lines[j]) \
                and not CAT.match(lines[j]):
            parts.insert(0, lines[j])
            j -= 1
        name = tidy(" ".join(parts))
        if not name or len(parts) > 2:
            raise Stop(f"{where}：找不到「{ln}」對應的產品名稱")
        out.append((name, tidy(m.group(1)), i))
    return out


def parse_header(head, where):
    if not any(len(r) > 1 and "紅利類" in squash(r[0]) and "貨幣" in squash(r[1]) for r in head):
        raise Stop(f"{where}：表格標題不是「紅利類別｜保單貨幣」")
    nums = issue = None
    for r in head:
        cells = values_of(r)
        if not cells:
            continue
        if all(NUM.match(c) for c in cells):
            nums = cells
        elif all(ISSUE.match(c) for c in cells):
            issue = [ISSUE.match(c).group(1) for c in cells]
    if not nums or not issue or len(nums) != len(issue):
        raise Stop(f"{where}：看不懂保單年度或保單簽發年一列")
    return list(zip(nums, issue))


def parse(content, label):
    with pdfplumber.open(io.BytesIO(content)) as doc:
        m = re.match(r"D:(\d{4})(\d{2})(\d{2})", str((doc.metadata or {}).get("CreationDate") or ""))
        created = "-".join(m.groups()) if m else ""
        year, years, rows = None, None, []
        pct_text = 0
        na_text = dict.fromkeys(NA, 0)
        for no, page in enumerate(doc.pages, 1):
            where = f"{label} 第 {no} 頁"
            text = page.extract_text() or ""
            lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
            for ln in lines:
                t = TITLE.match(squash(ln))
                if t:
                    if year is not None and int(t.group(1)) != year:
                        raise Stop(f"{where}：報告年度與前面不同")
                    year = int(t.group(1))
            prods = products_of(lines, where)
            tables = [[[c or "" for c in r] for r in tb] for tb in page.extract_tables()]
            tables = [g for g in tables if any(is_data(r) for r in g)]
            if len(prods) != len(tables):
                raise Stop(f"{where}：產品標題 {len(prods)} 個，但數據表格 {len(tables)} 個")
            if not prods:
                continue
            pct_text += len(re.findall(r"\d+(?:\.\d+)?\s*[%％]", text))
            for t in NA:
                na_text[t] += text.count(t)
            for k, ((name, cat, at), grid) in enumerate(zip(prods, tables)):
                end = prods[k + 1][2] if k + 1 < len(prods) else len(lines)
                segment = {squash(x) for x in lines[at + 1:end]}
                first = next(i for i, r in enumerate(grid) if is_data(r))
                ys = parse_header(grid[:first], where)
                if years is None:
                    years = ys
                elif ys != years:
                    raise Stop(f"{where}：{name} 的年度欄與前面不同")
                for r in grid[first:]:
                    if not any(squash(c) for c in r):
                        continue
                    if not is_data(r):
                        raise Stop(f"{where} 有看不懂的一行：{'｜'.join(squash(c) for c in r if squash(c))[:80]}")
                    kind, ccy, cells = squash(r[0]), squash(r[1]), values_of(r)
                    if len(cells) != len(years):
                        raise Stop(f"{where}：{name} 有一行的格數不符（{len(cells)} 格）")
                    if kind + ccy + "".join(cells) not in segment:
                        raise Stop(f"{where}：{name}／{kind}／{ccy} 的表格內容與頁面文字對不上")
                    vals = [float(PCT.match(c).group(1)) if PCT.match(c) else None for c in cells]
                    rows.append({"name": name, "cat": cat, "ccy": ccy, "kind": kind,
                                 "cells": cells, "vals": vals})

    if year is None or not rows or not years:
        raise Stop(f"{label}：讀不到報告年度或數據，可能不是分紅實現率的中文 PDF，或版面已改")
    if years[0][0] != "1" or int(years[0][1]) != year - 1:
        raise Stop(f"{label}：報告年度 {year} 與第 1 個保單年度的簽發年（{years[0][1]}）不相符")
    keys = [(r["name"], r["ccy"], r["kind"]) for r in rows]
    dup = sorted({k for k in keys if keys.count(k) > 1})
    if dup:
        raise Stop(f"{label}：有重複的資料行：{dup[:5]}")
    n_num = sum(1 for r in rows for v in r["vals"] if v is not None)
    if n_num != pct_text:
        raise Stop(f"{label}：表格讀到 {n_num} 個數值，但頁面文字有 {pct_text} 個百分比，可能漏讀")
    na_table = {t: sum(1 for r in rows for c in r["cells"] if c == t) for t in NA}
    return {"year": year, "years": years, "rows": rows, "created": created, "label": label,
            "na_table": na_table, "na_text": na_text}


# ---------- 保誠格式 ----------
def solid(argb):
    return PatternFill(fill_type="solid", start_color=argb, end_color=argb)


DARK = solid("FF" + NAVY_HEX)
LIGHT = solid("FFDDEBF7")
PALE = solid("FFF4F7FB")
WHITE = solid("FFFFFFFF")
SEP_FILL = solid(SEP_COLOR)
GOOD_FILL = solid("FFE2F0D9")
LOW_FILL = solid("FFFCE4E4")
WHITE_BOLD = Font(bold=True, color="FFFFFFFF")
BLACK_BOLD = Font(bold=True, color="FF000000")
NAME_FONT = Font(bold=True, color=NAME_COLOR)
CUR_FONT = Font(bold=True, color=CUR_COLOR)
LINK_FONT = Font(bold=True, underline="single", color="FF" + NAVY_HEX)
GREY_FONT = Font(color="FFBFBFBF")
GOOD_FONT = Font(color="FF375623")
LOW_FONT = Font(color="FFC00000")
TITLE_FONT = Font(bold=True, size=16, color="FFFFFFFF")
NOTE_FONT = Font(size=10, color="FF44546A")
SECTION_FONT = Font(bold=True, size=12, color="FF" + NAVY_HEX)
LINE = Side(style="thin", color="FFD0D7E2")
GROUP = Side(style="thin", color="FF7F9BBF")
BORDER = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
CENTER_FIT = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
LEFT_FIT = Alignment(horizontal="left", vertical="center", shrink_to_fit=True, indent=1)


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def pct_format(v):
    return "0%" if abs(v * 100 - round(v * 100)) < 1e-9 else "0.0%"


def mean_or_slash(nums):
    return sum(nums) / len(nums) if nums else "/"


def numbers(rows):
    return [v for r in rows for v in r[N_FIXED:] if is_num(v)]


def to_rows(data):
    return [[r["name"], r["cat"], r["ccy"], r["kind"]]
            + ["/" if v is None else round(v / 100, 4) for v in r["vals"]] for r in data["rows"]]


def make_header(years, fixed):
    return list(fixed) + [f"保单年度{n}\n{y}" for n, y in years]


def build_summary(years, groups):
    width = N_FIXED + len(years)
    head = ["所有年份平均值"] + [f"持滿{n}年平均值" for n, _ in years[1:]]
    cum = []
    for label, sub in groups:
        if not sub:
            continue
        line = [label]
        for start in range(N_FIXED, width):
            line.append(mean_or_slash([r[c] for r in sub for c in range(start, width) if is_num(r[c])]))
        cum.append(line)
    all_label, everything = groups[0]
    year_line = [all_label] + [mean_or_slash([r[c] for r in everything if is_num(r[c])])
                               for c in range(N_FIXED, width)]
    return {"head": head, "cum": cum, "year_head": [y for _, y in years], "year": [year_line]}


def write_row(ws, r, col, values):
    for i, v in enumerate(values):
        ws.cell(row=r, column=col + i, value=v)


def dark(cell):
    cell.fill = DARK
    cell.font = WHITE_BOLD


def light(cell):
    cell.fill = LIGHT
    cell.font = BLACK_BOLD


def heat(cell, v, base):
    if not is_num(v):
        cell.fill = base
        cell.font = GREY_FONT
    elif v >= HIGH_LIMIT:
        cell.fill = GOOD_FILL
        cell.font = GOOD_FONT
    elif v < LOW_LIMIT:
        cell.fill = LOW_FILL
        cell.font = LOW_FONT
    else:
        cell.fill = base


def cat_font(cat):
    s = str(cat)
    if "危疾" in s:
        return Font(color="FF2E7D32")
    if "年金" in s:
        return Font(color="FFC55A11")
    return Font(color="FF2F5597")


def units(s):
    return sum(2 if ord(ch) > 255 else 1 for ch in str(s))


def fit_width(values, low, cap):
    longest = max([units(v) for v in values] or [0])
    return min(max(math.ceil(longest * 1.1) + 4, low), cap)


def add_sheet(wb, title, heading, notes, header, rows, summary, head_label, tab):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    last_col = len(header)
    lows = [30, 12, 10, 12]
    caps = [52, 24, 12, 24]
    widths = [fit_width([header[k]] + [row[k] for row in rows], lows[k], caps[k])
              for k in range(N_FIXED)]

    ws.cell(row=1, column=1, value=heading)
    ws.cell(row=2, column=1, value=notes)
    for rr in (1, 2):
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
        ws.cell(row=rr, column=1).alignment = LEFT
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.cell(row=1, column=1).fill = DARK
    ws.cell(row=2, column=1).font = NOTE_FONT
    ws.cell(row=2, column=1).fill = PALE
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 22 * (notes.count("\n") + 1)

    write_row(ws, HEAD_ROW, 1, header)
    for c in range(1, last_col + 1):
        cell = ws.cell(row=HEAD_ROW, column=c)
        dark(cell)
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[HEAD_ROW].height = 42

    rr = FIRST
    for i, row in enumerate(rows):
        first = i == 0 or rows[i - 1][0] != row[0]
        last = i == len(rows) - 1 or rows[i + 1][0] != row[0]
        if first and i > 0:
            ws.cell(row=rr, column=1).fill = SEP_FILL
            ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
            ws.row_dimensions[rr].height = SEP_H
            rr += 1
        border = Border(left=LINE, right=LINE,
                        top=GROUP if first else LINE,
                        bottom=GROUP if last else LINE)
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=rr, column=c, value=v)
            cell.border = border
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
            if c == 1:
                cell.font = NAME_FONT
            elif c == 2:
                cell.font = cat_font(v)
            elif c == 3:
                cell.font = CUR_FONT
            elif c > N_FIXED:
                if is_num(v):
                    cell.number_format = pct_format(v)
                heat(cell, v, WHITE)
        ws.row_dimensions[rr].height = ROW_H
        rr += 1
    data_end = rr - 1

    if rows:
        ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(last_col)}{data_end}"

    lc = N_FIXED
    r = data_end + 2
    ws.cell(row=r, column=lc, value="平均值摘要").font = SECTION_FONT
    ws.row_dimensions[r].height = ROW_H
    r += 1
    table_start = r
    write_row(ws, r, lc, [head_label] + summary["head"])
    r += 1
    value_rows = []
    for line in summary["cum"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    year_head_row = r
    write_row(ws, r, lc, ["當年平均值"] + summary["year_head"])
    r += 1
    for line in summary["year"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    table_end = r - 1
    for row_i in range(table_start, table_end + 1):
        for c in range(lc, last_col + 1):
            cell = ws.cell(row=row_i, column=c)
            cell.alignment = CENTER
            cell.border = BORDER
    for c in range(lc, last_col + 1):
        dark(ws.cell(row=table_start, column=c))
    dark(ws.cell(row=year_head_row, column=lc))
    for c in range(lc + 1, last_col + 1):
        light(ws.cell(row=year_head_row, column=c))
    for row_i in (table_start, year_head_row):
        ws.row_dimensions[row_i].height = 32
    for vr, line in value_rows:
        label_cell = ws.cell(row=vr, column=lc)
        light(label_cell)
        label_cell.alignment = CENTER_FIT
        ws.row_dimensions[vr].height = ROW_H
        for k, v in enumerate(line[1:]):
            cell = ws.cell(row=vr, column=lc + 1 + k)
            cell.alignment = CENTER_FIT
            cell.number_format = "0.00%"
            heat(cell, v, WHITE)
    note_row = table_end + 1
    ws.cell(row=note_row, column=lc,
            value="「持滿N年平均值」＝保單年度N及之後所有數字的平均；"
                  "「當年平均值」＝該年度一欄所有數字的平均。平均值計入所有貨幣，為固定數字，不會隨篩選變動。")
    ws.merge_cells(start_row=note_row, start_column=lc, end_row=note_row, end_column=last_col)
    ws.cell(row=note_row, column=lc).font = NOTE_FONT
    ws.cell(row=note_row, column=lc).alignment = LEFT
    ws.row_dimensions[note_row].height = 32

    for i in range(1, last_col + 1):
        w = widths[i - 1] if i <= N_FIXED else YEAR_W
        if i == lc:
            w = max(w, 18)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


def groups_by(rows, idx):
    return [(k or "（未分類）", [r for r in rows if r[idx] == k]) for k in dict.fromkeys(r[idx] for r in rows)]


def add_overview(wb, year, today, title, rows, remarks):
    ws = wb.create_sheet("概覽", 0)
    ws.sheet_properties.tabColor = "FFC000"
    ws.sheet_view.showGridLines = False
    last_col = 6
    for i, w in enumerate([36, 18, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def banner(r, value, font, fill, height):
        cell = ws.cell(row=r, column=1, value=value)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = font
        cell.fill = fill
        cell.alignment = LEFT
        ws.row_dimensions[r].height = height

    def section(r, value):
        ws.cell(row=r, column=1, value=value).font = SECTION_FONT
        ws.row_dimensions[r].height = ROW_H

    def stat_table(r, heading, first_col, groups, link=None):
        section(r, heading)
        r += 1
        for c, v in enumerate([first_col, "產品數", "行數", "有數字的格數", "平均值", "低於 80% 的比例"], 1):
            cell = ws.cell(row=r, column=c, value=v)
            dark(cell)
            cell.alignment = CENTER_FIT
            cell.border = BORDER
        ws.row_dimensions[r].height = ROW_H
        for label, sub in groups:
            nums = numbers(sub)
            low = sum(1 for v in nums if v < LOW_LIMIT) / len(nums) if nums else "/"
            vals = [label, len({x[0] for x in sub}), len(sub), len(nums), mean_or_slash(nums), low]
            r += 1
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = BORDER
                cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
                cell.fill = WHITE
            head = ws.cell(row=r, column=1)
            if link:
                head.font = LINK_FONT
                head.hyperlink = Hyperlink(ref=head.coordinate, location=f"'{link}'!A1")
            else:
                head.font = BLACK_BOLD
            ws.cell(row=r, column=5).number_format = "0.00%"
            heat(ws.cell(row=r, column=5), vals[4], WHITE)
            ws.cell(row=r, column=6).number_format = "0.0%"
            ws.row_dimensions[r].height = ROW_H
        return r + 2

    def lines(r, texts):
        for t in texts:
            r += 1
            cell = ws.cell(row=r, column=1, value="・" + t)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
            cell.font = NOTE_FONT
            cell.alignment = LEFT
            ws.row_dimensions[r].height = 18 * max(1, math.ceil(units(t) / 110)) + 6
        return r

    banner(1, f"立橋人壽分紅表現概覽（報告年度 {year}）", TITLE_FONT, DARK, 40)
    banner(2, f"資料來源：立橋人壽網站｜更新日期：{today}", NOTE_FONT, PALE, ROW_H)

    r = stat_table(4, "資料概況", "工作表", [(title, rows)], link=title)
    r = stat_table(r, "按紅利類型", "紅利類型", groups_by(rows, 3))
    r = stat_table(r, "按產品種類", "產品種類", groups_by(rows, 1))
    r = stat_table(r, "按貨幣", "貨幣", groups_by(rows, 2))

    section(r, "顏色說明")
    samples = [(1.0, "100% 或以上：與銷售時的說明一致或更好"),
               (0.79, "低於 80%：明顯低於銷售時的說明"),
               ("/", "沒有數字：官方標示「不適用」、「N/A」、「尚未推出」或「產品下架」")]
    for v, desc in samples:
        r += 1
        cell = ws.cell(row=r, column=1, value=v)
        cell.border = BORDER
        cell.alignment = CENTER_FIT
        if is_num(v):
            cell.number_format = "0%"
        heat(cell, v, WHITE)
        d = ws.cell(row=r, column=2, value=desc)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
        d.alignment = LEFT_FIT
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "備註")
    r = lines(r, remarks)

    r += 2
    section(r, "使用提示")
    lines(r, ["點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
              "工作表第 3 行有篩選箭嘴，可按產品種類、貨幣或紅利類型篩選。",
              "向右或向下捲動時，標題列和產品名稱會固定不動。",
              "平均值摘要在工作表最底部。",
              "手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。"])


def build_excel(path, data, today):
    year, years = data["year"], data["years"]
    rows = to_rows(data)
    legend = ("綠色＝100% 或以上；紅色＝低於 80%；"
              "「/」＝官方標示「不適用」、「N/A」、「尚未推出」或「產品下架」。\n"
              f"資料來源：立橋人壽網站（報告年度 {year}）｜更新日期：{today}")
    notes = ("分紅實現率＝實際派發的非保證紅利 ÷ 銷售時建議書所示的金額。"
             "100% 即與銷售時一致，高於 100% 即比銷售時多。\n"
             "年度欄的年份為保單簽發年份。\n" + legend)
    na = "、".join(f"{t} {n} 格" for t, n in data["na_table"].items() if n)
    remarks = ([f"官方 PDF 建立日期：{data['created']}"] if data["created"] else []) + [
        "以下定義摘自 2025 年度的官方文件，如官方日後修改，以官方文件為準。",
        "涵蓋由 2019 年啟動業務至報告年度有新保單簽發的分紅產品，計算至報告年度之相應保單週年日。",
        "特別紅利／終期紅利的分紅實現率：報告年度終止的保單實際已支付的特別紅利／終期紅利 ÷ 銷售時建議書所示的相應數值。",
        "週年紅利的分紅實現率：報告年度有效的保單實際累積的週年紅利，以及實際紅利及可支取保證現金所累積的利息 ÷ 銷售時建議書所示的相應數值。",
        "歸原紅利的分紅實現率：報告年度有效的保單實際總歸原紅利的現金價值 ÷ 銷售時建議書所示的相應數值。",
        "「不適用」的原因：相應保單年度沒有生效保單；建議書所示的紅利為零；或（特別紅利／終期紅利）報告年度沒有相關保單終結。",
        f"沒有數字的格（顯示為「/」）：{na}。",
        f"官方來源：{PAGE}（下載「分紅保險產品之分紅實現率」）",
        f"本檔案讀取的 PDF：{data['label']}",
    ]
    groups = [("所有分紅", rows)]
    for label, sub in groups_by(rows, 3):
        if numbers(sub):
            groups.append((label, sub))

    title = f"{year}分紅實現率"
    wb = Workbook()
    wb.remove(wb.active)
    add_sheet(wb, title, f"立橋人壽分紅實現率（報告年度 {year}）", notes,
              make_header(years, ["产品名称", "产品种类", "货币", "红利类型"]), rows,
              build_summary(years, groups), "分紅類型", "2E75B6")
    add_overview(wb, year, today, title, rows, remarks)
    wb.save(path)
    print("工作表：", "、".join(wb.sheetnames))


# ---------- Notion（與泰禾版相同做法） ----------
def check(r, what):
    if r.status_code >= 300:
        hint = HINTS.get(r.status_code, "")
        raise Stop(f"{what} 失敗：{r.status_code} {r.text[:300]}" + (f"\n→ {hint}" if hint else ""))
    return r.json()


def resolve_source(raw, headers):
    m = re.search(r"[0-9a-fA-F]{32}", raw.replace("-", ""))
    ident = m.group(0) if m else raw

    def first_source(db_id):
        r = requests.get(f"{API}/databases/{db_id}", headers=headers, timeout=30)
        if r.status_code == 200:
            ds = r.json().get("data_sources") or []
            return ds[0]["id"] if ds else None
        return None

    found = first_source(ident)
    if found:
        return found
    r = requests.get(f"{API}/data_sources/{ident}", headers=headers, timeout=30)
    if r.status_code == 200:
        return ident
    status = r.status_code
    r = requests.get(f"{API}/blocks/{ident}/children?page_size=100", headers=headers, timeout=30)
    if r.status_code == 200:
        for b in r.json().get("results", []):
            if b.get("type") == "child_database":
                found = first_source(b["id"])
                if found:
                    return found
    raise Stop(f"找不到立橋的年份資料庫（狀態 {status}）。401＝金鑰不正確；"
               "404＝頁面未連接「AIA 自動更新」或 NOTION_DATABASE_ID 不正確。")


def page_title(page):
    for prop in page.get("properties", {}).values():
        if prop.get("type") == "title":
            return "".join(t.get("plain_text", "") for t in prop.get("title", []))
    return ""


def file_blocks(page_id, headers):
    r = requests.get(f"{API}/blocks/{page_id}/children?page_size=100", headers=headers, timeout=30)
    out = []
    for block in check(r, "讀取頁面內容").get("results", []):
        if block["type"] == "file":
            info = block["file"]
            cap = "".join(t.get("plain_text", "") for t in info.get("caption", []))
            out.append((block["id"], info.get("name", ""), cap))
    return out


def upload(path, year, digest, today):
    token = os.environ.get("NOTION_TOKEN", "").strip()
    target = os.environ.get("NOTION_DATABASE_ID", "").strip()
    if not token or not target:
        raise Stop("未設定 NOTION_TOKEN 或 NOTION_DATABASE_ID")
    headers = {"Authorization": f"Bearer {token}", "Notion-Version": VERSION,
               "Content-Type": "application/json"}
    source = resolve_source(target, headers)
    r = requests.post(f"{API}/data_sources/{source}/query", headers=headers, json={"page_size": 100}, timeout=30)
    pages = check(r, "讀取年份資料庫").get("results", [])
    page = next((p for p in pages if str(year) in page_title(p)), None)
    if page is None:
        body = {"parent": {"type": "data_source_id", "data_source_id": source},
                "properties": {"title": {"title": [{"text": {"content": str(year)}}]}}}
        r = requests.post(f"{API}/pages", headers=headers, json=body, timeout=30)
        page = check(r, "建立年度頁面")
        print(f"已建立頁面：{year}")
    else:
        print("使用現有頁面：", page_title(page))
    page_id = page["id"]

    blocks = file_blocks(page_id, headers)
    if any(AUTO in name and digest in cap for _, name, cap in blocks):
        print("數據和格式與上次上傳相同，略過上傳")
        return
    old_auto = [bid for bid, name, _ in blocks if AUTO in name]

    filename = os.path.basename(path)
    r = requests.post(f"{API}/file_uploads", headers=headers,
                      json={"mode": "single_part", "filename": filename, "content_type": XLSX}, timeout=30)
    upload_id = check(r, "建立上傳")["id"]
    with open(path, "rb") as fh:
        r = requests.post(f"{API}/file_uploads/{upload_id}/send",
                          headers={"Authorization": f"Bearer {token}", "Notion-Version": VERSION},
                          files={"file": (filename, fh.read(), XLSX)}, timeout=120)
    check(r, "傳送檔案")
    block = {"object": "block", "type": "file",
             "file": {"type": "file_upload", "file_upload": {"id": upload_id},
                      "caption": [{"type": "text", "text": {"content": f"更新 {today}｜{digest}"}}]}}
    r = requests.patch(f"{API}/blocks/{page_id}/children", headers=headers, json={"children": [block]}, timeout=60)
    check(r, "加入檔案到頁面")
    for bid in old_auto:
        r = requests.delete(f"{API}/blocks/{bid}", headers=headers, timeout=30)
        check(r, "刪除上一次的自動生成檔案")
    print(f"已上傳：{filename}（刪除舊的自動生成檔案 {len(old_auto)} 個）")


# ---------- 主流程 ----------
def main():
    parsed = []
    for label, content in load_sources():
        if not content.startswith(b"%PDF"):
            raise Stop(f"{label} 不是 PDF")
        d = parse(content, label)
        print(f"已讀取 {label}：報告年度 {d['year']}｜PDF 建立日期 {d['created'] or '讀不到'}｜{len(d['rows'])} 行")
        parsed.append(d)
    data = max(parsed, key=lambda d: (d["year"], d["created"]))
    year, rows = data["year"], data["rows"]
    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
    today = now.strftime("%Y-%m-%d")
    print(f"採用：{data['label']}｜報告年度 {year}")
    print("年度欄：", " ".join(f"{n}({y})" for n, y in data["years"]))

    last = None
    for r in rows:
        if r["name"] != last:
            last = r["name"]
            print(f"{r['name']}（{r['cat']}）")
        print(f"   {r['kind']}｜{r['ccy']}｜" + " ".join("/" if v is None else f"{v:g}%" for v in r["vals"]))

    nums = [v for r in rows for v in r["vals"] if v is not None]
    same = data["na_table"] == data["na_text"]
    print("沒有數字的格：" + "｜".join(f"{t} {n}" for t, n in data["na_table"].items())
          + ("｜與頁面文字一致 ✓" if same else f"｜頁面文字為 {data['na_text']} ⚠（請抽查）"))
    if nums:
        stats = (f"平均 {sum(nums) / len(nums):.2f}%｜範圍 {min(nums):g}%–{max(nums):g}%"
                 f"｜低於 80%：{sum(1 for x in nums if x < 80)}｜100% 或以上：{sum(1 for x in nums if x >= 100)}")
    else:
        stats = "沒有任何數值"
    print(f"合計：產品 {len({r['name'] for r in rows})}｜資料行 {len(rows)}｜數值 {len(nums)}｜{stats}"
          f"｜沒有數字 {sum(data['na_table'].values())}")
    if year == 2025:
        ok = (len(rows), len(nums)) == EXPECT_2025
        print("   與探查結果" + ("一致 ✓" if ok else f"不一致 ⚠（探查為 {EXPECT_2025}）"))
    if year < now.year - 1:
        print(f"⚠ 提醒：目前最新的 PDF 是 {year} 報告年度。如官網已公布新一年的分紅實現率，"
              f"請下載並放入 repo 根目錄：{PAGE}")

    path = f"立橋人壽分紅實現率（{year}年）{AUTO}.xlsx"
    build_excel(path, data, today)
    print(f"已生成 {path}（{os.path.getsize(path) // 1024} KB）")

    digest = hashlib.sha256(json.dumps(
        [FORMAT, data["years"]] + [[r["name"], r["cat"], r["ccy"], r["kind"], r["cells"]] for r in rows],
        ensure_ascii=False).encode()).hexdigest()[:12]
    if os.environ.get("UPLOAD", "").lower() == "true":
        upload(path, year, digest, today)
        print("完成")
    else:
        print("只讀試跑：沒有上傳")


if __name__ == "__main__":
    try:
        main()
    except Stop as e:
        print(f"✗ 已停止：{e}")
        sys.exit(1)
