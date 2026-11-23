import collections
import datetime
import glob
import io
import math
import os
import re

import pdfplumber
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties

TOKEN = os.environ.get("NOTION_TOKEN", "").strip()
API = "https://api.notion.com/v1"
VERSION = "2025-09-03"
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Notion-Version": VERSION,
    "Content-Type": "application/json",
}
SOURCE_PAGE = "https://www.blue.com.hk/hk/participating_policies"
CJK = r"\u3000-\u303f\u4e00-\u9fff\uff00-\uffef"
LINE_RE = re.compile(r"^\d{1,2}/\d{1,2}/(\d{4})\s+\d{1,2}/\d{1,2}/(\d{4})\s+(\d{1,2}\+?)\s+(.+)$")
TYPES = {"terminal": "終期紅利", "reversionary": "歸原紅利", "annual": "週年紅利"}
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
AUTO = "自動生成"
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


# ---------- 讀取微藍 PDF ----------
def product_name(text, path):
    for line in text.splitlines():
        if re.search(r"[\u4e00-\u9fff]", line):
            s = re.sub(r"\s+", " ", line).strip()
            s = re.sub(r"\s*([「」])\s*", r"\1", s)
            return re.sub(r"(?<=[%s])\s+(?=[%s])" % (CJK, CJK), "", s)
    return os.path.splitext(os.path.basename(path))[0]


def category(name):
    if "儲蓄" in name:
        return "儲蓄保險"
    if "人壽" in name:
        return "終身人壽"
    return "其他"


def fail(path, text, why):
    print(f"✗ {os.path.basename(path)}：{why}")
    print("--- PDF 文字（前 80 行） ---")
    for line in text.splitlines()[:80]:
        print(line)
    raise SystemExit(1)


def parse_pdf(path):
    with pdfplumber.open(path) as pdf:
        texts = [p.extract_text() or "" for p in pdf.pages]
        meta = (pdf.metadata or {}).get("CreationDate", "") or ""
    full = "\n".join(texts)
    m = re.search(r"reporting\s+year\s+(\d{4})", full, flags=re.I)
    if not m:
        fail(path, full, "找不到報告年度")
    cells, rtype = {}, ""
    for text in texts:
        hit = False
        for line in text.splitlines():
            g = LINE_RE.match(line.strip())
            if not g:
                continue
            hit = True
            y1, y2, label, rest = g.groups()
            vals = rest.split()
            if len(vals) != 1:
                fail(path, full, f"保單年度 {label} 有 {len(vals)} 個數值，版面可能改了")
            v = vals[0]
            if v.upper() in ("N/A", "NA"):
                value = "/"
            elif re.fullmatch(r"\d+(\.\d+)?%", v):
                value = round(float(v[:-1]) / 100, 4)
            else:
                fail(path, full, f"看不懂的數值：{v}")
                        item = (y1 if y1 == y2 else f"{y1}-{y2}", value)
            if label in cells and cells[label] != item:
                fail(path, full, f"保單年度 {label} 的中英文表格數值不一致")
            cells[label] = item
        if hit and not rtype:
            t = re.search(r"Fulfillment\s+Ratio\s+for\s+(Terminal|Reversionary|Annual)", text, flags=re.I)
            rtype = TYPES[t.group(1).lower()] if t else ""
    if len(cells) < 5:
        fail(path, full, f"只讀到 {len(cells)} 行資料")
    labels = sorted(cells, key=lambda s: int(s.rstrip("+")))
    created = re.search(r"D:(\d{4})(\d{2})(\d{2})", meta)
    return {"file": os.path.basename(path), "name": product_name(texts[0], path), "year": m.group(1),
            "rtype": rtype or "未標示", "labels": labels,
            "spans": [cells[k][0] for k in labels], "vals": [cells[k][1] for k in labels],
            "created": "-".join(created.groups()) if created else "未知"}


# ---------- 平均值 ----------
def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def mean_or_slash(nums):
    return sum(nums) / len(nums) if nums else "/"


def numbers(rows, n_fixed):
    return [v for r in rows for v in r[n_fixed:] if is_num(v)]


def build_summary(n_fixed, labels, spans, groups):
    width = n_fixed + len(labels)
    head = ["所有年份平均值"] + [f"持滿{k}年平均值" for k in labels[1:]]
    cum = []
    for label, sub in groups:
        if not sub:
            continue
        line = [label]
        for start in range(n_fixed, width):
            line.append(mean_or_slash([r[c] for r in sub for c in range(start, width) if is_num(r[c])]))
        cum.append(line)
    all_label, everything = groups[0]
    year_line = [all_label] + [mean_or_slash([r[c] for r in everything if is_num(r[c])])
                               for c in range(n_fixed, width)]
    return {"head": head, "cum": cum, "year_head": list(spans), "year": [year_line]}


def fmt_pct(v):
    return f"{v * 100:.2f}%" if is_num(v) else "/"


# ---------- 生成 Excel（保誠格式） ----------
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
    """100% 或以上綠底；低於 80% 紅底；沒有數字用淺灰字。"""
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
    if "人壽" in str(cat):
        return Font(color="FF2E7D32")
    return Font(color="FF2F5597")


def units(s):
    return sum(2 if ord(ch) > 255 else 1 for ch in str(s))


def fit_width(values, low, cap):
    longest = max([units(v) for v in values] or [0])
    return min(max(math.ceil(longest * 1.1) + 4, low), cap)


def add_sheet(wb, title, heading, notes, header, rows, summary, head_label, n_fixed, tab):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    last_col = len(header)
    lows = [30, 12, 10, 14]
    caps = [52, 24, 12, 44]
    widths = [fit_width([header[k]] + [row[k] for row in rows], lows[k], caps[k])
              for k in range(n_fixed)]

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
    ws.row_dimensions[2].height = 66

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
            elif c > n_fixed:
                if is_num(v):
                    cell.number_format = "0%"
                heat(cell, v, WHITE)
        ws.row_dimensions[rr].height = ROW_H
        rr += 1
    data_end = rr - 1

    if rows:
        ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(last_col)}{data_end}"

    lc = n_fixed
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
                  "「當年平均值」＝該年度一欄所有數字的平均。平均值為固定數字，不會隨篩選變動。")
    ws.merge_cells(start_row=note_row, start_column=lc, end_row=note_row, end_column=last_col)
    ws.cell(row=note_row, column=lc).font = NOTE_FONT
    ws.cell(row=note_row, column=lc).alignment = LEFT
    ws.row_dimensions[note_row].height = 32

    for i in range(1, last_col + 1):
        w = widths[i - 1] if i <= n_fixed else YEAR_W
        if i == lc:
            w = max(w, 18)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


def add_overview(wb, year, today, items, remarks):
    """items：[(簡稱, 工作表名稱, 資料行, 固定欄數)]"""
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

    def table_head(r, values):
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            dark(cell)
            cell.alignment = CENTER_FIT
            cell.border = BORDER
        ws.row_dimensions[r].height = ROW_H

    def bullets(r, lines):
        for t in lines:
            r += 1
            cell = ws.cell(row=r, column=1, value="・" + t)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
            cell.font = NOTE_FONT
            cell.alignment = LEFT_FIT
            ws.row_dimensions[r].height = ROW_H
        return r

    banner(1, f"微藍分紅實現率概覽（報告年度 {year}）", TITLE_FONT, DARK, 40)
    banner(2, f"資料來源：Blue 微藍網站的分紅實現率 PDF｜更新日期：{today}", NOTE_FONT, PALE, ROW_H)

    r = 4
    section(r, "資料概況")
    r += 1
    table_head(r, ["工作表", "產品／版本數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"])
    for _, title, rows, n_fixed in items:
        nums = numbers(rows, n_fixed)
        low = sum(1 for v in nums if v < LOW_LIMIT) / len(nums) if nums else "/"
        vals = [title, len({row[0] for row in rows}), len(rows), len(nums), mean_or_slash(nums), low]
        r += 1
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
        link = ws.cell(row=r, column=1)
        link.font = LINK_FONT
        link.hyperlink = Hyperlink(ref=link.coordinate, location=f"'{title}'!A1")
        ws.cell(row=r, column=5).number_format = "0.00%"
        heat(ws.cell(row=r, column=5), vals[4], WHITE)
        ws.cell(row=r, column=6).number_format = "0.0%"
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "按產品種類的整體平均值")
    r += 1
    table_head(r, ["產品種類"] + [short for short, _, _, _ in items])
    cats = []
    for _, _, rows, _ in items:
        for row in rows:
            if row[1] not in cats:
                cats.append(row[1])
    for cat in cats:
        r += 1
        cell = ws.cell(row=r, column=1, value=cat or "（未分類）")
        cell.border = BORDER
        cell.alignment = LEFT_FIT
        cell.font = Font(bold=True, color=cat_font(cat).color)
        for k, (_, _, rows, n_fixed) in enumerate(items):
            v = mean_or_slash(numbers([row for row in rows if row[1] == cat], n_fixed))
            c2 = ws.cell(row=r, column=k + 2, value=v)
            c2.border = BORDER
            c2.alignment = CENTER_FIT
            c2.number_format = "0.00%"
            heat(c2, v, WHITE)
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "顏色說明")
    samples = [(1.0, "100% 或以上：與銷售時的說明一致或更好"),
               (0.79, "低於 80%：明顯低於銷售時的說明"),
               ("/", "沒有數字：官方 PDF 顯示 N/A")]
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
    section(r, "使用提示")
    r = bullets(r, ["點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
                    "工作表第 3 行有篩選箭嘴，可按產品種類、貨幣或紅利類型篩選。",
                    "向右或向下捲動時，標題列和產品名稱會固定不動。",
                    "平均值摘要在工作表最底部。",
                    "手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。"])
    if remarks:
        r += 2
        section(r, "資料留意")
        bullets(r, remarks)


# ---------- Notion（沿用保誠） ----------
def check(r, what):
    if r.status_code >= 300:
        print(f"{what} 失敗：{r.status_code} {r.text[:400]}")
        raise SystemExit(1)
    return r.json()


def target_id():
    raw = (os.environ.get("NOTION_DATABASE_ID") or os.environ.get("NOTION_DATABASE_URL") or "").strip()
    return raw.split("?")[0].rstrip("/").replace("-", "")[-32:]


def resolve_source(raw):
    m = re.search(r"[0-9a-fA-F]{32}", raw.replace("-", ""))
    ident = m.group(0) if m else raw

    def first_source(db_id):
        r = requests.get(f"{API}/databases/{db_id}", headers=HEADERS, timeout=30)
        if r.status_code == 200:
            ds = r.json().get("data_sources") or []
            return ds[0]["id"] if ds else None
        return None

    found = first_source(ident)
    if found:
        return found
    r = requests.get(f"{API}/data_sources/{ident}", headers=HEADERS, timeout=30)
    if r.status_code == 200:
        return ident
    status = r.status_code
    r = requests.get(f"{API}/blocks/{ident}/children?page_size=100", headers=HEADERS, timeout=30)
    if r.status_code == 200:
        for b in r.json().get("results", []):
            if b.get("type") == "child_database":
                found = first_source(b["id"])
                if found:
                    return found
    print(f"找不到微藍的年份資料庫（狀態 {status}）。401＝金鑰不正確；"
          "404＝頁面未連接「AIA 自動更新」或 NOTION_DATABASE_ID 不正確。")
    raise SystemExit(1)


def query_pages(source):
    r = requests.post(f"{API}/data_sources/{source}/query", headers=HEADERS, json={"page_size": 100}, timeout=30)
    return check(r, "讀取年份資料庫").get("results", [])


def page_title(page):
    for prop in page.get("properties", {}).values():
        if prop.get("type") == "title":
            return "".join(t.get("plain_text", "") for t in prop.get("title", []))
    return ""


def file_blocks(page_id):
    r = requests.get(f"{API}/blocks/{page_id}/children?page_size=100", headers=HEADERS, timeout=30)
    out = []
    for block in check(r, "讀取頁面內容").get("results", []):
        if block["type"] == "file":
            out.append((block["id"], block["file"].get("name", "")))
    return out


def upload_file(page_id, filename, content):
    old_auto = [bid for bid, name in file_blocks(page_id) if AUTO in name]
    r = requests.post(f"{API}/file_uploads", headers=HEADERS,
                      json={"mode": "single_part", "filename": filename, "content_type": XLSX}, timeout=30)
    upload_id = check(r, "建立上傳")["id"]
    r = requests.post(f"{API}/file_uploads/{upload_id}/send",
                      headers={"Authorization": f"Bearer {TOKEN}", "Notion-Version": VERSION},
                      files={"file": (filename, content, XLSX)}, timeout=120)
    check(r, "傳送檔案")
    block = {"object": "block", "type": "file",
             "file": {"type": "file_upload", "file_upload": {"id": upload_id}}}
    r = requests.patch(f"{API}/blocks/{page_id}/children", headers=HEADERS, json={"children": [block]}, timeout=60)
    check(r, "加入檔案到頁面")
    for bid in old_auto:
        r = requests.delete(f"{API}/blocks/{bid}", headers=HEADERS, timeout=30)
        check(r, "刪除上一次的自動生成檔案")
    print(f"已上傳：{filename}（刪除舊的自動生成檔案 {len(old_auto)} 個）")


# ---------- 主流程 ----------
def main():
    print("=== 1. 讀取 repo 內的 PDF ===")
    paths = sorted(glob.glob("**/*.pdf", recursive=True))
    print(f"   PDF {len(paths)} 份")
    if not paths:
        raise SystemExit("✗ repo 內沒有 PDF，請先上傳微藍的分紅實現率 PDF")
    docs = [parse_pdf(p) for p in paths]
    for d in docs:
        print(f"   {d['file']}")
        print(f"      產品：{d['name']}｜報告年度 {d['year']}｜{d['rtype']}｜PDF 建立 {d['created']}")
        print("      " + " ".join(f"{k}:{fmt_pct(v)}" for k, v in zip(d["labels"], d["vals"])))

    print()
    print("=== 2. 核對 ===")
    years = sorted({d["year"] for d in docs})
    if len(years) != 1:
        raise SystemExit(f"✗ PDF 的報告年度不一致：{years}。請刪除 repo 內舊年度的 PDF")
    if any(d["labels"] != docs[0]["labels"] or d["spans"] != docs[0]["spans"] for d in docs):
        raise SystemExit("✗ 各 PDF 的保單年度欄不一致，暫不更新")
    names = [d["name"] for d in docs]
    if len(set(names)) != len(names):
        raise SystemExit(f"✗ 有重複的產品名稱：{names}")
    year, labels, spans = years[0], docs[0]["labels"], docs[0]["spans"]
    print(f"   報告年度 {year}｜保單年度欄：" + " ".join(f"{k}({s})" for k, s in zip(labels, spans)))

    rows = [[d["name"], category(d["name"]), "未標示", d["rtype"]] + d["vals"] for d in docs]
    rows.sort(key=lambda x: x[0])
    nums = numbers(rows, 4)
    slashes = sum(1 for r0 in rows for v in r0[4:] if not is_num(v))

    print()
    print("=== 3. 讀取結果 ===")
    print(f"   產品 {len(rows)} 個｜{len(rows)} 行｜數值 {len(nums)} 個｜「/」{slashes} 格"
          f"｜平均 {fmt_pct(mean_or_slash(nums))}"
          f"｜最高 {fmt_pct(max(nums)) if nums else '/'}｜最低 {fmt_pct(min(nums)) if nums else '/'}")
    types = collections.Counter(r0[3] for r0 in rows)
    print("   紅利類型：", "｜".join(f"{k} {v}" for k, v in types.items()))
    print("   產品種類（按名稱推斷）：", dict(collections.Counter(r0[1] for r0 in rows)))

    groups = [("所有分紅", rows)]
    for t in types:
        sub = [r0 for r0 in rows if r0[3] == t]
        if numbers(sub, 4):
            groups.append((t, sub))
    summary = build_summary(4, labels, spans, groups)

    remarks = ["貨幣一欄顯示「未標示」，因為官方 PDF 沒有列明貨幣。",
               "產品種類是按產品名稱推斷，並非官方 PDF 的欄位。",
               "只包括微藍自己的產品，不包括前英傑華產品。"]
    if not nums:
        remarks.insert(0, f"報告年度 {year} 的所有格都是 N/A，所以平均值全部顯示「/」。")

    print()
    print("=== 4. Excel ===")
    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")
    notes = ("分紅實現率＝實際派發的非保證紅利 ÷ 銷售時所示的金額。"
             "100% 即與銷售時一致，高於 100% 即比銷售時多。\n"
             "綠色＝100% 或以上；紅色＝低於 80%；「/」＝官方 PDF 顯示 N/A。\n"
             f"資料來源：Blue 微藍網站的分紅實現率 PDF（報告年度 {year}）｜更新日期：{today}")
    header = ["产品名称", "产品种类", "货币", "红利类型"] + [f"保单年度{k}\n{s}" for k, s in zip(labels, spans)]
    title = f"{year}分紅實現率"
    wb = Workbook()
    wb.remove(wb.active)
    add_sheet(wb, title, f"微藍分紅實現率（報告年度 {year}）", notes, header, rows, summary,
              "分紅類型", 4, "2E75B6")
    add_overview(wb, year, today, [("分紅實現率", title, rows, 4)], remarks)
    buf = io.BytesIO()
    wb.save(buf)
    content = buf.getvalue()
    path = f"blue_fulfillment_{year}.xlsx"
    with open(path, "wb") as f:
        f.write(content)
    print(f"   {path}｜{len(content) // 1024} KB｜工作表：", "、".join(wb.sheetnames))

    print()
    print("=== 5. 上傳到 Notion ===")
    if os.environ.get("UPLOAD") != "1":
        print("   試跑：沒有上傳")
        print("完成")
        return
    target = target_id()
    if not TOKEN or not target:
        raise SystemExit("未設定 NOTION_TOKEN 或 NOTION_DATABASE_ID")
    source = resolve_source(target)
    page = next((p for p in query_pages(source) if year in page_title(p)), None)
    if page is None:
        body = {"parent": {"type": "data_source_id", "data_source_id": source},
                "properties": {"title": {"title": [{"text": {"content": year}}]}}}
        page = check(requests.post(f"{API}/pages", headers=HEADERS, json=body, timeout=30), "建立年度頁面")
        print(f"已建立頁面：{year}")
    else:
        print("使用現有頁面：", page_title(page))
    upload_file(page["id"], f"微藍分紅實現率（{year}年）{AUTO}.xlsx", content)
    print("完成")


if __name__ == "__main__":
    main()
