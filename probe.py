import html as htmllib
import io
import re
import urllib.parse

import pdfplumber
import requests

BASE = "https://www.blue.com.hk"
PAGE = BASE + "/hk/participating_policies"
DATA = BASE + "/page-data/hk/participating_policies/page-data.json"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
     "Accept-Language": "zh-HK,zh;q=0.9"}
KEYS = ["分紅實現率", "實現率", "總現金價值", "現金價值比率", "報告年度", "終期紅利", "不適用", "派息率"]
WANT = re.compile(r"fulfil|ratio|bonus|dividend|participat|紅利|分紅|實現", re.I)


def plain(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", "", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def fetch(url):
    try:
        r = requests.get(url, headers=H, timeout=60)
    except Exception as e:
        print("   ✗ 讀取失敗｜%s｜%s" % (url, e))
        return ""
    r.encoding = "utf-8"
    print("   HTTP %s｜%s｜%d 字元｜%s" % (r.status_code, r.headers.get("Content-Type", ""), len(r.text), url))
    return r.text if r.status_code == 200 else ""


def keywords(body):
    text = plain(body.replace("\\n", " "))
    print("   關鍵字次數：" + "｜".join("%s %d" % (k, text.count(k)) for k in KEYS))
    for k in ("實現率", "總現金價值", "報告年度"):
        for m in list(re.finditer(k, text))[:3]:
            print("   …%s…" % text[max(0, m.start() - 60):m.end() + 100])


def pdf_links(body):
    found = re.findall(r"""["'(]([^"'()\s<>]+\.pdf)""", body, flags=re.I)
    out = []
    for u in found:
        u = urllib.parse.urljoin(PAGE, u.replace("\\u002F", "/").replace("\\/", "/"))
        if u not in out:
            out.append(u)
    return out


def dump_pdf(url):
    try:
        r = requests.get(url, headers=H, timeout=120)
    except Exception as e:
        print("   ✗ 下載失敗：%s" % e)
        return
    print("   HTTP %s｜%s｜%d KB｜Last-Modified：%s"
          % (r.status_code, r.headers.get("Content-Type", ""), len(r.content) // 1024, r.headers.get("Last-Modified")))
    if r.status_code != 200 or not r.content.startswith(b"%PDF"):
        return
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        print("   頁數：%d｜資料：%s" % (len(pdf.pages), pdf.metadata))
        for i, page in enumerate(pdf.pages[:8]):
            text = page.extract_text() or ""
            tables = page.extract_tables()
            print("\n--- 第 %d 頁｜文字 %d 字元｜表格 %d 個 ---" % (i + 1, len(text), len(tables)))
            for line in text.splitlines()[:60]:
                print(line)
            for j, t in enumerate(tables):
                print("   [表格 %d] %d 行 × %d 欄" % (j + 1, len(t), max(len(x) for x in t) if t else 0))
                for row in t[:12]:
                    print("      " + "｜".join("∅" if c is None else re.sub(r"\s+", " ", c) for c in row))
                if len(t) > 12:
                    print("      …（另有 %d 行）" % (len(t) - 12))


def main():
    print("=== 1. 紅利資料頁 ===")
    body = fetch(PAGE)
    keywords(body)

    print("\n=== 2. 頁面資料檔（如有） ===")
    data = fetch(DATA)
    if data:
        keywords(data)

    print("\n=== 3. 相關連結 ===")
    n = 0
    for href, label in re.findall(r'<a\b[^>]*href="([^"]*)"[^>]*>(.*?)</a>', body, flags=re.S | re.I):
        label = plain(label)
        if WANT.search(href) or WANT.search(label):
            n += 1
            print("   %s｜%s" % (label[:40], urllib.parse.urljoin(PAGE, href)))
    for src in re.findall(r'<iframe\b[^>]*src="([^"]*)"', body, flags=re.I):
        print("   iframe｜%s" % src)
    print("   相關連結 %d 個" % n)

    print("\n=== 4. 網頁表格 ===")
    tables = re.findall(r"<table\b.*?</table>", body, flags=re.S | re.I)
    print("   表格 %d 個" % len(tables))
    last = 0
    for i, t in enumerate(tables[:15]):
        pos = body.find(t, last)
        print("\n   [表格 %d] 之前的文字：%s" % (i + 1, plain(body[max(last, pos - 3000):pos])[-200:]))
        last = pos + len(t)
        trs = re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)
        for tr in trs[:12]:
            cells = [plain(c) for c in re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=re.S | re.I)]
            print("      " + "｜".join(cells))
        if len(trs) > 12:
            print("      …（另有 %d 行）" % (len(trs) - 12))

    print("\n=== 5. PDF ===")
    pdfs = pdf_links(body + data)
    pdfs.sort(key=lambda u: 0 if WANT.search(urllib.parse.unquote(u)) else 1)
    print("   PDF 連結 %d 個" % len(pdfs))
    for u in pdfs:
        print("   " + u)
    for u in pdfs[:4]:
        print("\n>>> " + u)
        dump_pdf(u)

    print("\n=== 探查完成：沒有產生 Excel，沒有連接 Notion ===")


if __name__ == "__main__":
    main()
