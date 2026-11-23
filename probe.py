import html as htmllib
import io
import re
import urllib.parse

import pdfplumber
import requests
from curl_cffi import requests as creq

S = "https://www.blue.com.hk/static/"
PDFS = [
    S + "523f1b70d7b48e55f0b291b785850f54/WeWealth_GoWealth_Generator_3_-_Fulfillment_Ratio_2025_Clean.pdf",
    S + "9d3c89483c732790a90759cd58acb625/WeProtect_GoWealth_Plus_-_Fulfillment_Ratio_2025_Clean.pdf",
    S + "a75513c6d6a7c17777649070922daf49/WeWealth_GoWealth_Generator_2_-_Fulfillment_Ratio_2025_Clean.pdf",
    S + "a266da73ede994e5579c7cf8ac83a50f/WeWeath_GoWealth_Generator_-_Fulfillment_Ratio_2025_Clean.pdf",
]
OLD = "https://www.blue.com.hk/other_products/hk/participating_policies/fulfillment_ratio/"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
H = {"User-Agent": UA, "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
LANG = {"Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
KEYS = ["分紅實現率", "實現率", "總現金價值", "報告年度", "終期紅利", "週年紅利", "不適用", "N/A"]
METHODS = [
    ("普通連線", lambda u: requests.get(u, headers=H, timeout=60)),
    ("模擬 Chrome", lambda u: creq.get(u, impersonate="chrome", headers=LANG, timeout=60)),
]


def plain(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", "", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def get(url):
    print("   網址：" + url)
    for name, fn in METHODS:
        try:
            r = fn(url)
        except Exception as e:
            print("   %s｜✗ 連線失敗：%s" % (name, str(e)[:150]))
            continue
        print("   %s｜HTTP %s｜%d KB｜server %s｜節點 %s"
              % (name, r.status_code, len(r.content) // 1024,
                 r.headers.get("server"), r.headers.get("x-volterra-location")))
        if r.status_code == 200 and r.content:
            return r
    return None


def dump_pdf(content):
    if not content.startswith(b"%PDF"):
        print("   ✗ 不是 PDF")
        return
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        print("   頁數：%d｜資料：%s" % (len(pdf.pages), pdf.metadata))
        for i, page in enumerate(pdf.pages[:6]):
            text = page.extract_text() or ""
            tables = page.extract_tables()
            print("\n--- 第 %d 頁｜文字 %d 字元｜表格 %d 個 ---" % (i + 1, len(text), len(tables)))
            for line in text.splitlines()[:45]:
                print(line)
            for j, t in enumerate(tables):
                print("   [表格 %d] %d 行 × %d 欄" % (j + 1, len(t), max(len(x) for x in t) if t else 0))
                for row in t[:14]:
                    print("      " + "｜".join("∅" if c is None else re.sub(r"\s+", " ", c) for c in row))


def main():
    print("=== 1. GitHub 伺服器所在地 ===")
    try:
        info = requests.get("https://ipinfo.io/json", timeout=30).json()
        print("   %s｜%s｜%s" % (info.get("country"), info.get("city"), info.get("org")))
    except Exception as e:
        print("   ✗ 查不到：%s" % e)

    print("\n=== 2. 微藍 4 份分紅實現率 PDF ===")
    ok = 0
    for u in PDFS:
        print()
        r = get(u)
        if r is not None:
            ok += 1
            dump_pdf(r.content)
    print("\n   讀到 %d／%d 份" % (ok, len(PDFS)))

    print("\n=== 3. 前英傑華產品分紅實現率頁 ===")
    r = get(OLD)
    if r is not None:
        body = r.content.decode("utf-8", errors="replace")
        text = plain(body)
        print("   關鍵字次數：" + "｜".join("%s %d" % (k, text.count(k)) for k in KEYS))
        print("   內容開頭：" + text[:1500])
        tables = re.findall(r"<table\b.*?</table>", body, flags=re.S | re.I)
        print("   表格 %d 個" % len(tables))
        for i, t in enumerate(tables[:10]):
            print("\n   [表格 %d]" % (i + 1))
            trs = re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)
            for tr in trs[:14]:
                cells = [plain(c) for c in re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=re.S | re.I)]
                print("      " + "｜".join(cells))
            if len(trs) > 14:
                print("      …（另有 %d 行）" % (len(trs) - 14))
        links = []
        for u in re.findall(r"""["'(]([^"'()\s<>]+\.pdf)""", body, flags=re.I):
            u = urllib.parse.urljoin(OLD, u)
            if u not in links:
                links.append(u)
        print("\n   PDF 連結 %d 個" % len(links))
        for u in links:
            print("   " + u)
        for u in links[:3]:
            print("\n>>> PDF")
            p = get(u)
            if p is not None:
                dump_pdf(p.content)

    print("\n=== 探查完成：沒有產生 Excel，沒有連接 Notion ===")


if __name__ == "__main__":
    main()
