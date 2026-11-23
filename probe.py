import re

import requests

PAGE = "https://www.blue.com.hk/hk/participating_policies"
PDF = ("https://www.blue.com.hk/static/a266da73ede994e5579c7cf8ac83a50f/"
       "WeWeath_GoWealth_Generator_-_Fulfillment_Ratio_2025_Clean.pdf")
H = {"User-Agent": "Mozilla/5.0 (blue-notion-sync yearly check)"}


def show(name, fn):
    try:
        r = fn()
    except Exception as e:
        print("   %s｜✗ 連線失敗：%s" % (name, str(e)[:200]))
        return None
    print("   %s｜HTTP %s｜%d KB｜最終網址 %s" % (name, r.status_code, len(r.content) // 1024, r.url[:150]))
    return r


def test(label, url):
    print("\n=== %s ===" % label)
    print("   " + url)
    r = show("現有存檔", lambda: requests.get("https://archive.org/wayback/available",
                                          params={"url": url}, headers=H, timeout=60))
    if r is not None and r.status_code == 200:
        print("   " + r.text[:300])
    show("要求重新存檔", lambda: requests.get("https://web.archive.org/save/" + url, headers=H, timeout=180))
    r = show("讀取最新存檔", lambda: requests.get("https://web.archive.org/web/2id_/" + url, headers=H, timeout=120))
    return r


def main():
    r = test("1. 紅利資料頁", PAGE)
    if r is not None and r.status_code == 200:
        body = r.content.decode("utf-8", errors="replace")
        links = sorted(set(re.findall(r"""["'(]([^"'()\s<>]+\.pdf)""", body, flags=re.I)))
        print("   「實現率」出現 %d 次｜PDF 連結 %d 個" % (body.count("實現率"), len(links)))
        for u in links[:30]:
            print("   " + u)

    r = test("2. GO息寶 PDF", PDF)
    if r is not None:
        print("   是 PDF：%s" % r.content.startswith(b"%PDF"))

    print("\n=== 探查完成：沒有產生 Excel，沒有連接 Notion ===")


if __name__ == "__main__":
    main()
