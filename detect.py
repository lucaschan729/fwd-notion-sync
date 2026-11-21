import html as htmllib
import json
import re
import subprocess
import sys

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi"])
from curl_cffi import requests as creq

BASE = "https://www.fwd.com.hk"
PATH = "/regulatory-disclosures/fulfillment-ratios/"
WORDS = ["分紅實現率", "總價值比率", "總現金價值", "週年紅利", "特別紅利", "歸原紅利", "保單年度", "不適用", "美元", "港元"]


def get(url):
    r = creq.get(url, impersonate="chrome", timeout=60)
    return r.status_code, r.content.decode("utf-8", "replace")


def unesc(s):
    for a, b in (('\\"', '"'), ("\\u003c", "<"), ("\\u003e", ">"), ("\\u0026", "&"), ("\\n", " "), ("\\/", "/")):
        s = s.replace(a, b)
    return s


def clean(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def flat(s):
    return re.sub(r"\s+", " ", s)


status, body = get(BASE + "/zh" + PATH)
slugs = list(dict.fromkeys(re.findall(re.escape(PATH) + r"([A-Za-z0-9_\-]+)", body)))
print("=== 產品頁 ===")
print(f"總頁狀態 {status}｜產品頁 {len(slugs)} 個")
for i in range(0, len(slugs), 3):
    print("  " + "｜".join(slugs[i:i + 3]))

picks = []
for s in ([slugs[0], slugs[len(slugs) // 2], slugs[-1]] if slugs else []):
    if s not in picks:
        picks.append(s)

for n, slug in enumerate(picks, 1):
    print(f"=== 樣本 {n}：{slug} ===")
    url = BASE + "/zh" + PATH + slug + "/"
    st, raw = get(url)
    if st != 200:
        print(f"  /zh 狀態 {st}，改試沒有 /zh")
        st, raw = get(BASE + PATH + slug + "/")
    b = unesc(raw)
    title = re.search(r"<title[^>]*>(.*?)</title>", raw, flags=re.S | re.I)
    print(f"狀態 {st}｜長度 {len(raw)}｜標題 {clean(title.group(1))[:50] if title else '沒有'}")
    t1 = re.findall(r"<table\b.*?</table>", raw, flags=re.S | re.I)
    t2 = re.findall(r"<table\b.*?</table>", b, flags=re.S | re.I)
    print(f"表格：原文 {len(t1)} 個｜還原後 {len(t2)} 個")
    print("關鍵字：", "｜".join(f"{w} {b.count(w)}" for w in WORDS))
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', raw, flags=re.S)
    if m:
        try:
            d = json.loads(m.group(1))
            pp = d.get("props", {}).get("pageProps", {})
            print(f"NEXT_DATA：page {d.get('page')}｜pageProps 鍵 {list(pp.keys())[:12]}")
        except Exception as e:
            print(f"NEXT_DATA 解析失敗：{e}")
    else:
        print("NEXT_DATA：沒有")
    hits = list(re.finditer(r"\d{2,3}(?:\.\d+)?\s*%", b))
    print(f"百分比數值 {len(hits)} 個")
    for t in (t2 or t1)[:2]:
        rows = re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)
        print(f"  表格 {len(rows)} 行：")
        for tr in rows[:6]:
            cells = [clean(c)[:14] for c in re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=re.S | re.I)]
            print("     " + " | ".join(cells[:14]))
    if hits:
        p = hits[0].start()
        print("  第一個百分比附近原文：")
        print("  " + flat(b[max(0, p - 450):p + 250]))
        if len(hits) > 1:
            p = hits[-1].start()
            print("  最後一個百分比附近原文：")
            print("  " + flat(b[max(0, p - 250):p + 150]))
    else:
        p = max(b.rfind("週年紅利"), b.rfind("特別紅利"))
        print("  沒有百分比，紅利字眼附近原文：")
        print("  " + (flat(b[max(0, p - 350):p + 350]) if p >= 0 else "找不到"))
print("=== 完 ===")
