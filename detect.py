import html as htmllib
import re
import subprocess
import sys

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi"])
from curl_cffi import requests as creq

URL = "https://www.fwd.com.hk/zh/regulatory-disclosures/fulfillment-ratios/"
WORDS = ["分紅實現率", "總價值比率", "總現金價值", "週年紅利", "特別紅利", "終期紅利", "保單年度", "不適用", "報告年度"]


def clean(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


r = creq.get(URL, impersonate="chrome", timeout=60)
body = r.content.decode("utf-8", "replace")
print("=== 總頁 ===")
print(f"狀態 {r.status_code}｜長度 {len(body)}｜最終網址 {r.url}")
tables = re.findall(r"<table\b.*?</table>", body, flags=re.S | re.I)
print(f"表格 {len(tables)} 個｜__NEXT_DATA__ {'有' if '__NEXT_DATA__' in body else '沒有'}"
      f"｜iframe {len(re.findall(r'<iframe', body, flags=re.I))} 個")
print("關鍵字次數：", "｜".join(f"{w} {body.count(w)}" for w in WORDS))
heads = [clean(h) for h in re.findall(r"<h[1-4]\b[^>]*>(.*?)</h[1-4]>", body, flags=re.S | re.I)]
heads = [h for h in dict.fromkeys(heads) if h]
print(f"=== 標題 {len(heads)} 個（頭 25 個）===")
for h in heads[:25]:
    print("  " + h[:60])
links = list(dict.fromkeys(re.findall(r"[\"'(]((?:https?:)?/[^\"'()\s<>]+)", body)))
keep = [u for u in links if re.search(r"\.pdf|\.json|\.xlsx?|\.csv|fulfil|ratio", u, flags=re.I)
        and not re.search(r"\.(js|css|png|jpg|svg|woff2?)(\?|$)", u, flags=re.I)]
print(f"=== 相關連結 {len(keep)} 個（頭 25 個）===")
for u in keep[:25]:
    print("  " + u[-110:])
print("=== 頭 4 個表格 ===")
for i, t in enumerate(tables[:4], 1):
    rows = re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)
    print(f"表 {i}：{len(rows)} 行")
    for tr in rows[:5]:
        cells = [clean(c)[:16] for c in re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=re.S | re.I)]
        print("   " + " | ".join(cells[:14]))
if not tables:
    text = clean(body)
    p = text.find("分紅實現率")
    print("=== 沒有表格，內文片段 ===")
    print(text[max(0, p):p + 600] if p >= 0 else text[:600])
print("=== 完 ===")
