import html as htmllib
import json
import re
import subprocess
import sys

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi"])
from curl_cffi import requests as creq

BASE = "https://www.fwd.com.hk"
PATH = "/regulatory-disclosures/fulfillment-ratios/"


def get(url):
    r = creq.get(url, impersonate="chrome", timeout=60)
    return r.status_code, r.content.decode("utf-8", "replace")


def clean(s):
    s = re.sub(r"<[^>]+>", " ", str(s or ""))
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def is_cells(o):
    return (isinstance(o, list) and len(o) > 0
            and all(isinstance(x, dict) and "content" in x for x in o)
            and any("row_span" in x or "column_span" in x for x in o))


def walk(o, path, out):
    if isinstance(o, dict):
        for k, v in o.items():
            walk(v, path + (k,), out)
    elif isinstance(o, list):
        if is_cells(o):
            out.append(path)
            return
        for i, v in enumerate(o):
            walk(v, path + (i,), out)


def resolve(root, path):
    for k in path:
        root = root[k]
    return root


def load(slug):
    st, raw = get(BASE + "/zh" + PATH + slug + "/")
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', raw, flags=re.S)
    if st != 200 or not m:
        return st, None, {}
    root = json.loads(m.group(1))
    out = []
    walk(root, (), out)
    tables = {}
    for p in out:
        idx = [i for i, x in enumerate(p) if isinstance(x, int)]
        tid = p[:idx[-1]] if idx else p
        tables.setdefault(tid, []).append(p)
    return st, root, tables


def cell_text(c):
    t = clean(c.get("content"))[:18] or "∅"
    rs, cs = c.get("row_span"), c.get("column_span")
    if rs not in (None, 1, "1"):
        t += f"[r{rs}]"
    if cs not in (None, 1, "1"):
        t += f"[c{cs}]"
    return t


def show(slug):
    st, root, tables = load(slug)
    print(f"=== 樣本：{slug}｜狀態 {st}｜表格 {len(tables)} 個 ===")
    for n, (tid, rows) in enumerate(list(tables.items())[:4], 1):
        print(f"-- 表 {n}｜{len(rows)} 行｜路徑尾 {'/'.join(str(x) for x in tid[-5:])}")
        up, shown = tid[:-1], 0
        while shown < 3:
            try:
                d = resolve(root, up)
            except Exception:
                break
            if isinstance(d, dict):
                vals = [f"{k}={clean(v)[:40]}" for k, v in d.items()
                        if isinstance(v, str) and 0 < len(clean(v)) <= 60][:8]
                print(f"   上層{shown + 1} 鍵 {list(d.keys())[:10]}")
                if vals:
                    print("   文字：" + "｜".join(vals))
                shown += 1
            if not up:
                break
            up = up[:-1]
        for p in rows[:25]:
            cells = resolve(root, p)
            print(f"   ({len(cells)}格) " + " | ".join(cell_text(c) for c in cells[:16]))
        if len(rows) > 25:
            print(f"   …其餘 {len(rows) - 25} 行略")


status, body = get(BASE + "/zh" + PATH)
slugs = list(dict.fromkeys(re.findall(re.escape(PATH) + r"([A-Za-z0-9_\-]+)", body)))
print("=== 全部產品頁（表格數/行數/百分比格數）===")
stats, line = [], []
for s in slugs:
    try:
        st, root, tables = load(s)
        nrow = sum(len(r) for r in tables.values())
        npct = sum(1 for rows in tables.values() for p in rows for c in resolve(root, p)
                   if re.search(r"\d\s*%", clean(c.get("content"))))
        stats.append((s, len(tables), nrow, npct))
        line.append(f"{s[:28]} {len(tables)}/{nrow}/{npct}")
    except Exception as e:
        stats.append((s, 0, 0, 0))
        line.append(f"{s[:28]} 失敗:{str(e)[:30]}")
    if len(line) == 3:
        print("  " + "｜".join(line))
        line = []
if line:
    print("  " + "｜".join(line))
print(f"合計：{len(stats)} 頁｜表格 {sum(x[1] for x in stats)}｜行 {sum(x[2] for x in stats)}"
      f"｜百分比格 {sum(x[3] for x in stats)}｜沒有表格的頁 {sum(1 for x in stats if x[1] == 0)}")

picks = []
if stats:
    for s in (stats[0][0], max(stats, key=lambda x: x[2])[0]):
        if s not in picks:
            picks.append(s)
for s in picks:
    show(s)
print("=== 完 ===")
