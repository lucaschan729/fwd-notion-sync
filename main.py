import datetime
import html as htmllib
import json
import os
import re
import sys
import time
from collections import Counter

import requests
from curl_cffi import requests as creq
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

BASE = "https://www.fwd.com.hk"
PATH = "/regulatory-disclosures/fulfillment-ratios/"
LABELS = [str(i) for i in range(1, 11)] + ["10+"]
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ASCII_NAME = "FWD_fulfillment_auto.xlsx"
NAVY = "1F4E78"
TYPE_COLORS = {"週年紅利": "1F4E78", "特別紅利": "C55A11", "歸原紅利": "548235"}


def get(url):
    err = ""
    for _ in range(3):
        try:
            r = creq.get(url, impersonate="chrome", timeout=60)
            if r.status_code == 200:
                return r.content.decode("utf-8", "replace")
            err = f"狀態 {r.status_code}"
        except Exception as e:
            err = str(e)[:80]
        time.sleep(3)
    raise RuntimeError(err)


def clean(s):
    s = re.sub(r"<[^>]+>", " ", str(s or ""))
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def is_cells(o):
    return (isinstance(o, list) and len(o) > 0
            and all(isinstance(x, dict) and "content" in x for x in o)
            and any("row_span" in x or "column_span" in x for x in o))


def find_rows(o, out):
    if isinstance(o, dict):
        for v in o.values():
            find_rows(v, out)
    elif isinstance(o, list):
        if is_cells(o):
            out.append(o)
            return
        for v in o:
            find_rows(v, out)


def find_sections(o, out):
    if isinstance(o, dict):
        if "table" in o and "description" in o:
            out.append(o)
            return
        for v in o.values():
            find_sections(v, out)
    elif isinstance(o, list):
        for v in o:
            find_sections(v, out)


def parse_desc(text):
    m = re.search(r"^(.*?)\s*產品類別\s*[－\-–—―:：]?\s*(.*?)\s*(\d{4})\s*年度", text)
    if m:
        return m.group(1).strip(), m.group(2).strip(), m.group(3)
    m = re.search(r"^(.*?)\s*(\d{4})\s*年度", text)
    if m:
        return m.group(1).strip(), "", m.group(2)
    return text.strip(), "", ""


def to_val(s):
    m = re.search(r"(\d+(?:\.\d+)?)\s*%", s)
    return round(float(m.group(1)) / 100, 6) if m else "/"


def name_key(name):
    return re.sub(r"\s+", "", name)


def parse_page(slug, raw, years, seen, dups):
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', raw, flags=re.S)
    if not m:
        raise RuntimeError("沒有 NEXT_DATA")
    secs, recs = [], []
    find_sections(json.loads(m.group(1)), secs)
    for sec in secs:
        rows = []
        find_rows(sec.get("table"), rows)
        if not rows:
            continue
        texts = [[clean(c.get("content")) for c in r] for r in rows]
        name, cat, year = parse_desc(clean(sec.get("description")))
        name = name or clean(sec.get("title")) or slug
        sig = json.dumps([name_key(name), texts], ensure_ascii=False)
        if sig in seen:
            if seen[sig] != slug:
                dups.append(f"{name}（{slug} 與 {seen[sig]} 相同）")
            continue
        seen[sig] = slug
        head = next((t for t in texts if sum(1 for x in t if re.search(r"年度\s*\d", x)) >= 3), None)
        colmap, cur_idx = {}, None
        if head:
            for i, x in enumerate(head):
                hm = re.search(r"年度\s*(\d+)\s*(\+|以上)?", x)
                if hm:
                    lab = hm.group(1) + ("+" if hm.group(2) else "")
                    if lab in colmap.values():
                        lab += "+"
                    colmap[i] = lab
                    ym = re.search(r"[（(]\s*(\d{4}(?:\s*[-–—至]\s*\d{4})?)", x)
                    if ym and lab not in years:
                        years[lab] = re.sub(r"\s+", "", ym.group(1))
                elif "貨幣" in x:
                    cur_idx = i
        else:
            n = max(len(t) for t in texts)
            start = n - len(LABELS)
            cur_idx = 1 if start >= 2 else None
            colmap = {start + i: lab for i, lab in enumerate(LABELS)}
        last_type = ""
        for t in texts:
            if t is head or t == head or not any(t):
                continue
            if head and len(t) == len(head) - 1 and last_type:
                t = [""] + t
            typ = t[0] or last_type
            last_type = typ
            cur = t[cur_idx] if cur_idx is not None and cur_idx < len(t) else ""
            if "不適用" in cur or cur in ("-", "/", "—", "–"):
                cur = ""
            vals = {lab: to_val(t[i]) for i, lab in colmap.items() if i < len(t)}
            recs.append({"slug": slug, "name": name, "cat": cat, "year": year,
                         "typ": typ, "cur": cur, "vals": vals})
    return recs


def group_of(typ):
    for k in ("週年紅利", "特別紅利", "歸原紅利"):
        if k[:2] in typ:
            return k
    return typ or "其他"


def avg(rs, lab):
    v = [r["vals"][lab] for r in rs if isinstance(r["vals"].get(lab), float)]
    return (sum(v) / len(v), len(v)) if v else (None, 0)


def suffix_avg(rs, i):
    v = [r["vals"][lab] for r in rs for lab in LABELS[i:] if isinstance(r["vals"].get(lab), float)]
    return sum(v) / len(v) if v else None


def products(recs):
    out = {}
    for x in recs:
        p = out.setdefault(name_key(x["name"]), {"name": x["name"], "cat": x["cat"], "rows": 0, "vals": []})
        p["rows"] += 1
        if not p["cat"] and x["cat"]:
            p["cat"] = x["cat"]
        p["vals"] += [v for v in x["vals"].values() if isinstance(v, float)]
    return list(out.values())


def build_excel(recs, years, year, path):
    wb = Workbook()
    ov = wb.active
    ov.title = "概覽"
    sheet_name = f"{year}分紅實現率"
    ws = wb.create_sheet(sheet_name)

    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    sep_border = Border(bottom=Side(style="thin", color="5B9BD5"))
    sep_fill = PatternFill("solid", fgColor="DCE6F1")
    head_fill = PatternFill("solid", fgColor=NAVY)
    green = PatternFill("solid", fgColor="E2F0D9")
    red = PatternFill("solid", fgColor="FCE4E4")
    white = PatternFill("solid", fgColor="FFFFFF")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")
    source = f"資料來源：富衛香港網站（報告年度 {year}）｜更新日期：{today}"
    prods = products(recs)
    nums = [v for rec in recs for v in rec["vals"].values() if isinstance(v, float)]

    def pct_fmt(v):
        return "0%" if abs(v * 100 - round(v * 100)) < 1e-6 else "0.0%"

    def head_row(sh, row, heads, start=1, height=34):
        for c, h in enumerate(heads, start):
            cell = sh.cell(row=row, column=c, value=h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = head_fill
            cell.alignment = center
            cell.border = border
        sh.row_dimensions[row].height = height

    def ratio_cell(cell, v, fmt=None):
        cell.alignment = center
        cell.border = border
        cell.fill = white
        if isinstance(v, float):
            cell.value = v
            cell.number_format = fmt or pct_fmt(v)
            if v >= 1:
                cell.fill = green
                cell.font = Font(color="375623")
            elif v < 0.8:
                cell.fill = red
                cell.font = Font(color="C00000")
            else:
                cell.font = Font(color="262626")
        else:
            cell.value = "/"
            cell.font = Font(color="BFBFBF")

    def text_row(sh, row, text, ncols, font, height=20, start=1):
        sh.merge_cells(start_row=row, start_column=start, end_row=row, end_column=ncols)
        cell = sh.cell(row=row, column=start, value=text)
        cell.font = font
        cell.alignment = left
        sh.row_dimensions[row].height = height

    def setup_print(sh, rows):
        sh.page_setup.orientation = "landscape"
        sh.page_setup.fitToWidth = 1
        sh.page_setup.fitToHeight = 0
        sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        sh.print_title_rows = rows

    # ---------- 分紅實現率 ----------
    ncol = 4 + len(LABELS)
    last = get_column_letter(ncol)
    grey = Font(size=10, color="595959")
    text_row(ws, 1, f"富衛分紅實現率（報告年度 {year}）", ncol, Font(bold=True, size=14, color=NAVY), 28)
    text_row(ws, 2, "分紅實現率＝實際派發的非保證紅利 ÷ 銷售時說明的金額。100% 即與預期一致，高於 100% 即比預期多。定義以富衛網站的說明為準。", ncol, grey)
    text_row(ws, 3, "綠色＝100% 或以上；紅色＝低於 80%；「/」＝網站沒有提供該年度的數字（不適用）。", ncol, grey)
    text_row(ws, 4, source, ncol, grey)
    head_row(ws, 5, ["產品名稱", "產品種類", "貨幣", "紅利類型"] +
             [f"保單年度{lab}\n{years.get(lab, '')}".strip() for lab in LABELS])
    r, prev = 6, None
    for rec in recs:
        key = name_key(rec["name"])
        if prev is not None and key != prev:
            for c in range(1, ncol + 1):
                cell = ws.cell(row=r, column=c)
                cell.fill = sep_fill
                cell.border = sep_border
            ws.row_dimensions[r].height = 5
            r += 1
        prev = key
        for c, v in enumerate([rec["name"], rec["cat"], rec["cur"], rec["typ"]], 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = border
            cell.fill = white
            cell.alignment = left if c == 1 else center
            if c == 1:
                cell.font = Font(bold=True, color="262626")
            elif c == 2:
                cell.font = Font(size=10, color="595959")
            elif c == 3:
                cell.font = Font(bold=True, color="7030A0")
            else:
                cell.font = Font(color=TYPE_COLORS.get(group_of(rec["typ"]), "404040"))
        for c, lab in enumerate(LABELS, 5):
            ratio_cell(ws.cell(row=r, column=c), rec["vals"].get(lab, "/"))
        ws.row_dimensions[r].height = 21
        r += 1
    ws.auto_filter.ref = f"A5:{last}{r - 1}"
    ws.freeze_panes = "B6"

    r += 1
    ws.cell(row=r, column=4, value="平均值摘要").font = Font(bold=True, size=12, color=NAVY)
    r += 1
    head_row(ws, r, ["分紅類型", "所有年份平均值"] + [f"持滿{i}年平均值" for i in range(2, 10)] +
             ["第10年平均值", "持滿10年+平均值"], start=4)
    types = list(dict.fromkeys(group_of(x["typ"]) for x in recs))
    groups = [("所有分紅", recs)] + [(g, [x for x in recs if group_of(x["typ"]) == g]) for g in types]
    for g, rs in groups:
        r += 1
        cell = ws.cell(row=r, column=4, value=g)
        cell.font = Font(bold=True, color=TYPE_COLORS.get(g, "262626"))
        cell.alignment = center
        cell.border = border
        for i in range(len(LABELS)):
            a = suffix_avg(rs, i)
            ratio_cell(ws.cell(row=r, column=5 + i), a if a is not None else "/", "0.00%")
        ws.row_dimensions[r].height = 21
    r += 1
    head_row(ws, r, ["當年平均值"] + [years.get(lab, lab) for lab in LABELS], start=4, height=24)
    r += 1
    cell = ws.cell(row=r, column=4, value="所有分紅")
    cell.font = Font(bold=True, color="262626")
    cell.alignment = center
    cell.border = border
    for i, lab in enumerate(LABELS):
        a, _ = avg(recs, lab)
        ratio_cell(ws.cell(row=r, column=5 + i), a if a is not None else "/", "0.00%")
    ws.row_dimensions[r].height = 21
    r += 1
    text_row(ws, r, "「持滿N年平均值」＝保單年度N及之後所有數字的平均；「當年平均值」＝該年度一欄所有數字的平均。平均值計入所有貨幣，為固定數字，不會隨篩選變動。",
             ncol, Font(size=9, color="7F7F7F"), 30, start=4)
    for c, w in enumerate([38, 18, 10, 12] + [12] * len(LABELS), 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    setup_print(ws, "5:5")

    # ---------- 概覽 ----------
    text_row(ov, 1, f"富衛分紅表現概覽（報告年度 {year}）", 6, Font(bold=True, size=14, color=NAVY), 28)
    text_row(ov, 2, f"資料來源：富衛香港網站｜更新日期：{today}", 6, grey)
    sec_font = Font(bold=True, size=12, color=NAVY)
    ov.cell(row=4, column=1, value="資料概況").font = sec_font
    head_row(ov, 5, ["工作表", "產品數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"], height=26)
    n = len(nums) or 1
    row = [sheet_name, len(prods), len(recs), len(nums), sum(nums) / n, sum(1 for v in nums if v < 0.8) / n]
    for c, v in enumerate(row, 1):
        cell = ov.cell(row=6, column=c, value=v)
        cell.border = border
        cell.alignment = center
        if c == 1:
            cell.hyperlink = f"#'{sheet_name}'!A1"
            cell.font = Font(bold=True, color="0563C1", underline="single")
        elif c == 5:
            cell.number_format = "0.00%"
        elif c == 6:
            cell.number_format = "0.0%"
    ov.row_dimensions[6].height = 22

    ov.cell(row=8, column=1, value="按產品種類的整體平均值").font = sec_font
    head_row(ov, 9, ["產品種類", "分紅實現率", "產品數"], height=26)
    cats = list(dict.fromkeys(p["cat"] or "未識別種類" for p in prods))
    rr = 10
    for g in cats:
        ps = [p for p in prods if (p["cat"] or "未識別種類") == g]
        v = [x for p in ps for x in p["vals"]]
        cell = ov.cell(row=rr, column=1, value=g)
        cell.font = Font(bold=True, color="262626")
        cell.border = border
        cell.alignment = left
        ratio_cell(ov.cell(row=rr, column=2), sum(v) / len(v) if v else "/", "0.00%")
        cell = ov.cell(row=rr, column=3, value=len(ps))
        cell.border = border
        cell.alignment = center
        ov.row_dimensions[rr].height = 22
        rr += 1

    rr += 1
    ov.cell(row=rr, column=1, value="顏色說明").font = sec_font
    for v, t in [(1.0, "100% 或以上：與銷售時的說明一致或更好"),
                 (0.79, "低於 80%：明顯低於銷售時的說明"),
                 ("/", "沒有數字：網站沒有提供該年度的數字（不適用）")]:
        rr += 1
        ratio_cell(ov.cell(row=rr, column=1), v, "0%")
        text_row(ov, rr, t, 6, Font(size=10, color="404040"), 22, start=2)
    rr += 2
    ov.cell(row=rr, column=1, value="使用提示").font = sec_font
    for t in ["・點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
              "・產品數按產品名稱計算，同名產品只算一次。",
              "・工作表第 5 行有篩選箭嘴，可按產品種類、貨幣或紅利類型篩選。",
              "・產品之間有淺藍色分隔行；排序會打亂分隔行，建議只用篩選。",
              "・向右或向下捲動時，標題列和產品名稱會固定不動。",
              "・平均值摘要在工作表最底部。",
              "・手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。",
              "・過往表現不代表未來。"]:
        rr += 1
        text_row(ov, rr, t, 6, Font(size=10, color="404040"))
    ov.column_dimensions["A"].width = 28
    for c in "BCDEF":
        ov.column_dimensions[c].width = 16
    ov.sheet_view.showGridLines = False

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)


def notion_upload(path, year):
    token = os.environ.get("NOTION_TOKEN", "").strip()
    raw = os.environ.get("NOTION_DATABASE_ID", "").strip()
    if not token or not raw:
        print("未設定 Notion Secrets，略過上傳（只產生 Excel）")
        return True
    seg = raw.split("?")[0].rstrip("/").split("/")[-1].replace("-", "")
    dbid = seg[-32:].lower()
    if not re.fullmatch(r"[0-9a-f]{32}", dbid):
        print("Notion 上傳失敗｜NOTION_DATABASE_ID 格式不對，請貼上資料庫整頁連結")
        return False
    H = {"Authorization": f"Bearer {token}", "Notion-Version": "2025-09-03"}

    def call(method, url, what, **kw):
        r = requests.request(method, "https://api.notion.com/v1" + url, headers=H, timeout=120, **kw)
        if r.status_code >= 300:
            msg = re.sub(r"[0-9a-f]{8}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{12}", "（已隱藏）", r.text[:300])
            raise RuntimeError(f"{what}：狀態 {r.status_code}｜{msg}")
        return r.json()

    try:
        try:
            db = call("GET", f"/databases/{dbid}", "讀取資料庫")
            ds = db["data_sources"][0]["id"]
        except RuntimeError as e:
            if "狀態 404" not in str(e):
                raise
            ds = dbid
        src = call("GET", f"/data_sources/{ds}", "讀取資料來源")
        tprop = next(k for k, v in src["properties"].items() if v.get("type") == "title")
        pages, cursor = {}, None
        while True:
            body = {"page_size": 100}
            if cursor:
                body["start_cursor"] = cursor
            res = call("POST", f"/data_sources/{ds}/query", "查詢年度頁面", json=body)
            for p in res.get("results", []):
                if p.get("archived") or p.get("in_trash"):
                    continue
                title = "".join(t.get("plain_text", "") for t in p["properties"][tprop]["title"]).strip()
                pages.setdefault(title, p["id"])
            cursor = res.get("next_cursor")
            if not res.get("has_more") or not cursor:
                break
        want = f"{year}年"
        pid = pages.get(want) or pages.get(str(year))
        if pid:
            print(f"Notion：找到「{want if want in pages else year}」頁面")
        else:
            page = call("POST", "/pages", "建立年度頁面", json={
                "parent": {"type": "data_source_id", "data_source_id": ds},
                "properties": {tprop: {"title": [{"text": {"content": want}}]}}})
            pid = page["id"]
            print(f"Notion：已建立「{want}」")
        cursor, old = None, []
        while True:
            q = "?page_size=100" + (f"&start_cursor={cursor}" if cursor else "")
            res = call("GET", f"/blocks/{pid}/children{q}", "讀取頁面內容")
            for b in res.get("results", []):
                if b.get("type") == "file":
                    nm = (b["file"].get("name") or "")
                    if "自動生成" in nm or ASCII_NAME in nm:
                        old.append(b["id"])
            cursor = res.get("next_cursor")
            if not res.get("has_more") or not cursor:
                break
        fname = os.path.basename(path)
        fu = call("POST", "/file_uploads", "建立檔案上傳", json={
            "mode": "single_part", "filename": fname, "content_type": XLSX})
        with open(path, "rb") as fh:
            call("POST", f"/file_uploads/{fu['id']}/send", "傳送檔案", files={"file": (ASCII_NAME, fh, XLSX)})
        call("PATCH", f"/blocks/{pid}/children", "放入頁面", json={"children": [{
            "object": "block", "type": "file",
            "file": {"type": "file_upload", "file_upload": {"id": fu["id"]}}}]})
        for bid in old:
            call("DELETE", f"/blocks/{bid}", "刪除舊自動檔")
        print(f"已把檔案放在頁面內容最底｜已替換舊自動檔 {len(old)} 個（手動檔不受影響）")
        return True
    except Exception as e:
        print(f"Notion 上傳失敗｜{e}")
        return False


def main():
    body = get(BASE + "/zh" + PATH)
    slugs = list(dict.fromkeys(re.findall(re.escape(PATH) + r"([A-Za-z0-9_\-]+)", body)))
    print(f"=== 讀取產品頁：{len(slugs)} 個 ===")
    recs, years, failed, seen, dups = [], {}, [], {}, []
    for slug in slugs:
        try:
            before = len(dups)
            got = parse_page(slug, get(BASE + "/zh" + PATH + slug + "/"), years, seen, dups)
            if not got and len(dups) == before:
                failed.append(f"{slug}（沒有資料行）")
            recs.extend(got)
        except Exception as e:
            failed.append(f"{slug}（{str(e)[:50]}）")
    if not recs:
        print("讀不到任何資料，停止")
        sys.exit(1)
    year = Counter(x["year"] for x in recs if x["year"]).most_common(1)
    year = year[0][0] if year else str(datetime.date.today().year - 1)
    dups = list(dict.fromkeys(dups))
    prods = products(recs)

    print("=== 各產品（行數/數值格數）===")
    for p in prods:
        print(f"  {p['name']}｜{p['cat'] or '未識別種類'}｜{p['rows']}/{len(p['vals'])}")
    nums = sum(len(p["vals"]) for p in prods)
    print("=== 核對 ===")
    print(f"報告年度 {year}｜產品頁 {len(slugs)} 個｜產品 {len(prods)} 個（同名只算一次）｜資料行 {len(recs)}｜數值格 {nums}")
    print(f"已略過的重複表格：{len(dups)} 個" + ("｜" + "、".join(dups) if dups else ""))
    pairs = {(x["slug"], name_key(x["name"])): x["name"] for x in recs}
    cnt = Counter(k[1] for k in pairs)
    same = list(dict.fromkeys(nm for (s, k), nm in pairs.items() if cnt[k] > 1))
    print(f"同名但數值不同、出現在多個產品頁的產品：{len(same)} 個" + ("｜" + "、".join(same) if same else ""))
    print("年份欄：" + "｜".join(f"{lab}={years.get(lab, '?')}" for lab in LABELS))
    for g in ["所有分紅"] + list(dict.fromkeys(group_of(x["typ"]) for x in recs)):
        rs = recs if g == "所有分紅" else [x for x in recs if group_of(x["typ"]) == g]
        hold = [suffix_avg(rs, i) for i in range(len(LABELS))]
        print(f"{g[:12]}（{len(rs)} 行）持滿平均｜" +
              "｜".join(f"{lab}:{a * 100:.2f}%" if a is not None else f"{lab}:/" for lab, a in zip(LABELS, hold)))
    parts = []
    for lab in LABELS:
        a, n = avg(recs, lab)
        parts.append(f"{lab}:{a * 100:.2f}%({n})" if a is not None else f"{lab}:/")
    print("所有分紅 當年平均｜" + "｜".join(parts))
    print(f"讀取失敗或沒有資料的頁：{len(failed)} 個" + ("｜" + "、".join(failed) if failed else ""))

    path = f"output/富衛分紅實現率（{year}年）自動生成.xlsx"
    build_excel(recs, years, year, path)
    print(f"已產生 {path}")
    if not notion_upload(path, year):
        sys.exit(1)


if __name__ == "__main__":
    main()
