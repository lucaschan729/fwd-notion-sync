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

BASE = "https://www.fwd.com.hk"
PATH = "/regulatory-disclosures/fulfillment-ratios/"
LABELS = [str(i) for i in range(1, 11)] + ["10+"]
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ASCII_NAME = "FWD_fulfillment_auto.xlsx"


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


def parse_page(slug, raw, years):
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', raw, flags=re.S)
    if not m:
        raise RuntimeError("沒有 NEXT_DATA")
    secs, recs, seen = [], [], set()
    find_sections(json.loads(m.group(1)), secs)
    for sec in secs:
        rows = []
        find_rows(sec.get("table"), rows)
        if not rows:
            continue
        texts = [[clean(c.get("content")) for c in r] for r in rows]
        sig = json.dumps(texts, ensure_ascii=False)
        if sig in seen:
            continue
        seen.add(sig)
        name, cat, year = parse_desc(clean(sec.get("description")))
        name = name or clean(sec.get("title")) or slug
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


def pool(rs, labs):
    return [r["vals"][lab] for r in rs for lab in labs if isinstance(r["vals"].get(lab), float)]


def mean(v):
    return sum(v) / len(v) if v else None


def avg(rs, lab):
    v = pool(rs, [lab])
    return mean(v), len(v)


def build_excel(recs, years, year, path, n_pages):
    today = (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime("%Y-%m-%d")
    wb = Workbook()
    ov = wb.active
    ov.title = "概覽"
    sname = f"{year}分紅實現率"
    ws = wb.create_sheet(sname)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="1F4E78")
    sub_fill = PatternFill("solid", fgColor="DDEBF7")
    sep_fill = PatternFill("solid", fgColor="F2F2F2")
    green = PatternFill("solid", fgColor="E2F0D9")
    red = PatternFill("solid", fgColor="FCE4E4")
    white = PatternFill("solid", fgColor="FFFFFF")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    title_font = Font(bold=True, size=14, color="1F4E78")
    sec_font = Font(bold=True, size=12, color="1F4E78")
    note_font = Font(size=9, color="7F7F7F")
    white_bold = Font(bold=True, color="FFFFFF")
    ncol = 4 + len(LABELS)
    last = get_column_letter(ncol)

    def rate(sheet, r, c, v, fmt):
        cell = sheet.cell(row=r, column=c, value=v if v is not None else "/")
        cell.alignment = center
        cell.border = border
        if isinstance(v, float):
            cell.number_format = fmt
            cell.fill = green if round(v, 6) >= 1 else (red if v < 0.8 else white)
        else:
            cell.font = Font(color="BFBFBF")
        return cell

    def head(sheet, r, c, v, fill=head_fill, font=white_bold):
        cell = sheet.cell(row=r, column=c, value=v)
        cell.font = font
        cell.fill = fill
        cell.alignment = center
        cell.border = border
        return cell

    # ---------- 資料工作表 ----------
    notes = [
        (f"富衛分紅實現率（報告年度 {year}）", title_font, 26),
        ("分紅實現率＝實際派發的非保證紅利 ÷ 銷售時說明的金額。100% 即與預期一致，高於 100% 即比預期多。定義以富衛網站的說明為準。", Font(size=10, color="404040"), 20),
        ("綠色＝100% 或以上；紅色＝低於 80%；「/」＝網站顯示「不適用」。", Font(size=10, color="404040"), 18),
        (f"資料來源：富衛香港網站（報告年度 {year}）｜更新日期：{today}", note_font, 18),
    ]
    for i, (text, font, h) in enumerate(notes, 1):
        ws.merge_cells(f"A{i}:{last}{i}")
        ws.cell(row=i, column=1, value=text).font = font
        ws.cell(row=i, column=1).alignment = left
        ws.row_dimensions[i].height = h
    HR = len(notes) + 1
    heads = ["產品名稱", "產品種類", "貨幣", "紅利類型"] + [
        f"保單年度{lab}" + (f"\n{years[lab]}" if lab in years else "") for lab in LABELS]
    for c, h in enumerate(heads, 1):
        head(ws, HR, c, h)
    ws.row_dimensions[HR].height = 36

    r, prev = HR + 1, None
    for rec in recs:
        key = (rec["slug"], rec["name"])
        if prev is not None and key != prev:
            ws.merge_cells(f"A{r}:{last}{r}")
            ws.cell(row=r, column=1).fill = sep_fill
            ws.row_dimensions[r].height = 6
            r += 1
        prev = key
        for c, v in enumerate([rec["name"], rec["cat"], rec["cur"], rec["typ"]], 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = border
            cell.fill = white
            cell.alignment = left if c <= 2 else center
            if c == 1:
                cell.font = Font(bold=True, color="262626")
            elif c == 2:
                cell.font = Font(color="595959")
            elif c == 3:
                cell.font = Font(bold=True, color="7030A0")
        for c, lab in enumerate(LABELS, 5):
            v = rec["vals"].get(lab, "/")
            fmt = "0%"
            if isinstance(v, float) and abs(v * 100 - round(v * 100)) > 1e-6:
                fmt = "0.0%"
            rate(ws, r, c, v if isinstance(v, float) else None, fmt)
        ws.row_dimensions[r].height = 20
        r += 1
    data_end = r - 1
    ws.auto_filter.ref = f"A{HR}:{last}{data_end}"
    ws.freeze_panes = f"B{HR + 1}"
    for c, w in enumerate([38, 26, 10, 14] + [12] * len(LABELS), 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    # ---------- 平均值摘要（工作表最底）----------
    groups = ["所有紅利"] + list(dict.fromkeys(group_of(x["typ"]) for x in recs))

    def rows_of(g):
        return recs if g == "所有紅利" else [x for x in recs if group_of(x["typ"]) == g]

    r += 1
    ws.cell(row=r, column=4, value="平均值摘要").font = sec_font
    r += 1
    cum_heads = ["所有年份平均值"] + [f"持滿{n}年平均值" for n in range(2, 10)] + ["第10年平均值", "持滿10年+平均值"]
    head(ws, r, 4, "分紅類型")
    for c, h in enumerate(cum_heads, 5):
        head(ws, r, c, h)
    ws.row_dimensions[r].height = 34
    r += 1
    for g in groups:
        cell = ws.cell(row=r, column=4, value=g)
        cell.font = Font(bold=True)
        cell.alignment = center
        cell.border = border
        for i in range(len(LABELS)):
            rate(ws, r, 5 + i, mean(pool(rows_of(g), LABELS[i:])), "0.00%")
        ws.row_dimensions[r].height = 20
        r += 1
    head(ws, r, 4, "當年平均值", sub_fill, Font(bold=True, color="1F4E78"))
    for c, lab in enumerate(LABELS, 5):
        head(ws, r, c, years.get(lab, lab), sub_fill, Font(bold=True, color="1F4E78"))
    r += 1
    cell = ws.cell(row=r, column=4, value="所有紅利")
    cell.font = Font(bold=True)
    cell.alignment = center
    cell.border = border
    for c, lab in enumerate(LABELS, 5):
        rate(ws, r, c, mean(pool(recs, [lab])), "0.00%")
    ws.row_dimensions[r].height = 20
    r += 1
    ws.merge_cells(f"D{r}:{last}{r}")
    ws.cell(row=r, column=4, value="「持滿N年平均值」＝保單年度N及之後所有數字的平均；「當年平均值」＝該年度一欄所有數字的平均。平均值計入所有貨幣，為固定數字，不會隨篩選變動。").font = note_font
    ws.cell(row=r, column=4).alignment = left
    ws.row_dimensions[r].height = 30

    # ---------- 概覽 ----------
    nums = pool(recs, LABELS)
    ov["A1"] = f"富衛分紅表現概覽（報告年度 {year}）"
    ov["A1"].font = title_font
    ov.row_dimensions[1].height = 26
    ov["A2"] = f"資料來源：富衛香港網站｜更新日期：{today}"
    ov["A2"].font = note_font

    ov["A4"] = "資料概況"
    ov["A4"].font = sec_font
    for c, h in enumerate(["工作表", "產品數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"], 1):
        head(ov, 5, c, h)
    link = ov.cell(row=6, column=1, value=sname)
    link.hyperlink = f"#'{sname}'!A1"
    link.font = Font(color="0563C1", underline="single")
    link.border = border
    link.alignment = center
    stats = [len({(x["slug"], x["name"]) for x in recs}), len(recs), len(nums)]
    for c, v in enumerate(stats, 2):
        cell = ov.cell(row=6, column=c, value=v)
        cell.alignment = center
        cell.border = border
    rate(ov, 6, 5, mean(nums), "0.00%")
    cell = ov.cell(row=6, column=6, value=(sum(1 for v in nums if v < 0.8) / len(nums)) if nums else "/")
    cell.number_format = "0.0%"
    cell.alignment = center
    cell.border = border

    ov["A8"] = "按產品種類的整體平均值"
    ov["A8"].font = sec_font
    head(ov, 9, 1, "產品種類")
    head(ov, 9, 2, "分紅實現率")
    rr = 10
    for cat in dict.fromkeys(x["cat"] or "未識別類別" for x in recs):
        rs = [x for x in recs if (x["cat"] or "未識別類別") == cat]
        cell = ov.cell(row=rr, column=1, value=cat)
        cell.border = border
        cell.alignment = left
        rate(ov, rr, 2, mean(pool(rs, LABELS)), "0.00%")
        rr += 1

    rr += 1
    ov.cell(row=rr, column=1, value="顏色說明").font = sec_font
    rr += 1
    legend = [
        ("100%", green, "100% 或以上：與銷售時的說明一致或更好"),
        ("79%", red, "低於 80%：明顯低於銷售時的說明"),
        ("/", white, "沒有數字：網站顯示「不適用」"),
    ]
    for sample, fill, text in legend:
        cell = ov.cell(row=rr, column=1, value=sample)
        cell.fill = fill
        cell.alignment = center
        cell.border = border
        ov.cell(row=rr, column=2, value=text)
        rr += 1

    rr += 1
    ov.cell(row=rr, column=1, value="使用提示").font = sec_font
    rr += 1
    tips = [
        "・點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
        f"・工作表第 {HR} 行有篩選箭嘴，可按產品種類、貨幣或紅利類型篩選。",
        "・向右或向下捲動時，標題列和產品名稱會固定不動。",
        "・平均值摘要在工作表最底部。",
        "・手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。",
        "・過往表現不代表未來。",
    ]
    for t in tips:
        ov.cell(row=rr, column=1, value=t)
        rr += 1
    ov.column_dimensions["A"].width = 34
    for c in range(2, 7):
        ov.column_dimensions[get_column_letter(c)].width = 16

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
    recs, years, failed = [], {}, []
    for slug in slugs:
        try:
            got = parse_page(slug, get(BASE + "/zh" + PATH + slug + "/"), years)
            if not got:
                failed.append(f"{slug}（沒有資料行）")
            recs.extend(got)
        except Exception as e:
            failed.append(f"{slug}（{str(e)[:50]}）")
    if not recs:
        print("讀不到任何資料，停止")
        sys.exit(1)
    year = Counter(x["year"] for x in recs if x["year"]).most_common(1)
    year = year[0][0] if year else str(datetime.date.today().year - 1)

    print("=== 各產品（行數/數值格數）===")
    per = {}
    for x in recs:
        k = (x["name"], x["cat"])
        a = per.setdefault(k, [0, 0])
        a[0] += 1
        a[1] += sum(1 for v in x["vals"].values() if isinstance(v, float))
    for (name, cat), (nr, nv) in per.items():
        print(f"  {name}｜{cat or '未識別類別'}｜{nr}/{nv}")
    nums = sum(1 for x in recs for v in x["vals"].values() if isinstance(v, float))
    print("=== 核對 ===")
    print(f"報告年度 {year}｜產品 {len(per)} 個｜資料行 {len(recs)}｜數值格 {nums}")
    print("年份欄：" + "｜".join(f"{lab}={years.get(lab, '?')}" for lab in LABELS))
    for g in ["所有紅利"] + list(dict.fromkeys(group_of(x["typ"]) for x in recs)):
        rs = recs if g == "所有紅利" else [x for x in recs if group_of(x["typ"]) == g]
        parts = []
        for lab in LABELS:
            a, n = avg(rs, lab)
            parts.append(f"{lab}:{a * 100:.2f}%({n})" if a is not None else f"{lab}:/")
        print(f"{g[:12]}（{len(rs)} 行）平均｜" + "｜".join(parts))
    print(f"讀取失敗或沒有資料的頁：{len(failed)} 個" + ("｜" + "、".join(failed) if failed else ""))

    path = f"output/富衛分紅實現率（{year}年）自動生成.xlsx"
    build_excel(recs, years, year, path, len(slugs))
    print(f"已產生 {path}")
    if not notion_upload(path, year):
        sys.exit(1)


if __name__ == "__main__":
    main()
