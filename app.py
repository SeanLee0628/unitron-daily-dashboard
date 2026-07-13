#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""일일 입출고 리포트 대시보드 (업로드형).

암호화된 'daily material shipping & receiving' 엑셀을 업로드하면 → 날짜 시트
(어제·오늘)만 읽어 → 일일 입출고 리포트 대시보드. 입고/출고 마스터 시트는 안 씀.

표준 라이브러리만 사용. 엑셀은 브라우저에서 base64로 전송한다.
"""
from __future__ import annotations

import base64
import io
import json
import os
import re
import socket
import threading
import webbrowser
from collections import Counter, defaultdict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import msoffcrypto
import openpyxl                                   # Excel 내보내기에만 쓴다 (쓰기)

import fastxl                                     # 읽기 — openpyxl 보다 11배 빠름

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_PW = os.environ.get("XLSX_PW", "")  # 비밀번호는 코드에 두지 않음 (업로드 화면 입력 또는 환경변수)
CHART_JS = os.path.join(HERE, "chart.umd.min.js")
DATA_FILE = os.path.join(os.environ.get("DATA_DIR", HERE), "saved_data.json")  # 데이터 저장(공유)


def clean(x):
    if x is None:
        return ""
    s = str(x).strip()
    return "" if s in (".", "-", "nan", "None") else s


def to_num(x):
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


# ---------------------------------------------------------------- 이름 정규화
# 같은 사람/같은 회사가 다른 표기로 들어와 집계가 쪼개진다.
#   담당: '박현진' / '박현진 책임매니저'  → 같은 사람
#   거래처: 'Mobis' / 'MOBIS'            → 같은 회사 (대소문자만 다름)
TITLES = ("책임매니저", "선임매니저", "수석매니저", "매니저", "책임연구원", "선임연구원",
          "연구원", "책임", "선임", "수석", "프로", "사원", "주임", "대리", "과장",
          "차장", "부장", "팀장", "실장", "이사", "상무", "전무", "대표", "님")
_TITLE_RE = re.compile(r"\s*(" + "|".join(TITLES) + r")\s*$")


def norm_sales(s):
    """담당자 이름 → 직급 제거. '박현진 책임매니저' → '박현진'."""
    if not s:
        return "(미지정)"
    n = str(s).split("/")[0].strip()
    prev = None
    while n and n != prev:                 # '책임매니저' 처럼 겹친 직급도 벗겨낸다
        prev = n
        n = _TITLE_RE.sub("", n).strip()
    return n or "(미지정)"


def company_key(s):
    """비교용 키 — 대소문자·공백 무시. 거래처와 담당자 모두 같은 규칙."""
    return re.sub(r"\s+", "", str(s or "")).upper()


class Companies:
    """같은 회사의 여러 표기를 하나로 모은다 (Mobis / MOBIS).

    표시 이름은 가장 많이 쓰인 표기를 쓴다 — 임의로 대문자화하지 않는다.
    담당자에는 쓰지 않는다 (영어 이름 James/JAMES 는 그대로 둔다).
    """

    def __init__(self):
        self.seen = defaultdict(Counter)   # 키 -> Counter(원본 표기)

    def add(self, name):
        n = clean(name)
        if n:
            self.seen[company_key(n)][n] += 1
        return n

    def canon(self, name):
        n = clean(name)
        if not n:
            return n
        c = self.seen.get(company_key(n))
        return c.most_common(1)[0][0] if c else n


# ---------------------------------------------------------------- 파싱
def _hkey(h):
    """헤더 정규화: 대문자·공백/'/'/'  제거 → PART#, QTY, CUSTOMER, 담당SALES 등."""
    return clean(h).upper().replace(" ", "").replace("'", "").replace("/", "")


def parse_daily(ws):
    """날짜 시트 1개 → (inbound[], outbound[]). 컬럼은 헤더로 동적 인식(실마다 배치 달라도 OK)."""
    mode = None
    cols = {}
    inbound, outbound = [], []
    for r in ws.iter_rows(values_only=True):
        if not r:
            continue
        c0 = clean(r[0])
        if c0 and not c0.isdigit():
            if c0 == "NO":                          # 헤더행 → 컬럼 위치 매핑
                cols = {_hkey(h): i for i, h in enumerate(r) if clean(h)}
                continue
            if "입고" in c0:
                mode = "in"; continue
            if "출고" in c0:
                mode = "out"; continue
            continue                                # 기타 텍스트(영업실 소제목 등)
        if not c0 or not c0.isdigit():
            continue

        def col(*keys):
            for k in keys:
                i = cols.get(k)
                if i is not None and i < len(r):
                    v = clean(r[i])
                    if v:
                        return v
            return ""

        def coln(*keys):
            for k in keys:
                i = cols.get(k)
                if i is not None and i < len(r):
                    return to_num(r[i])
            return 0.0

        part = col("PART#", "PART", "MPN")
        if not part:
            continue
        qty = coln("QTY", "QUANTITY")
        sales = norm_sales(col("담당SALES", "SALES"))
        if mode == "in":
            inbound.append(dict(
                part=part, qty=qty, customer=col("CUSTOMER", "SR#", "공급처"),
                sales=sales, fab=col("FAB"), remark=col("REMARK"),
            ))
        elif mode == "out":
            outbound.append(dict(
                customer=col("CUSTOMER"), part=part, qty=qty, sales=sales,
                doc=col("문서번호", "DOC"), remark=col("REMARK"), vendor=col("VENDOR"),
            ))
    return inbound, outbound


def parse_workbook(wb):
    """워크북 → {날짜: (inbound, outbound)} (날짜 시트만)."""
    per = {}
    for s in wb.sheetnames:
        if re.match(r"\d{4}-\d{2}-\d{2}", s):
            per[s] = parse_daily(wb[s])
    return per


# ---------------------------------------------------------------- 재고 현황
OLD_YEAR = 2022          # 이 해 이전 datecode = 장기 재고


def inventory_sheet(wb):
    """'Jul inventory' 처럼 이름이 inventory 로 끝나는 시트를 찾는다."""
    for s in wb.sheetnames:
        if s.strip().lower().endswith("inventory"):
            return wb[s]
    return None


def parse_inventory(ws):
    """재고 시트 → 대시보드용 집계. 헤더는 2행, 데이터는 3행부터."""
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    if len(rows) < 3:
        return None
    hdr = [clean(h) for h in rows[1]]
    i_part = next((i for i, h in enumerate(hdr)
                   if h.replace(" ", "").upper() == "PART#"), -1)
    if i_part < 0:
        return None
    # 시트 하단에 Part# 없는 합계 행이 섞여 있다. 그대로 더하면 정확히 2배가 된다.
    data = [r for r in rows[2:] if r and i_part < len(r) and clean(r[i_part])]

    def idx(*names):
        for n in names:
            for i, h in enumerate(hdr):
                if h.replace(" ", "").upper() == n.replace(" ", "").upper():
                    return i
        return -1

    C = dict(
        office=idx("Sales team"), central=idx("Central"), vender=idx("VENDER"),
        family=idx("FAMILY"), part=idx("Part#"), mobis=idx("MOBIS ID"),
        pn=idx("품번"), qty=idx("Q'ty"), avail=idx("available Q'ty"),
        booking=idx("booking"), customer=idx("CUSTOMER"), sales=idx("SALES"),
        crd=idx("CRD"),
    )
    if C["part"] < 0 or C["qty"] < 0:
        return None

    # Datecode 연도 컬럼 (헤더에 4자리 연도가 들어있음)
    years = {}
    for i, h in enumerate(hdr):
        if "DATECODE" in h.upper():
            m = re.search(r"(20\d{2})", h)
            if m:
                years[int(m.group(1))] = i

    # 일별 컬럼: '1일'~'31일' 이 두 번 반복 (앞 31개=입고, 뒤 31개=출고)
    days = [i for i, h in enumerate(hdr) if re.fullmatch(r"\d{1,2}일", h)]
    d_in, d_out = days[:31], days[31:62]
    i_prev = idx("전월")

    def g(r, key):
        i = C[key]
        return r[i] if 0 <= i < len(r) else None

    items, tot_q, tot_a, tot_b = [], 0.0, 0.0, 0.0
    by = {k: defaultdict(lambda: [0, 0.0]) for k in ("office", "vender", "family")}
    dc = {y: [0, 0.0] for y in years}
    cust_bk = defaultdict(float)

    for r in data:
        if not r or not clean(g(r, "part")):
            continue
        q = to_num(g(r, "qty"))
        if q <= 0:                                  # 재고 없는 품목은 제외
            continue
        a = to_num(g(r, "avail"))
        b = to_num(g(r, "booking"))
        tot_q += q; tot_a += a; tot_b += b

        for k in ("office", "vender", "family"):
            key = clean(g(r, k)) or "(미지정)"
            by[k][key][0] += 1
            by[k][key][1] += q

        old_q, oldest = 0.0, None
        for y, ci in years.items():
            v = to_num(r[ci]) if ci < len(r) else 0.0
            if v > 0:
                dc[y][0] += 1; dc[y][1] += v
                if y <= OLD_YEAR:
                    old_q += v
                if oldest is None or y < oldest:
                    oldest = y

        c = clean(g(r, "customer"))
        if c and c != "." and b:
            cust_bk[c] += b

        items.append(dict(
            part=clean(g(r, "part")), mobis=clean(g(r, "mobis")),
            pn=clean(g(r, "pn")), family=clean(g(r, "family")),
            vender=clean(g(r, "vender")), office=clean(g(r, "office")),
            sales=clean(g(r, "sales")), customer=c, crd=clean(g(r, "crd")),
            qty=q, avail=a, booking=b, old=old_q, oldest=oldest,
        ))

    def top(k):
        return [dict(name=n, items=v[0], qty=v[1])
                for n, v in sorted(by[k].items(), key=lambda x: -x[1][1])]

    daily_in = [sum(to_num(r[i]) for r in data if i < len(r)) for i in d_in]
    daily_out = [sum(to_num(r[i]) for r in data if i < len(r)) for i in d_out]

    return dict(
        sheet=ws.title,
        n_items=len(items),
        total_qty=tot_q, avail_qty=tot_a, booking_qty=tot_b,
        old_qty=sum(v[1] for y, v in dc.items() if y <= OLD_YEAR),
        old_year=OLD_YEAR,
        by_office=top("office"), by_vender=top("vender"), by_family=top("family"),
        datecode=[dict(year=y, items=dc[y][0], qty=dc[y][1]) for y in sorted(dc)],
        month=dict(
            prev=sum(to_num(r[i_prev]) for r in data if 0 <= i_prev < len(r)),
            inbound=sum(daily_in), outbound=sum(daily_out),
            daily_in=daily_in, daily_out=daily_out,
        ),
        customers=[dict(name=n, booking=v)
                   for n, v in sorted(cust_bk.items(), key=lambda x: -x[1])[:12]],
        items=items,
    )


def day_block(date, inbound, outbound):
    cust = defaultdict(lambda: [0.0, 0])
    for r in outbound:
        c = r["customer"] or "(미지정)"
        cust[c][0] += r["qty"]; cust[c][1] += 1
    top_cust = sorted(cust.items(), key=lambda kv: kv[1][0], reverse=True)[:8]
    sales_d = defaultdict(int)
    for r in inbound + outbound:
        sales_d[r["sales"]] += 1
    top_sales = sorted(sales_d.items(), key=lambda kv: kv[1], reverse=True)[:8]
    hl = max(outbound, key=lambda r: r["qty"], default=None)
    kpi = dict(
        in_cnt=len(inbound), in_qty=round(sum(r["qty"] for r in inbound)),
        out_cnt=len(outbound), out_qty=round(sum(r["qty"] for r in outbound)),
        customers=len({r["customer"] for r in outbound if r["customer"]}),
    )
    kpi["net"] = kpi["in_qty"] - kpi["out_qty"]
    return dict(
        date=date, kpi=kpi,
        cust=[dict(name=k, qty=round(v[0]), cnt=v[1]) for k, v in top_cust],
        sales=[dict(name=k, cnt=v) for k, v in top_sales],
        highlight=(dict(customer=hl["customer"], part=hl["part"], qty=round(hl["qty"]), sales=hl["sales"]) if hl else None),
        in_rows=[dict(part=r["part"], qty=round(r["qty"]), customer=r["customer"],
                      sales=r["sales"], fab=r["fab"], remark=r["remark"]) for r in inbound],
        out_rows=[dict(customer=r["customer"], part=r["part"], qty=round(r["qty"]),
                       sales=r["sales"], doc=r["doc"], remark=r["remark"], vendor=r["vendor"]) for r in outbound],
    )


def open_wb(raw, password):
    """읽기는 fastxl 로 한다. openpyxl 은 styles.xml(14MB)을 통째로 객체화하느라
    파일당 5~8초를 쓰는데, 우리는 값만 필요하다. 실측 46초 → 4초."""
    if raw[:4] != b"\xd0\xcf\x11\xe0":                   # 1) 평문 xlsx
        try:
            return fastxl.load(io.BytesIO(raw))
        except Exception:
            pass
    try:                                                 # 2) 암호화 → 복호화
        off = msoffcrypto.OfficeFile(io.BytesIO(raw))
        off.load_key(password=password or "")
        dec = io.BytesIO(); off.decrypt(dec)
        return fastxl.load(dec)
    except Exception:
        raise ValueError("엑셀을 열 수 없습니다. 비밀번호가 비었거나 틀렸을 수 있어요 — 업로드 화면의 '비밀번호' 칸에 입력하세요 (예: 9178).")


def office_block(name, per):
    """{날짜:(inb,outb)} → 실 1개 대시보드 데이터."""
    dates = sorted(per)
    days = [day_block(d, per[d][0], per[d][1]) for d in dates]
    return dict(name=name, dates=dates, today_idx=len(days) - 1, days=days)


ALL = "전체 합계"


def merge_inventories(invs):
    """실별 재고 → 전체 합계 재고. 품목 리스트를 합치고 집계를 다시 계산한다."""
    if not invs:
        return None
    items = [dict(it) for inv in invs for it in inv["items"]]
    years = sorted({d["year"] for inv in invs for d in inv["datecode"]})
    ndays = max((len(inv["month"]["daily_in"]) for inv in invs), default=31)

    def agg(key):
        acc = defaultdict(lambda: [0, 0.0])
        for inv in invs:
            for r in inv[key]:
                acc[r["name"]][0] += r["items"]
                acc[r["name"]][1] += r["qty"]
        return [dict(name=n, items=v[0], qty=v[1])
                for n, v in sorted(acc.items(), key=lambda x: -x[1][1])]

    def dsum(key, i):
        return sum(inv["month"][key][i] if i < len(inv["month"][key]) else 0.0
                   for inv in invs)

    cust = defaultdict(float)
    for inv in invs:
        for c in inv["customers"]:
            cust[c["name"]] += c["booking"]

    dc = []
    for y in years:
        n = q = 0
        for inv in invs:
            for d in inv["datecode"]:
                if d["year"] == y:
                    n += d["items"]; q += d["qty"]
        dc.append(dict(year=y, items=n, qty=q))

    s = lambda k: sum(inv[k] for inv in invs)
    return dict(
        sheet=", ".join(sorted({inv["sheet"] for inv in invs})),
        n_items=len(items),
        total_qty=s("total_qty"), avail_qty=s("avail_qty"),
        booking_qty=s("booking_qty"), old_qty=s("old_qty"),
        old_year=invs[0]["old_year"],
        by_office=agg("by_office"), by_vender=agg("by_vender"), by_family=agg("by_family"),
        datecode=dc,
        month=dict(prev=sum(inv["month"]["prev"] for inv in invs),
                   inbound=sum(inv["month"]["inbound"] for inv in invs),
                   outbound=sum(inv["month"]["outbound"] for inv in invs),
                   daily_in=[dsum("daily_in", i) for i in range(ndays)],
                   daily_out=[dsum("daily_out", i) for i in range(ndays)]),
        customers=[dict(name=n, booking=v)
                   for n, v in sorted(cust.items(), key=lambda x: -x[1])[:12]],
        items=items,
    )


def build_payload(files, password):
    """files: [{name, file(base64)}] → {offices:[...], inventories:{실: {...}}}

    파일 종류는 시트 이름으로 자동 판별한다.
      날짜 시트(YYYY-MM-DD) 있음  → 일일 입출고 파일
      '~ inventory' 시트 있음      → 재고 현황 파일
    재고도 실별로 보관한다. (전에는 하나만 남아 마지막 파일이 앞의 것을 덮어썼다)
    """
    raw_offices, inventories = [], {}     # [(실이름, {날짜:(inb,outb)})]
    for f in files:
        b64 = f["file"]
        if "," in b64[:64]:
            b64 = b64.split(",", 1)[1]
        name = f.get("name") or "실"
        wb = open_wb(base64.b64decode(b64), password)
        try:
            per = parse_workbook(wb)
            if per:                                     # 입출고 파일
                raw_offices.append((name, per))
                continue
            ws = inventory_sheet(wb)                    # 재고 파일
            if ws is not None:
                inv = parse_inventory(ws)
                if inv:
                    inventories[name] = inv
        finally:
            wb.close()

    if not raw_offices and not inventories:
        raise ValueError("날짜 시트(YYYY-MM-DD)가 있는 입출고 파일이나 "
                         "'inventory' 시트가 있는 재고 파일을 찾지 못했습니다.")

    # ---- 거래처 표기 통일 (Mobis / MOBIS → 가장 많이 쓰인 표기 하나로) ----
    # 파일 전체를 본 뒤에야 대표 표기를 고를 수 있어 여기서 일괄 처리한다.
    C = Companies()
    for _, per in raw_offices:
        for inb, outb in per.values():
            for r in inb + outb:
                C.add(r.get("customer"))
    for inv in inventories.values():
        for it in inv["items"]:
            C.add(it.get("customer"))

    for _, per in raw_offices:
        for inb, outb in per.values():
            for r in inb + outb:
                r["customer"] = C.canon(r.get("customer"))
    for inv in inventories.values():
        bk = defaultdict(float)
        for it in inv["items"]:
            it["customer"] = C.canon(it.get("customer"))
            it["sales"] = norm_sales(it.get("sales")) if it.get("sales") else ""
            if it["customer"] and it["customer"] != "." and it["booking"]:
                bk[it["customer"]] += it["booking"]
        inv["customers"] = [dict(name=n, booking=v) for n, v in
                            sorted(bk.items(), key=lambda x: -x[1])[:12]]

    offices = []
    allbydate = defaultdict(lambda: ([], []))
    for name, per in raw_offices:
        offices.append(office_block(name, per))
        for d, (inb, outb) in per.items():
            allbydate[d][0].extend(inb); allbydate[d][1].extend(outb)

    if len(offices) > 1:                                # 전체 합계 탭 (맨 앞)
        agg = {d: (allbydate[d][0], allbydate[d][1]) for d in allbydate}
        offices.insert(0, office_block(ALL, agg))
    if len(inventories) > 1:
        inventories[ALL] = merge_inventories(list(inventories.values()))

    return dict(offices=offices, inventories=inventories)


# ---------------------------------------------------------------- Excel 내보내기
def export_xlsx(payload):
    """지금 보고 있는 화면을 엑셀로. {office, view, day?/inventory?} → xlsx bytes."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    HDR = PatternFill("solid", fgColor="1D1D20")
    HDRF = Font(color="FFFFFF", bold=True, size=10)
    TITLE = Font(bold=True, size=14, color="1D1D20")
    RED = Font(bold=True, color="C43A3A")
    THIN = Side(style="thin", color="E4E4E9")
    BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    def sheet(name, title, headers, rows, widths=None, red_col=None):
        ws = wb.create_sheet(name[:31])
        ws["A1"] = title
        ws["A1"].font = TITLE
        ws.append([])
        ws.append(headers)
        for c in ws[3]:
            c.fill, c.font, c.border = HDR, HDRF, BOX
            c.alignment = Alignment(horizontal="center", vertical="center")
        for r in rows:
            ws.append(r)
        for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
            for c in row:
                c.border = BOX
                if isinstance(c.value, (int, float)):
                    c.number_format = "#,##0"
                    c.alignment = Alignment(horizontal="right")
        if red_col:
            for row in ws.iter_rows(min_row=4, max_row=ws.max_row,
                                    min_col=red_col, max_col=red_col):
                for c in row:
                    if isinstance(c.value, (int, float)) and c.value > 0:
                        c.font = RED
        for i, w in enumerate(widths or [], start=1):
            ws.column_dimensions[ws.cell(row=3, column=i).column_letter].width = w
        ws.freeze_panes = "A4"
        return ws

    office = payload.get("office") or "전체"
    view = payload.get("view") or "io"

    if view == "inv":
        inv = payload["inventory"]
        sheet("요약", f"재고 현황 · {office}",
              ["항목", "값"],
              [["재고 품목", inv["n_items"]], ["총 재고", inv["total_qty"]],
               ["가용 재고", inv["avail_qty"]], ["예약(booking)", inv["booking_qty"]],
               [f"장기재고 ({inv['old_year']}년 이전)", inv["old_qty"]],
               ["당월 입고", inv["month"]["inbound"]],
               ["당월 출고", inv["month"]["outbound"]]],
              widths=[26, 16])
        sheet("품목", f"품목별 재고 · {office}",
              ["PART#", "MOBIS ID", "FAMILY", "VENDER", "실", "담당", "고객",
               "재고", "가용", "예약", "장기재고", "Datecode"],
              [[i["part"], i["mobis"], i["family"], i["vender"], i["office"],
                i["sales"], i["customer"], i["qty"], i["avail"], i["booking"],
                i["old"], i["oldest"] or ""] for i in
               sorted(inv["items"], key=lambda x: -x["qty"])],
              widths=[26, 16, 16, 12, 10, 10, 18, 12, 12, 12, 12, 11],
              red_col=11)
        # Datecode 는 영업1,2실에만 채워져 있다 → 있는 실에서만 시트를 만든다
        dc = [d for d in inv["datecode"] if d["qty"] > 0]
        if dc:
            sheet("노후화", f"재고 노후화 (Datecode) · {office}",
                  ["연도", "품목 수", "수량"],
                  [[d["year"], d["items"], d["qty"]] for d in dc],
                  widths=[10, 12, 16])
        sheet("MOBIS별", f"MOBIS ID별 · {office}",
              ["MOBIS ID", "품목 수", "수량"],
              [[m["name"], m["items"], m["qty"]] for m in inv["by_family"]],
              widths=[20, 12, 16])
    else:
        day = payload["day"]
        k = day.get("kpi", {})
        sheet("요약", f"일일 입출고 · {office} · {day.get('date','')}",
              ["항목", "값"],
              [["입고 건수", k.get("in_cnt")], ["입고 수량", k.get("in_qty")],
               ["출고 건수", k.get("out_cnt")], ["출고 수량", k.get("out_qty")],
               ["순물동(입-출)", k.get("net")], ["출고 거래처", k.get("customers")]],
              widths=[20, 16])
        sheet("출고", f"출고 내역 · {office} · {day.get('date','')}",
              ["#", "거래처", "PART#", "수량", "담당", "문서번호", "비고"],
              [[i + 1, r.get("customer"), r.get("part"), r.get("qty"),
                r.get("sales"), r.get("doc"), r.get("remark")]
               for i, r in enumerate(day.get("out_rows", []))],
              widths=[6, 22, 26, 12, 10, 16, 22])
        sheet("입고", f"입고 내역 · {office} · {day.get('date','')}",
              ["#", "거래처/공급", "PART#", "수량", "담당", "FAB", "비고"],
              [[i + 1, r.get("customer"), r.get("part"), r.get("qty"),
                r.get("sales"), r.get("fab"), r.get("remark")]
               for i, r in enumerate(day.get("in_rows", []))],
              widths=[6, 22, 26, 12, 10, 10, 22])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- 이메일 (Outlook)
def _e(s):
    return (str(s) if s is not None else "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# 메일 하단 문의처. 답장(Reply-To)은 첫 번째 사람에게 간다.
CONTACTS = [
    ("자재관리팀", "안성우 책임", "sw.ahn@unitrontech.com"),
    ("기술지원 경영기획팀", "이희서 매니저", "seanlee@unitrontech.com"),
]
CONTACT_NAME = f"유니트론텍 {CONTACTS[0][1]}"
CONTACT_MAIL = CONTACTS[0][2]
# 받는 사람 메일함에 보이는 발신자 이름. 없으면 메일주소가 그대로 노출된다.
FROM_NAME = os.environ.get("MAIL_FROM_NAME", "유니트론텍 입출고 대시보드")


# 이메일은 웹이 아니다 — Outlook 은 flex/grid/외부CSS 를 못 읽는다.
# 테이블 레이아웃 + 인라인 스타일로만 짠다. (하우스 스타일: 차콜 + 시그니처 레드)
INK = "#1d1d20"
RED = "#c43a3a"
MUT = "#8a8a92"
LINE = "#e4e4e9"
FONT = "'Malgun Gothic','맑은 고딕',-apple-system,'Segoe UI',Roboto,sans-serif"


def compose_email_text(office, date, link=""):
    """텍스트 버전 — HTML 을 못 읽는 클라이언트용. 스팸 점수도 낮춰준다."""
    return (
        f"UNITRONTECH\n"
        f"입출고 및 재고현황\n"
        f"{office} · {date}\n\n"
        f"아래 링크에서 조회하실 수 있습니다.\n"
        f"{link}\n\n"
        f"----------------------------------------\n"
        f"문의\n"
        + "".join(f"  {team} {who} ({mail})\n" for team, who, mail in CONTACTS)
        + f"\n본 메일은 자동 발송되었습니다.\n"
    )


def compose_contacts_html():
    """푸터 문의처 — 팀 / 이름 / 메일 한 줄씩."""
    rows = []
    for i, (team, who, mail) in enumerate(CONTACTS):
        pad = "padding-top:6px;" if i else ""
        rows.append(
            f'<tr><td style="{pad}font-family:{FONT};font-size:12.5px;'
            f'color:#55555d;line-height:1.7;white-space:nowrap">'
            f'<span style="color:{MUT}">{_e(team)}</span>'
            f'&nbsp;<span style="color:{INK};font-weight:700">{_e(who)}</span>'
            f'&nbsp;<a href="mailto:{_e(mail)}" style="color:{RED};'
            f'text-decoration:none">{_e(mail)}</a></td></tr>')
    return ('<table role="presentation" cellpadding="0" cellspacing="0" border="0">'
            + "".join(rows) + "</table>")


def compose_email_html(office, date, day=None, link=""):
    cta = (
        # 버튼은 <table> 로 감싸야 Outlook 에서 모양이 유지된다
        f'<table role="presentation" cellpadding="0" cellspacing="0" border="0">'
        f'<tr><td align="center" bgcolor="{RED}" style="border-radius:6px">'
        f'<a href="{_e(link)}" style="display:inline-block;padding:15px 34px;'
        f'font-family:{FONT};font-size:14.5px;font-weight:700;color:#ffffff;'
        f'text-decoration:none;letter-spacing:-0.2px">입출고 및 재고현황 조회</a>'
        f'</td></tr></table>'
    ) if link else ""

    raw = (f'<div style="margin-top:14px;font-size:11.5px;color:{MUT};'
           f'word-break:break-all;line-height:1.6">{_e(link)}</div>') if link else ""

    return f"""<!DOCTYPE html>
<html><body style="margin:0;padding:0;background:#f2f2f5">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
       style="background:#f2f2f5;padding:32px 16px">
  <tr><td align="center">
    <table role="presentation" width="540" cellpadding="0" cellspacing="0" border="0"
           style="width:540px;max-width:100%;background:#ffffff;border:1px solid {LINE};border-radius:10px">

      <!-- 상단 시그니처 라인 -->
      <tr><td style="height:3px;background:{RED};font-size:0;line-height:0;
                     border-radius:10px 10px 0 0">&nbsp;</td></tr>

      <!-- 헤더 -->
      <tr><td style="padding:30px 36px 0 36px;font-family:{FONT}">
        <div style="font-size:10.5px;font-weight:700;color:{MUT};letter-spacing:1.6px">UNITRONTECH</div>
        <div style="margin-top:10px;font-size:20px;font-weight:700;color:{INK};letter-spacing:-0.5px">
          입출고 및 재고현황</div>
        <div style="margin-top:7px;font-size:13px;color:{MUT}">
          {_e(office)} <span style="color:{LINE}">|</span> {_e(date)}</div>
      </td></tr>

      <!-- 구분선 -->
      <tr><td style="padding:22px 36px 0 36px">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr><td style="height:1px;background:{LINE};font-size:0;line-height:0">&nbsp;</td></tr>
        </table>
      </td></tr>

      <!-- 본문 + CTA -->
      <tr><td style="padding:24px 36px 32px 36px;font-family:{FONT}">
        <div style="font-size:13.5px;color:#4a4a52;line-height:1.75;margin-bottom:22px">
          해당 일자의 입출고 내역과 재고현황을 아래에서 조회하실 수 있습니다.</div>
        {cta}
        {raw}
      </td></tr>

      <!-- 푸터 -->
      <tr><td style="padding:18px 36px 24px 36px;background:#fafafb;
                     border-top:1px solid {LINE};border-radius:0 0 10px 10px;font-family:{FONT}">
        <div style="font-size:10.5px;font-weight:700;color:{MUT};letter-spacing:1.2px;
                    margin-bottom:9px">문의</div>
        {compose_contacts_html()}
        <div style="margin-top:12px;font-size:11px;color:{MUT}">본 메일은 자동 발송되었습니다.</div>
      </td></tr>

    </table>
  </td></tr>
</table>
</body></html>"""


def resolve_smtp(cfg):
    """SMTP 설정 결정: 요청(cfg) > 환경변수. host 없으면 None."""
    cfg = cfg or {}
    host = cfg.get("host") or os.environ.get("SMTP_HOST")
    if not host:
        return None
    return dict(
        host=host,
        port=int(cfg.get("port") or os.environ.get("SMTP_PORT") or 587),
        user=cfg.get("user") or os.environ.get("SMTP_USER") or "",
        pw=cfg.get("pass") or os.environ.get("SMTP_PASS") or "",
        sender=cfg.get("from") or cfg.get("user") or os.environ.get("MAIL_FROM")
               or os.environ.get("SMTP_USER") or "no-reply@localhost",
    )


def _send_smtp(emails, subject, body, cfg, text=None):
    """SMTP 발송. cfg = {host, port, user, pw, sender}. text = 대체 텍스트 본문."""
    import smtplib, ssl
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.utils import formataddr
    host, port = cfg["host"], cfg["port"]
    user, pw, sender = cfg["user"], cfg["pw"], cfg["sender"]

    # HTML 전용 메일은 스팸 점수가 올라간다 → 텍스트 버전을 같이 보낸다
    if text:
        msg = MIMEMultipart("alternative")
        msg.attach(MIMEText(text, "plain", "utf-8"))
        msg.attach(MIMEText(body, "html", "utf-8"))
    else:
        msg = MIMEText(body, "html", "utf-8")

    msg["Subject"] = subject
    # 메일주소 대신 이름으로 보이게 (한글은 RFC2047 로 자동 인코딩됨)
    msg["From"] = formataddr((FROM_NAME, sender)) if FROM_NAME else sender
    msg["To"] = ", ".join(emails)
    # 답장은 발송용 계정이 아니라 문의처로 가야 한다
    msg["Reply-To"] = formataddr((CONTACT_NAME, CONTACT_MAIL))
    with smtplib.SMTP(host, port, timeout=25) as s:
        s.ehlo()
        try:
            s.starttls(context=ssl.create_default_context()); s.ehlo()
        except Exception:
            pass
        if user:
            s.login(user, pw)
        s.sendmail(sender, emails, msg.as_string())
    return {"ok": True, "sent": True, "via": "smtp", "n": len(emails)}


def _send_outlook(emails, subject, body, send):
    """로컬 PC Outlook 발송(폴백)."""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise ValueError("메일 발송 설정이 없습니다. 클라우드는 SMTP_HOST 등 환경변수, 로컬은 Outlook이 필요합니다.")
    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        mail = app.CreateItem(0)
        mail.To = "; ".join(emails)
        mail.Subject = subject
        mail.HTMLBody = body
        if send:
            mail.Send()
            return {"ok": True, "sent": True, "via": "outlook", "n": len(emails)}
        mail.Display(False)
        return {"ok": True, "sent": False, "via": "outlook", "n": len(emails)}
    finally:
        pythoncom.CoUninitialize()


def send_sms(numbers, text):
    """솔라피(Solapi)로 문자 자동 발송. 환경변수 SOLAPI_API_KEY/SOLAPI_API_SECRET/SMS_FROM 필요."""
    import hmac, hashlib, urllib.request
    from datetime import datetime, timezone
    key = os.environ.get("SOLAPI_API_KEY")
    secret = os.environ.get("SOLAPI_API_SECRET")
    sender = os.environ.get("SMS_FROM")
    if not (key and secret and sender):
        raise ValueError("문자 설정이 없습니다. 환경변수 SOLAPI_API_KEY / SOLAPI_API_SECRET / SMS_FROM(발신번호) 를 설정하세요.")
    nums = [re.sub(r"\D", "", n) for n in numbers if re.sub(r"\D", "", n)]
    if not nums:
        raise ValueError("받는 휴대폰 번호가 없습니다.")
    date = datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")
    salt = os.urandom(16).hex()
    sig = hmac.new(secret.encode(), (date + salt).encode(), hashlib.sha256).hexdigest()
    auth = f"HMAC-SHA256 apiKey={key}, date={date}, salt={salt}, signature={sig}"
    frm = re.sub(r"\D", "", sender)
    messages = [{"to": n, "from": frm, "text": text} for n in nums]
    body = json.dumps({"messages": messages}).encode("utf-8")
    req = urllib.request.Request("https://api.solapi.com/messages/v4/send-many",
                                 data=body, method="POST",
                                 headers={"Authorization": auth, "Content-Type": "application/json"})
    try:
        resp = json.load(urllib.request.urlopen(req, timeout=20))
    except urllib.error.HTTPError as e:
        raise ValueError(f"문자 발송 실패: {e.read().decode('utf-8', 'ignore')[:300]}")
    gi = resp.get("groupInfo", {}).get("count", {}) if isinstance(resp, dict) else {}
    return {"ok": True, "n": len(nums), "count": gi}


def send_email(payload):
    """실 링크(+요약·내역) 이메일 발송. SMTP 설정(요청/환경변수) 있으면 SMTP, 없으면 로컬 Outlook."""
    emails = [e for e in (payload.get("emails") or []) if e]
    office = payload.get("office", "")
    date = payload.get("date", "")
    day = payload.get("day") or {}
    link = payload.get("link", "")
    subject = f"[입출고 및 재고현황 {date}]"
    body = compose_email_html(office, date, day, link)
    text = compose_email_text(office, date, link)
    cfg = resolve_smtp(payload.get("smtp"))
    if cfg:                                             # SMTP (요청 설정 또는 환경변수)
        if not emails:
            raise ValueError("받는 사람 이메일을 입력하세요.")
        return _send_smtp(emails, subject, body, cfg, text=text)
    return _send_outlook(emails, subject, body, payload.get("send"))  # 폴백: 로컬 Outlook


# ---------------------------------------------------------------- HTTP
def chart_js():
    if os.path.exists(CHART_JS):
        with open(CHART_JS, encoding="utf-8") as f:
            return f"<script>{f.read()}</script>"
    return '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>'


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        path = self.path.split("?")[0]
        if path == "/mailcfg":                           # 서버에 보내는 메일계정(SMTP) 설정됐나
            self._send(200, json.dumps({"configured": bool(resolve_smtp(None))}))
        elif path == "/data":                            # 저장된 데이터(있으면)
            if os.path.exists(DATA_FILE):
                with open(DATA_FILE, encoding="utf-8") as f:
                    self._send(200, f.read())
            else:
                self._send(200, json.dumps({"offices": []}))
        else:                                            # / , /12 , /3 , /4 , /5 ... 모두 같은 페이지(클라이언트 라우팅)
            self._send(200, PAGE.replace("<!--CHARTJS-->", chart_js()), "text/html; charset=utf-8")

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            payload = json.loads(self.rfile.read(length))
            if self.path == "/build":
                files = payload.get("files") or [{"name": "", "file": payload["file"]}]
                result = build_payload(files, payload.get("password") or DEFAULT_PW)
                try:                                     # 업로드 결과 저장 → 링크 연 사람 모두 공유
                    with open(DATA_FILE, "w", encoding="utf-8") as f:
                        json.dump(result, f, ensure_ascii=False)
                except Exception:
                    pass
            elif self.path == "/export":
                data = export_xlsx(payload)
                self.send_response(200)
                self.send_header("Content-Type",
                                 "application/vnd.openxmlformats-officedocument."
                                 "spreadsheetml.sheet")
                self.send_header("Content-Disposition", 'attachment; filename="export.xlsx"')
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
            elif self.path == "/send":
                result = send_email(payload)
            elif self.path == "/sms":
                result = send_sms(payload.get("numbers") or [], payload.get("text") or "")
            else:
                self._send(404, json.dumps({"error": "not found"})); return
            self._send(200, json.dumps(result, ensure_ascii=False))
        except Exception as e:  # noqa: BLE001
            self._send(500, json.dumps({"error": f"{type(e).__name__}: {e}"}, ensure_ascii=False))


def find_free_port(pref):
    for p in (pref, pref + 1, pref + 2, 0):
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.bind(("127.0.0.1", p)); port = s.getsockname()[1]; s.close(); return port
        except OSError:
            s.close()
    return pref


PAGE = r"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>일일 입출고 리포트</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<!--CHARTJS-->
<style>
:root{--red:#c43a3a;--red2:#e15a5a;--blue:#3a6ea5;--green:#3f9d6b;--amber:#e0a93a;--ink:#15151a;--ink2:#23232b;--mut:#8a8a94;--line:#ededf1;--bg:#f4f4f7;}
*{box-sizing:border-box;}html{scroll-behavior:smooth;}
body{margin:0;background:var(--bg);color:var(--ink);font-family:'Pretendard',-apple-system,system-ui,sans-serif;-webkit-font-smoothing:antialiased;}
.tab{font-variant-numeric:tabular-nums;}
/* 업로드 화면 */
#land{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;
  background:radial-gradient(1200px 600px at 70% -10%,rgba(196,58,58,.18),transparent 60%),linear-gradient(160deg,#15151a,#23232b 60%,#3a2326);}
#land .box{width:100%;max-width:560px;text-align:center;color:#fff;}
#land h1{font-size:30px;font-weight:900;letter-spacing:-1px;margin:0 0 6px;}
#land h1 .r{color:var(--red2);}
#land p{color:#b9b9c4;font-size:13.5px;margin:0 0 26px;line-height:1.6;}
#drop{border:2.5px dashed rgba(255,255,255,.28);border-radius:18px;padding:42px 28px;cursor:pointer;transition:.18s;background:rgba(255,255,255,.03);}
#drop.hot{border-color:var(--red2);background:rgba(196,58,58,.12);}
#drop .ic{font-size:40px;}
#drop .t{font-size:16px;font-weight:700;margin-top:10px;}
#drop .h{font-size:12px;color:#aaa;margin-top:6px;}
.pwrow{margin-top:18px;display:flex;gap:8px;justify-content:center;align-items:center;color:#ccc;font-size:12.5px;}
.pwrow input{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:#fff;border-radius:8px;padding:7px 11px;width:120px;font-family:inherit;}
.loading{color:#fff;margin-top:22px;font-size:14px;display:none;}
.spin{display:inline-block;width:16px;height:16px;border:2.5px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:sp .8s linear infinite;vertical-align:-3px;margin-right:7px;}
@keyframes sp{to{transform:rotate(360deg);}}
.errmsg{color:#ff9b9b;margin-top:16px;font-size:13px;display:none;}
/* 대시보드 */
#dash{display:none;}
header{background:linear-gradient(135deg,#15151a,#23232b 55%,#3a2326);color:#fff;padding:22px 40px 20px;position:sticky;top:0;z-index:50;overflow:hidden;box-shadow:0 4px 18px rgba(0,0,0,.18);}
header::after{content:"";position:absolute;right:-80px;top:-90px;width:300px;height:300px;border-radius:50%;background:radial-gradient(circle,rgba(196,58,58,.4),transparent 70%);}
header .ttl{font-size:13px;color:#c9c9d2;font-weight:700;letter-spacing:3px;text-transform:uppercase;}
header h1{margin:6px 0 0;font-size:30px;font-weight:900;letter-spacing:-1px;position:relative;}
header h1 .d{color:var(--red2);}
header .meta{margin-top:8px;font-size:12.5px;color:#b9b9c4;}
header .reload{position:absolute;right:40px;top:30px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);color:#fff;font-size:12px;font-weight:600;padding:8px 15px;border-radius:9px;cursor:pointer;z-index:2;}
.wrap{max-width:1280px;margin:0 auto;padding:26px 40px 70px;}
.offseg{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px;align-items:center;}
.offseg .lab{font-size:11.5px;color:var(--mut);font-weight:700;margin-right:4px;}
.offseg button{border:1.5px solid var(--line);background:#fff;font-family:inherit;font-size:13px;font-weight:700;color:#777;padding:8px 15px;border-radius:10px;cursor:pointer;transition:.15s;}
.offseg button:hover{border-color:#cfcfd6;}
.offseg button.on{background:var(--ink);color:#fff;border-color:var(--ink);}
.offseg button.all{color:var(--red);border-color:#f1cccc;}
.offseg button.all.on{background:var(--red);color:#fff;border-color:var(--red);}
.dayseg{display:inline-flex;background:#e9e9ef;border-radius:12px;padding:4px;gap:4px;margin-bottom:20px;flex-wrap:wrap;}
.dayseg button{border:none;background:transparent;font-family:inherit;font-size:13.5px;font-weight:700;color:#777;padding:9px 18px;border-radius:9px;cursor:pointer;transition:.15s;}
.dayseg button.on{background:#fff;color:var(--ink);box-shadow:0 2px 8px rgba(0,0,0,.1);}
.dayseg button .tg{font-size:10px;font-weight:800;color:#fff;background:var(--red);border-radius:5px;padding:1px 6px;margin-left:6px;}
.kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:14px;margin-bottom:26px;}
@media(max-width:1080px){.kpis{grid-template-columns:repeat(3,1fr);}}
@media(max-width:620px){.kpis{grid-template-columns:repeat(2,1fr);}}
.kpi{background:#fff;border-radius:16px;padding:18px 18px;box-shadow:0 1px 2px rgba(0,0,0,.04),0 6px 22px rgba(0,0,0,.05);position:relative;overflow:hidden;transition:transform .15s;}
.kpi:hover{transform:translateY(-3px);}
.kpi::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--c,var(--ink));}
.kpi.in{--c:var(--blue);} .kpi.out{--c:var(--red);} .kpi.net{--c:var(--green);} .kpi.cu{--c:var(--amber);}
.kpi.old{--c:var(--red);} .kpi.av{--c:var(--green);} .kpi.bk{--c:var(--amber);} .kpi.it{--c:var(--blue);}
/* 뷰 전환 (입출고 / 재고) */
.xlsbtn{margin-bottom:4px;background:#1f7a4c;color:#fff;border:none;font-family:inherit;
  font-size:13px;font-weight:700;padding:11px 18px;border-radius:10px;cursor:pointer;}
.xlsbtn:hover{background:#186139;} .xlsbtn[disabled]{opacity:.5;cursor:default;}
.viewseg{display:inline-flex;background:#e9e9ef;border-radius:12px;padding:4px;gap:4px;margin-bottom:4px;}
.viewseg button{border:none;background:transparent;font-family:inherit;font-size:14px;font-weight:800;color:#777;padding:10px 22px;border-radius:9px;cursor:pointer;transition:.15s;}
.viewseg button.on{background:#fff;color:var(--ink);box-shadow:0 2px 8px rgba(0,0,0,.12);}
tr.oldrow td{background:#fdf1f1;}
td.old{color:var(--red);font-weight:800;}
.kpi .l{font-size:11.5px;color:var(--mut);font-weight:600;}
.kpi .v{font-size:25px;font-weight:900;margin-top:8px;letter-spacing:-.6px;line-height:1;}
.kpi .u{font-size:12px;font-weight:600;color:var(--mut);margin-left:2px;}
.kpi .d{font-size:11.5px;font-weight:700;margin-top:8px;}
.kpi .d.up{color:var(--green);} .kpi .d.dn{color:var(--red);} .kpi .d.fl{color:var(--mut);}
.hl{background:linear-gradient(110deg,#fff,#fff7f4);border:1px solid #f3dcdc;border-radius:16px;padding:16px 20px;margin-bottom:24px;display:flex;align-items:center;gap:16px;box-shadow:0 4px 18px rgba(196,58,58,.06);}
.hl .tag{background:var(--red);color:#fff;font-size:11px;font-weight:800;padding:5px 11px;border-radius:8px;white-space:nowrap;}
.hl .txt{font-size:14px;} .hl .txt b{font-weight:800;} .hl .txt .q{color:var(--red);font-weight:900;font-variant-numeric:tabular-nums;}
/* 오늘 요약 패널 — 카드 하단이 비지 않게 채운다 */
#splitcard{display:flex;flex-direction:column;}
.stable{border-top:1px solid var(--line);}
.srow{display:grid;grid-template-columns:106px 1fr auto;align-items:center;gap:10px;
  padding:11px 2px;border-bottom:1px solid var(--line);}
.srow .sl{font-size:12.5px;color:var(--mut);font-weight:600;}
.srow .sv{font-size:14px;font-variant-numeric:tabular-nums;}
.srow .sv b{font-weight:800;font-size:15.5px;}
.srow .ss{text-align:right;}
.sd{font-size:11.5px;font-weight:700;white-space:nowrap;}
.sd.up{color:var(--green);} .sd.dn{color:var(--red);} .sd.fl{color:var(--mut);}
.sbox{margin-top:14px;background:linear-gradient(110deg,#fdf6f5,#fff);border:1px solid #f0dcdc;
  border-radius:12px;padding:13px 15px;}
.sbox .sbt{font-size:11px;font-weight:800;color:var(--red);letter-spacing:.2px;
  display:flex;justify-content:space-between;align-items:center;}
.sbox .sbp{background:var(--red);color:#fff;border-radius:6px;padding:2px 7px;font-size:10.5px;}
.sbox .sbb{margin-top:7px;font-size:13px;color:#333;}
.sbox .sbb b{font-weight:800;}
.sbox .sbq{margin-top:5px;font-size:21px;font-weight:900;color:var(--ink);
  font-variant-numeric:tabular-nums;letter-spacing:-.4px;display:flex;align-items:baseline;gap:6px;}
.sbox .sbq span{font-size:12px;font-weight:700;color:var(--mut);}
.sbox .sbq .sbs{margin-left:auto;font-size:11.5px;font-weight:600;}
.schips{margin-top:auto;padding-top:14px;display:flex;align-items:center;gap:7px;flex-wrap:wrap;}
.schips .scl{font-size:11px;color:var(--mut);font-weight:700;margin-right:2px;}
.schip{font-size:12px;background:#f3f3f6;border-radius:7px;padding:5px 10px;color:#444;}
.schip b{font-weight:800;color:var(--ink);margin-left:3px;}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:22px;}
@media(max-width:900px){.grid{grid-template-columns:1fr;}}
.card{background:#fff;border-radius:16px;padding:20px 22px;box-shadow:0 1px 2px rgba(0,0,0,.04),0 6px 22px rgba(0,0,0,.05);}
.card h2{margin:0 0 4px;font-size:15.5px;font-weight:800;letter-spacing:-.3px;}
.card .desc{font-size:11.5px;color:var(--mut);margin:0 0 14px;}
.cbox{position:relative;height:260px;}
.tabs{display:flex;gap:8px;margin:6px 0 0;}
.tabbtn{font-size:13px;font-weight:700;padding:9px 18px;border:none;background:#ececf1;color:#666;border-radius:10px 10px 0 0;cursor:pointer;font-family:inherit;}
.tabbtn.on{background:#fff;color:var(--ink);box-shadow:0 -2px 8px rgba(0,0,0,.04);}
.tabbtn .n{font-size:11px;color:var(--mut);margin-left:5px;font-weight:600;}
.tablewrap{background:#fff;border-radius:0 16px 16px 16px;box-shadow:0 1px 2px rgba(0,0,0,.04),0 6px 22px rgba(0,0,0,.05);overflow:hidden;}
.scroll{max-height:520px;overflow:auto;}
table{width:100%;border-collapse:collapse;font-size:12.5px;}
thead th{position:sticky;top:0;background:#fafafb;color:var(--mut);font-weight:700;font-size:10.5px;text-transform:uppercase;letter-spacing:.4px;text-align:left;padding:11px 14px;border-bottom:1.5px solid var(--line);z-index:1;}
tbody td{padding:10px 14px;border-bottom:1px solid var(--line);}
tbody tr:hover{background:#fbfafc;}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;}
td.part{font-family:ui-monospace,Consolas,monospace;font-size:11.5px;}
td.qty{font-weight:800;font-variant-numeric:tabular-nums;}
.pill{display:inline-block;font-size:10.5px;font-weight:700;padding:2px 8px;border-radius:6px;background:#eef2f7;color:var(--blue);}
.pill.o{background:#fbe9e9;color:var(--red);}
.foot{text-align:center;color:#aab;font-size:11px;margin-top:18px;}
.rank{display:inline-flex;width:20px;height:20px;align-items:center;justify-content:center;background:var(--ink);color:#fff;border-radius:6px;font-size:10px;font-weight:800;margin-right:8px;}
.recip{display:flex;align-items:center;gap:12px;padding:8px 0;border-bottom:1px solid var(--line);}
.recip .nm{width:200px;flex:none;font-size:13px;font-weight:700;}
.recip input{flex:1;min-width:200px;font-size:13px;padding:9px 12px;border:1.5px solid var(--line);border-radius:9px;font-family:inherit;outline:none;transition:border .15s;}
.recip input:focus{border-color:var(--blue);}
.confirmbtn{background:var(--red);color:#fff;border:none;font-family:inherit;font-size:14px;font-weight:800;padding:12px 24px;border-radius:11px;cursor:pointer;transition:.15s;}
.confirmbtn:hover{background:#a93030;}
</style></head><body>

<div id="land">
  <div class="box">
    <h1>일일 입출고 <span class="r">리포트</span></h1>
    <p>영업실별 일일 입출고 엑셀을 <b style="color:#fff">한 번에 여러 개</b> 올리면, 실별로 각각 볼 수 있는 리포트로 정리합니다.<br>(예: 영업1,2실 / 3실 / 4실 / 5실 파일을 한꺼번에 — 실별 링크 /12 /3 /4 /5)</p>
    <div id="drop">
      <div class="ic">📥</div>
      <div class="t">엑셀 파일들을 여기로 끌어다 놓거나 클릭 (여러 개 가능)</div>
      <div class="h">.xlsx · 비밀번호 보호 지원 · 파일명의 (영업N실)로 실 구분</div>
      <input id="file" type="file" accept=".xlsx" multiple style="display:none">
    </div>
    <div class="pwrow"><span>🔒 비밀번호</span><input id="pw" type="text" placeholder="엑셀 비밀번호"></div>
    <div class="loading" id="loading"><span class="spin"></span>분석 중…</div>
    <div class="errmsg" id="errmsg"></div>
  </div>
</div>

<div id="dash">
  <header>
    <button class="reload" onclick="showUpload()">↻ 새 파일 업로드</button>
    <div class="ttl">Daily Shipping &amp; Receiving</div>
    <h1 id="h-date">—</h1>
    <div class="meta" id="h-meta"></div>
  </header>
  <div class="wrap" id="viewnav" style="display:none;padding-bottom:0">
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <div class="viewseg">
        <button class="on" id="vw-io" onclick="showView('io')">📦 일일 입출고</button>
        <button id="vw-inv" onclick="showView('inv')">📊 재고 현황</button>
      </div>
      <button id="btn-xls" class="xlsbtn" onclick="exportXlsx()">⬇ Excel 내보내기</button>
    </div>
  </div>

  <div class="wrap" id="ioview">
    <div class="offseg" id="offseg"></div>
    <div class="dayseg" id="dayseg"></div>
    <div class="kpis" id="kpis"></div>
    <div class="hl" id="hl" style="display:none"></div>
    <div class="grid">
      <div class="card"><h2>어제 vs 오늘 물동</h2><p class="desc">입고·출고 수량 비교</p><div class="cbox"><canvas id="cCompare"></canvas></div></div>
      <div class="card"><h2>오늘 출고 Top 거래처</h2><p class="desc">수량 기준 상위</p><div class="cbox"><canvas id="cCust"></canvas></div></div>
    </div>
    <div class="grid">
      <div class="card"><h2>담당자별 처리 건수</h2><p class="desc">오늘 입고+출고</p><div class="cbox"><canvas id="cSales"></canvas></div></div>
      <div class="card" id="splitcard"><h2>오늘 요약</h2><p class="desc">한눈에</p><div id="summary"></div></div>
    </div>
    <div class="tabs">
      <button class="tabbtn on" id="tb-out" onclick="showTab('out')">출고 내역<span class="n" id="n-out"></span></button>
      <button class="tabbtn" id="tb-in" onclick="showTab('in')">입고 내역<span class="n" id="n-in"></span></button>
    </div>
    <div class="tablewrap"><div class="scroll" id="tablearea"></div></div>

    <!-- 자동 발송 결과만 표시. 수동 발송(문자/이메일) UI는 없앴다. -->
    <div class="hl" id="mailnotice" style="display:none;margin-top:22px">
      <span class="tag">자동 발송</span>
      <div class="txt" id="emailstatus"></div>
    </div>

    <div class="foot" id="foot"></div>
  </div>

  <!-- ================= 재고 현황 ================= -->
  <div class="wrap" id="invview" style="display:none">
    <div class="offseg" id="inv-offseg"></div>
    <div class="kpis" id="inv-kpis"></div>
    <div class="hl" id="inv-hl" style="display:none"></div>
    <div class="grid">
      <div class="card"><h2 id="cAge-title">재고 노후화 (Datecode 연도별)</h2>
        <p class="desc" id="cAge-desc">—</p>
        <div class="cbox"><canvas id="cAge"></canvas></div></div>
      <div class="card"><h2 id="cFam-title">FAMILY별 재고</h2>
        <p class="desc" id="cFam-desc">수량 기준 상위</p>
        <div class="cbox"><canvas id="cFam"></canvas></div></div>
    </div>
    <div class="grid">
      <div class="card"><h2>당월 일별 입출고</h2><p class="desc">재고 시트 기준</p>
        <div class="cbox"><canvas id="cMon"></canvas></div></div>
      <div class="card"><h2>고객사 예약(booking)</h2><p class="desc">예약 수량 상위</p>
        <div class="cbox"><canvas id="cBook"></canvas></div></div>
    </div>
    <div class="grid" id="offgrid" style="display:none">
      <div class="card" style="grid-column:1/-1"><h2>영업실별 재고</h2>
        <p class="desc">실별 재고 수량 비교 · 클릭하면 해당 실로 이동</p>
        <div class="cbox"><canvas id="cOff"></canvas></div></div>
    </div>

    <div class="tabs">
      <button class="tabbtn on" id="ib-all" onclick="showInvTab('all')">전체 품목<span class="n" id="in-all"></span></button>
      <button class="tabbtn" id="ib-old" onclick="showInvTab('old')">장기재고<span class="n" id="in-old"></span></button>
      <button class="tabbtn" id="ib-bk" onclick="showInvTab('bk')">예약분<span class="n" id="in-bk"></span></button>
    </div>
    <div class="tablewrap">
      <div style="padding:12px 16px;border-bottom:1px solid var(--line)">
        <input id="invq" placeholder="Part# / MOBIS ID / FAMILY / 담당 검색…"
          style="width:100%;max-width:420px;font-size:13.5px;padding:9px 13px;border:1.5px solid var(--line);border-radius:10px;font-family:inherit;outline:none">
      </div>
      <div class="scroll" id="invtable"></div>
    </div>
    <div class="foot" id="inv-foot"></div>
  </div>
</div>

<script>
const fmt=n=>(n==null?'—':Number(n).toLocaleString('ko-KR'));
const esc=s=>String(s==null?'':s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const RED="#c43a3a",BLUE="#3a6ea5",GREEN="#3f9d6b",AMBER="#e0a93a",INK="#23232b";
if(window.Chart){Chart.defaults.font.family="'Pretendard',system-ui,sans-serif";Chart.defaults.color="#6b6b74";Chart.defaults.font.size=12;}
let DATA=null, O=null, OFFI=0, SERVER_MAIL=false;
let INVS={}, INVK=null, INVNAMES=[], VIEW='io';   // 실별 재고 / 선택된 실 / 볼 수 있는 실 / 현재 뷰

// 각 실 페이지(/12 /3 /4 /5)는 자기 실 것만 본다. 재고도 마찬가지.
// selectOffice 가 주소를 바꾸므로 시작할 때 한 번 붙잡아 둔다.
const URLSLUG=location.pathname.replace(/\//g,'');

// ── 고정 수신자 ──────────────────────────────────────────────
// 키는 실 이름에서 뽑은 숫자 (예: '영업4실'/'Inv4' → '4', '영업1,2실' → '12').
// 여기 없는 실은 DEFAULT_EMAILS 로 간다.
// 스팸함 통과 여부를 먼저 확인하는 중 — 당분간 본인에게만 보낸다.
// 확인되면 아래로 되돌린다:
//   const TEAM_EMAILS=['frankie@unitrontech.com','mh.choi@unitrontech.com',
//                      'yj.park@unitrontech.com','seanlee@unitrontech.com'].join(', ');
//   const OFFICE_EMAILS={'4':TEAM_EMAILS, '5':TEAM_EMAILS};
const OFFICE_EMAILS={};
const DEFAULT_EMAILS='seanlee@unitrontech.com';
function emailsFor(name){
  return parseEmails(OFFICE_EMAILS[officeSlug(name)]||DEFAULT_EMAILS);
}
// 서버에 보내는 메일계정(SMTP)이 설정돼 있으면 업로드 직후 자동 발송된다
fetch('/mailcfg').then(r=>r.json()).then(d=>{ SERVER_MAIL=!!d.configured; }).catch(()=>{});

// 업로드 (여러 파일)
const drop=document.getElementById('drop'),file=document.getElementById('file');
drop.onclick=()=>file.click();
['dragover','dragenter'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('hot');}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('hot');}));
drop.addEventListener('drop',ev=>{if(ev.dataTransfer.files.length)go(ev.dataTransfer.files);});
file.addEventListener('change',ev=>{if(ev.target.files.length)go(ev.target.files);});

// 파일명 → 실 이름.  (영업N실) 이 있으면 그걸 쓰고,
// 없으면 숫자를 뽑는다: Inv5/inv5_ → 영업5실, Inv12/inv12_ → 영업1,2실
function officeName(fn){
  const p=fn.match(/\(([^)]+)\)/);
  if(p) return p[1];
  const d=(fn.replace(/\.xlsx$/i,'').match(/\d+/)||[''])[0];
  if(d==='12') return '영업1,2실';
  return d?`영업${d}실`:fn.replace(/\.xlsx$/i,'');
}
function go(files){
  files=[...files].filter(f=>/\.xlsx$/i.test(f.name));
  if(!files.length) return;
  document.getElementById('errmsg').style.display='none';
  document.getElementById('loading').style.display='block';
  Promise.all(files.map(f=>new Promise((res,rej)=>{
    const rd=new FileReader();
    rd.onload=()=>res({name:officeName(f.name), file:rd.result}); rd.onerror=rej;
    rd.readAsDataURL(f);
  }))).then(arr=>{
    fetch('/build',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({files:arr, password:document.getElementById('pw').value})})
    .then(r=>r.json()).then(res=>{
      document.getElementById('loading').style.display='none';
      if(res.error){ const e=document.getElementById('errmsg'); e.textContent='오류: '+res.error; e.style.display='block'; return; }
      DATA=res; renderApp(); autoSend();
    }).catch(e=>{document.getElementById('loading').style.display='none';
      const el=document.getElementById('errmsg'); el.textContent='전송 오류: '+e; el.style.display='block';});
  }).catch(e=>{document.getElementById('loading').style.display='none';
    const el=document.getElementById('errmsg'); el.textContent='파일 읽기 오류: '+e; el.style.display='block';});
}

let CUR=0, TBcur='out', CH={}, TB={};

function delta(t,y){
  if(y==null) return '<span class="d fl">&nbsp;</span>';
  const d=t-y, p=y?Math.round(d/y*100):0;
  if(d===0) return '<span class="d fl">전일과 동일</span>';
  const cls=d>0?'up':'dn', ar=d>0?'▲':'▼';
  return `<span class="d ${cls}">${ar} ${fmt(Math.abs(d))} (${y?(d>0?'+':'')+p+'%':'신규'})</span>`;
}

function renderApp(){
  document.getElementById('land').style.display='none';
  document.getElementById('dash').style.display='block';
  document.getElementById('foot').textContent='자료: 사내 일일 입출고 엑셀 (날짜 시트) · 입고/출고 마스터 시트 미사용 · 파일명 (영업N실)로 실 구분';

  INVS=DATA.inventories||{};
  // 실 페이지에서는 자기 실 재고만 남긴다 — 다른 실 재고는 목록에도 없다.
  // 루트(/)에서만 실 선택 버튼이 나온다.
  INVNAMES=Object.keys(INVS);
  if(URLSLUG) INVNAMES=INVNAMES.filter(n=>officeSlug(n)===URLSLUG);

  const hasIO=DATA.offices && DATA.offices.length, hasInv=INVNAMES.length>0;
  document.getElementById('viewnav').style.display=(hasIO||hasInv)?'block':'none';
  document.getElementById('vw-io').style.display=hasIO?'':'none';
  document.getElementById('vw-inv').style.display=hasInv?'':'none';

  if(hasIO){
    let idx=0;
    if(URLSLUG){ const i=DATA.offices.findIndex(o=>officeSlug(o.name)===URLSLUG); if(i>=0) idx=i; }
    selectOffice(idx);
  }
  if(hasInv){
    // 재고는 실별로 보는 게 기본. '전체 합계'는 루트에서 버튼으로만 볼 수 있다.
    const pref=INVNAMES.filter(n=>n!=='전체 합계');
    INVK=pref[0]||INVNAMES[0];
    renderInventory();
  }
  showView(hasIO?'io':(hasInv?'inv':'io'));
}

function showView(v){
  const io=v==='io';
  VIEW=io?'io':'inv';
  document.getElementById('ioview').style.display=io?'block':'none';
  document.getElementById('invview').style.display=io?'none':'block';
  document.getElementById('vw-io').classList.toggle('on',io);
  document.getElementById('vw-inv').classList.toggle('on',!io);
  if(!io){
    const I=INVS[INVK];
    if(!I) return;
    document.getElementById('h-date').innerHTML=`재고 현황 <span class="d">·</span> ${esc(INVK)}`;
    document.getElementById('h-meta').textContent=
      `품목 ${fmt(I.n_items)}건 · 총 재고 ${fmt(I.total_qty)} EA`;
  }else if(O){ showDay(CUR); }
}

// ── Excel 내보내기 (지금 보고 있는 화면 그대로) ──
function exportXlsx(){
  const b=document.getElementById('btn-xls');
  const body = VIEW==='inv'
    ? {office:INVK, view:'inv', inventory:INVS[INVK]}
    : {office:O.name, view:'io', day:O.days[CUR]};
  const label = VIEW==='inv' ? `재고_${INVK}` : `입출고_${O.name}_${O.days[CUR].date}`;
  b.disabled=true; b.textContent='⏳ 만드는 중…';
  fetch('/export',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify(body)})
  .then(r=>{ if(!r.ok) throw new Error('HTTP '+r.status); return r.blob(); })
  .then(bl=>{
    const a=document.createElement('a');
    a.href=URL.createObjectURL(bl);
    a.download=label.replace(/[\\/:*?"<>|]/g,'_')+'.xlsx';
    a.click(); URL.revokeObjectURL(a.href);
    b.textContent='✅ 내려받음';
    setTimeout(()=>{b.textContent='⬇ Excel 내보내기'; b.disabled=false;},1800);
  })
  .catch(e=>{ b.textContent='❌ 실패: '+e.message;
    setTimeout(()=>{b.textContent='⬇ Excel 내보내기'; b.disabled=false;},2500); });
}
function buildOffseg(){
  // 각 실 페이지는 자기 실만 표시 — 다른 실로 가는 버튼 없음 (실 간 이동 불가)
  const seg=document.getElementById('offseg');
  seg.innerHTML=`<span class="lab">영업실</span><button class="on" style="cursor:default" disabled>${esc(O.name)}</button>`;
}
function selectOffice(i){
  OFFI=i; O=DATA.offices[i];
  history.replaceState(null,'','/'+officeSlug(O.name));
  buildOffseg();
  renderOffice();
}
function renderOffice(){
  document.getElementById('dayseg').innerHTML=O.dates.map((d,i)=>{
    const lab=i===O.today_idx?'오늘':(i===O.today_idx-1?'어제':'');
    return `<button onclick="showDay(${i})">${esc(d)}${lab?'<span class="tg">'+lab+'</span>':''}</button>`;
  }).join('');
  drawCompare();
  showDay(O.today_idx);
}

function officeSlug(name){ const d=(String(name).match(/\d/g)||[]).join(''); return d||'all'; }
function officeLink(){ return location.origin+'/'+officeSlug(O.name); }

// 로컬 테스트용 SMTP 설정(브라우저 저장). 배포본은 서버 환경변수를 쓰므로 보통 비어 있다.
function getSmtp(){ try{return JSON.parse(localStorage.getItem('smtp_config')||'{}');}catch(e){return {};} }

// ── 링크 이메일 발송 ──
function officeLinkFor(name){ return location.origin+'/'+officeSlug(name); }
function parseEmails(raw){ return [...new Set(String(raw||'').split(/[;,\s]+/).map(s=>s.trim()).filter(Boolean))]; }
function sendOfficeEmail(office, emails){
  const day=office.days[office.today_idx];
  return fetch('/send',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({office:office.name, date:day.date, link:officeLinkFor(office.name),
      emails, day, smtp:getSmtp()})}).then(r=>r.json());
}
// ── 업로드 완료 시 자동 발송 (전송 버튼을 누르지 않아도 나간다) ──
// 실별로 각자의 링크가 담긴 메일이 나간다. '전체 합계'는 실이 아니므로 제외.
function notice(msg){
  const box=document.getElementById('mailnotice');
  const st=document.getElementById('emailstatus');
  if(!box||!st) return;
  st.innerHTML=msg;
  box.style.display='flex';
}

async function autoSend(){
  const targets=(DATA.offices||[])
    .filter(o=>o.name!=='전체 합계')
    .map(o=>({o, emails:emailsFor(o.name)}))
    .filter(t=>t.emails.length);
  if(!targets.length) return;

  if(!SERVER_MAIL && !getSmtp().host){
    notice('⚠️ 자동 발송 안 됨 — 서버에 메일 계정(SMTP)이 설정돼 있지 않습니다.');
    return;
  }
  notice(`업로드 완료 — 자동 발송 중… (${targets.length}개 실)`);

  const ok=[], fail=[];
  for(const t of targets){
    try{
      const r=await sendOfficeEmail(t.o, t.emails);
      (r && r.ok && r.sent ? ok : fail).push(t.o.name+(r&&r.error?(' ('+r.error+')'):''));
    }catch(e){ fail.push(t.o.name+' ('+e+')'); }
  }
  notice(
    (ok.length?`✅ 자동 발송 완료 — <b>${ok.join(', ')}</b>`:'')+
    (fail.length?`${ok.length?' · ':''}❌ 실패: ${fail.join(', ')}`:''));
}

function showUpload(){
  document.getElementById('dash').style.display='none';
  document.getElementById('land').style.display='flex';
}
// 저장된 데이터가 있으면 업로드 없이 바로 보여줌 (링크 공유용)
fetch('/data').then(r=>r.json()).then(d=>{
  if(d && d.offices && d.offices.length){ DATA=d; renderApp(); }
}).catch(()=>{});

function showDay(idx){
  CUR=idx;
  const day=O.days[idx], prev=idx>0?O.days[idx-1]:null, k=day.kpi, pk=prev?prev.kpi:null;
  document.querySelectorAll('#dayseg button').forEach((b,i)=>b.classList.toggle('on',i===idx));
  const tag=idx===O.today_idx?'오늘':(idx===O.today_idx-1?'어제':'선택일');
  document.getElementById('h-date').innerHTML=`${esc(O.name)} <span class="d">·</span> ${esc(day.date)} ${tag}`;
  document.getElementById('h-meta').textContent=`${prev?'비교 기준 '+prev.date+' · ':'이전일 데이터 없음 · '}입고 ${fmt(k.in_cnt)}건 / 출고 ${fmt(k.out_cnt)}건`;
  document.getElementById('kpis').innerHTML=[
    ['in','입고 건수',k.in_cnt,'건',pk?pk.in_cnt:null],
    ['in','입고 수량',k.in_qty,'EA',pk?pk.in_qty:null],
    ['out','출고 건수',k.out_cnt,'건',pk?pk.out_cnt:null],
    ['out','출고 수량',k.out_qty,'EA',pk?pk.out_qty:null],
    ['net','순물동(입-출)',k.net,'EA',null],
    ['cu','거래처(출고)',k.customers,'곳',null],
  ].map(([c,l,v,u,y])=>`<div class="kpi ${c}"><div class="l">${l}</div>
     <div class="v tab">${fmt(v)}<span class="u">${u}</span></div>${delta(v,y)}</div>`).join('');

  const hlEl=document.getElementById('hl');
  if(day.highlight){const h=day.highlight; hlEl.style.display='flex';
    hlEl.innerHTML=`<span class="tag">${tag} 최대 출고</span>
      <div class="txt"><b>${esc(h.customer)}</b> 에 <b>${esc(h.part)}</b> <span class="q">${fmt(h.qty)}</span> EA 출고 · 담당 ${esc(h.sales)}</div>`;
  } else hlEl.style.display='none';

  buildSummary(day, prev, tag);
  drawDayCharts(day); buildTables(day); showTab(TBcur);
}

function drawCompare(){
  if(CH.c) CH.c.destroy();
  const labels=O.dates.map((d,i)=>d+(i===O.today_idx?' (오늘)':(i===O.today_idx-1?' (어제)':'')));
  CH.c=new Chart(document.getElementById('cCompare'),{type:'bar',
    data:{labels,datasets:[
      {label:'입고',data:O.days.map(d=>d.kpi.in_qty),backgroundColor:BLUE,borderRadius:6,maxBarThickness:46},
      {label:'출고',data:O.days.map(d=>d.kpi.out_qty),backgroundColor:RED,borderRadius:6,maxBarThickness:46},
    ]},options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'bottom',labels:{usePointStyle:true,boxWidth:8,padding:16}},
        tooltip:{callbacks:{label:c=>c.dataset.label+': '+fmt(c.raw)+' EA'}}},
      scales:{y:{grid:{color:'#f0f0f3'},ticks:{callback:v=>fmt(v)}},x:{grid:{display:false}}}}});
}

function drawDayCharts(day){
  if(CH.cu) CH.cu.destroy(); if(CH.s) CH.s.destroy();
  const cx=document.getElementById('cCust').getContext('2d');
  const g=cx.createLinearGradient(0,0,0,260);g.addColorStop(0,'rgba(196,58,58,.95)');g.addColorStop(1,'rgba(225,90,90,.6)');
  CH.cu=new Chart(cx,{type:'bar',
    data:{labels:day.cust.map(c=>c.name),datasets:[{data:day.cust.map(c=>c.qty),backgroundColor:g,borderRadius:5,maxBarThickness:22}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmt(c.raw)+' EA ('+day.cust[c.dataIndex].cnt+'건)'}}},
      scales:{x:{grid:{color:'#f0f0f3'},ticks:{callback:v=>fmt(v)}},y:{grid:{display:false}}}}});
  CH.s=new Chart(document.getElementById('cSales'),{type:'bar',
    data:{labels:day.sales.map(s=>s.name),datasets:[{data:day.sales.map(s=>s.cnt),backgroundColor:GREEN,borderRadius:5,maxBarThickness:22}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>c.raw+'건'}}},
      scales:{x:{grid:{color:'#f0f0f3'},ticks:{precision:0}},y:{grid:{display:false}}}}});
}

// ── 오늘 요약 (카드를 채우는 요약 패널) ──
function sRow(label, value, sub){
  return `<div class="srow"><div class="sl">${label}</div>
    <div class="sv">${value}</div><div class="ss">${sub||''}</div></div>`;
}
function pctOf(a,b){ return b?Math.round(a/b*100):0; }

function buildSummary(day, prev, tag){
  const k=day.kpi, pk=prev?prev.kpi:null;
  const net=k.net;
  const d=(v,y)=>{
    if(y==null||y===undefined) return '';
    const diff=v-y;
    if(!y && !v) return '<span class="sd fl">변동 없음</span>';
    if(!y) return '<span class="sd up">신규</span>';
    const p=Math.round(Math.abs(diff)/Math.abs(y)*100);
    const cls=diff>0?'up':(diff<0?'dn':'fl');
    const ar=diff>0?'▲':(diff<0?'▼':'·');
    return `<span class="sd ${cls}">${ar} ${fmt(Math.abs(diff))} (${diff>0?'+':''}${diff<0?'-':''}${p}%)</span>`;
  };

  const h=day.highlight;
  const share=h&&k.out_qty?pctOf(h.qty,k.out_qty):0;
  const top=(day.cust||[])[0];
  const sales=(day.sales||[]).slice(0,3);

  document.getElementById('summary').innerHTML=`
    <div class="stable">
      ${sRow('📦 입고', `${fmt(k.in_cnt)}건 · <b>${fmt(k.in_qty)}</b> EA`, d(k.in_qty, pk?pk.in_qty:null))}
      ${sRow('🚚 출고', `${fmt(k.out_cnt)}건 · <b>${fmt(k.out_qty)}</b> EA`, d(k.out_qty, pk?pk.out_qty:null))}
      ${sRow('⚖️ 순물동', `<span style="color:${net>=0?GREEN:RED};font-weight:800">${net>=0?'+':''}${fmt(net)}</span> EA`,
             net>=0?'<span class="sd up">순입고</span>':'<span class="sd dn">순출고</span>')}
      ${sRow('🏢 출고 거래처', `<b>${fmt(k.customers)}</b>곳`,
             top?`<span class="sd fl">1위 ${esc(top.name)} ${pctOf(top.qty,k.out_qty)}%</span>`:'')}
    </div>
    ${h?`<div class="sbox">
      <div class="sbt">${tag} 최대 출고 <span class="sbp">전체의 ${share}%</span></div>
      <div class="sbb"><b>${esc(h.customer)}</b> · ${esc(h.part)}</div>
      <div class="sbq">${fmt(h.qty)} <span>EA</span> <span class="sbs">담당 ${esc(h.sales)||'—'}</span></div>
    </div>`:''}
    ${sales.length?`<div class="schips"><span class="scl">담당 처리 건수</span>
      ${sales.map(s=>`<span class="schip">${esc(s.name)||'—'} <b>${s.cnt}</b></span>`).join('')}
    </div>`:''}`;
}

function emptyMsg(){return '<div style="padding:40px;text-align:center;color:#aaa;font-size:13px">이 날짜의 내역이 없습니다</div>';}
function buildTables(day){
  document.getElementById('n-out').textContent=day.out_rows.length;
  document.getElementById('n-in').textContent=day.in_rows.length;
  const out=!day.out_rows.length?emptyMsg():`<table><thead><tr><th>#</th><th>거래처</th><th>PART#</th><th class="n">수량</th><th>담당</th><th>문서번호</th><th>비고</th></tr></thead><tbody>${
    day.out_rows.map((r,i)=>`<tr><td class="n">${i+1}</td><td><b>${esc(r.customer)||'—'}</b></td><td class="part">${esc(r.part)}</td>
      <td class="qty">${fmt(r.qty)}</td><td>${esc(r.sales)}</td>
      <td>${esc(r.doc)||'—'}</td><td style="color:#888">${esc(r.remark)||''}</td></tr>`).join('')}</tbody></table>`;
  const inn=!day.in_rows.length?emptyMsg():`<table><thead><tr><th>#</th><th>거래처/공급</th><th>PART#</th><th class="n">수량</th><th>담당</th><th>FAB</th><th>비고</th></tr></thead><tbody>${
    day.in_rows.map((r,i)=>`<tr><td class="n">${i+1}</td><td><b>${esc(r.customer)||'—'}</b></td><td class="part">${esc(r.part)}</td>
      <td class="qty">${fmt(r.qty)}</td><td>${esc(r.sales)}</td><td>${esc(r.fab)?'<span class="pill">FAB '+esc(r.fab)+'</span>':'—'}</td>
      <td style="color:#888">${esc(r.remark)||''}</td></tr>`).join('')}</tbody></table>`;
  TB={out,in:inn};
}
function showTab(which){
  TBcur=which;
  document.getElementById('tb-out').classList.toggle('on',which==='out');
  document.getElementById('tb-in').classList.toggle('on',which==='in');
  document.getElementById('tablearea').innerHTML=TB[which]||'';
}

// ================= 재고 현황 =================
let ICH={}, ITAB='all';
const pct=(a,b)=>b?Math.round(a/b*1000)/10:0;

// 볼 수 있는 실만 연다 (실 페이지에서는 자기 실뿐 — 차트 드릴다운으로도 못 넘어간다)
function selectInv(name){
  if(!INVNAMES.includes(name)) return;
  INVK=name; renderInventory(); showView('inv');
}

function renderInventory(){
  const I=INVS[INVK];
  if(!I) return;

  // 실 선택 버튼. 실 페이지에서는 볼 수 있는 실이 자기 실 하나뿐이라 버튼도 하나 (누를 수 없음).
  const seg=document.getElementById('inv-offseg');
  if(INVNAMES.length<2){
    seg.innerHTML=`<span class="lab">영업실</span>`+
      `<button class="on" style="cursor:default" disabled>${esc(INVK)}</button>`;
  }else{
    const names=INVNAMES.slice().sort((a,b)=>
      (a==='전체 합계'?-1:0)-(b==='전체 합계'?-1:0));
    seg.innerHTML='<span class="lab">영업실</span>'+names.map(n=>{
      const all=n==='전체 합계'?' all':'';
      const on=n===INVK?' on':'';
      return `<button class="${all}${on}" onclick="selectInv('${n.replace(/'/g,"\\'")}')">${esc(n)}</button>`;
    }).join('');
  }

  document.getElementById('inv-foot').textContent=
    `자료: 재고 엑셀의 '${I.sheet}' 시트 · 재고 수량이 있는 품목만 집계 (합계 행 제외) · 장기재고 = Datecode ${I.old_year}년 이전`;

  // 노후화(Datecode)는 영업1,2실에만 데이터가 있다. 없는 실에서는 관련 요소를 아예 뺀다.
  const hasDC=I.datecode.some(d=>d.qty>0);
  document.getElementById('inv-kpis').innerHTML=[
    ['it','재고 품목',fmt(I.n_items),'건'],
    ['in','총 재고',fmt(I.total_qty),'EA'],
    ['av','가용 재고',fmt(I.avail_qty),'EA'],
    ['bk','예약(booking)',fmt(I.booking_qty),'EA'],
    hasDC ? ['old',`장기재고 (${I.old_year}년 이전)`,fmt(I.old_qty),'EA']
          : ['it','당월 입고',fmt(I.month.inbound),'EA'],
    ['cu','당월 출고',fmt(I.month.outbound),'EA'],
  ].map(([c,l,v,u])=>`<div class="kpi ${c}"><div class="l">${l}</div>
     <div class="v tab">${v}<span class="u">${u}</span></div></div>`).join('');

  const hl=document.getElementById('inv-hl');
  if(hasDC && I.old_qty>0){
    hl.style.display='flex';
    hl.innerHTML=`<span class="tag">장기재고</span><div class="txt">
      Datecode <b>${I.old_year}년 이전</b> 재고가 <span class="q">${fmt(I.old_qty)} EA</span>
      — 전체 재고의 <b>${pct(I.old_qty,I.total_qty)}%</b></div>`;
  }else hl.style.display='none';

  Object.values(ICH).forEach(c=>c&&c.destroy()); ICH={};

  // 재고 노후화 — 장기재고는 빨강
  // Datecode 는 영업1,2실에만 채워져 있다(실측). 없는 실은 담당별 재고로 대체한다.
  const dc=I.datecode.filter(d=>d.qty>0);
  const logY={type:'logarithmic',grid:{color:'#f0f0f3'},
              ticks:{callback:v=>{const s=String(v);
                return /^[125]0*$/.test(s)?fmt(v):'';}}};
  const T=document.getElementById('cAge-title'), DSC=document.getElementById('cAge-desc');

  if(dc.length){
    T.textContent='재고 노후화 (Datecode 연도별)';
    DSC.innerHTML=`<b>${I.old_year}년 이전 = 장기재고</b> (빨강) · 편차가 커서 <b>로그 스케일</b> — 막대 길이를 그대로 비교하지 마세요`;
    ICH.age=new Chart(document.getElementById('cAge'),{type:'bar',
      data:{labels:dc.map(d=>d.year),datasets:[{data:dc.map(d=>d.qty),
        backgroundColor:dc.map(d=>d.year<=I.old_year?RED:BLUE),borderRadius:6,maxBarThickness:52}]},
      options:{responsive:true,maintainAspectRatio:false,
        plugins:{legend:{display:false},tooltip:{callbacks:{
          label:c=>fmt(c.raw)+' EA ('+dc[c.dataIndex].items+'품목) · 전체의 '
                   +pct(c.raw,I.total_qty)+'%'}}},
        scales:{y:logY,x:{grid:{display:false}}}}});
  }else{
    // 담당(SALES)별 재고 → 그것도 없으면 재고 상위 품목
    const bs={};
    I.items.forEach(x=>{ const s=(x.sales||'').trim();
      if(s && s!=='.') bs[s]=(bs[s]||0)+x.qty; });
    const sal=Object.entries(bs).sort((a,b)=>b[1]-a[1]).slice(0,8);

    if(sal.length){
      T.textContent='담당자별 재고';
      DSC.innerHTML='수량 기준 상위';
      ICH.age=new Chart(document.getElementById('cAge'),{type:'bar',
        data:{labels:sal.map(s=>s[0]),datasets:[{data:sal.map(s=>s[1]),
          backgroundColor:BLUE,borderRadius:5,maxBarThickness:22}]},
        options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
          plugins:{legend:{display:false},tooltip:{callbacks:{
            label:c=>fmt(c.raw)+' EA · 전체의 '+pct(c.raw,I.total_qty)+'%'}}},
          scales:{x:{grid:{color:'#f0f0f3'},ticks:{callback:v=>fmt(v)}},
                  y:{grid:{display:false}}}}});
    }else{
      const top=[...I.items].sort((a,b)=>b.qty-a.qty).slice(0,8);
      T.textContent='재고 상위 품목';
      DSC.innerHTML='수량 기준 상위';
      ICH.age=new Chart(document.getElementById('cAge'),{type:'bar',
        data:{labels:top.map(x=>x.part),datasets:[{data:top.map(x=>x.qty),
          backgroundColor:BLUE,borderRadius:5,maxBarThickness:22}]},
        options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
          plugins:{legend:{display:false},tooltip:{callbacks:{
            label:c=>fmt(c.raw)+' EA · 전체의 '+pct(c.raw,I.total_qty)+'%'}}},
          scales:{x:{grid:{color:'#f0f0f3'},ticks:{callback:v=>fmt(v)}},
                  y:{grid:{display:false},ticks:{font:{size:10}}}}}});
    }
  }

  // FAMILY 는 영업1,2실만 채워져 있다. 비어 있으면 VENDER 로 자동 전환한다.
  const named=a=>a.filter(x=>x.name && x.name!=='(미지정)' && x.name!=='.');
  const useFam=named(I.by_family).length>0;
  const dim=useFam?named(I.by_family):named(I.by_vender);
  const dimName=useFam?'FAMILY':'VENDER';
  const unk=(useFam?I.by_family:I.by_vender).find(x=>!x.name||x.name==='(미지정)'||x.name==='.');
  document.getElementById('cFam-title').textContent=`${dimName}별 재고`;
  document.getElementById('cFam-desc').innerHTML=
    `수량 기준 상위` + (unk?` · <b>미분류 ${fmt(unk.qty)} EA (${unk.items}품목)</b> 는 제외`:'')
    + (useFam?'':' · FAMILY 가 비어 있어 VENDER 로 표시');

  const fam=dim.slice(0,8);
  ICH.fam=new Chart(document.getElementById('cFam'),{type:'bar',
    data:{labels:fam.map(f=>f.name),datasets:[{data:fam.map(f=>f.qty),
      backgroundColor:BLUE,borderRadius:5,maxBarThickness:22}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{
        label:c=>fmt(c.raw)+' EA ('+fam[c.dataIndex].items+'품목)'}}},
      scales:{x:{grid:{color:'#f0f0f3'},ticks:{callback:v=>fmt(v)}},y:{grid:{display:false}}}}});

  const days=I.month.daily_in.map((_,i)=>i+1);
  ICH.mon=new Chart(document.getElementById('cMon'),{type:'bar',
    data:{labels:days,datasets:[
      {label:'입고',data:I.month.daily_in,backgroundColor:BLUE,borderRadius:3,maxBarThickness:14},
      {label:'출고',data:I.month.daily_out,backgroundColor:RED,borderRadius:3,maxBarThickness:14},
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'bottom',labels:{usePointStyle:true,boxWidth:8,padding:16}},
        tooltip:{callbacks:{label:c=>c.dataset.label+': '+fmt(c.raw)+' EA'}}},
      scales:{y:{grid:{color:'#f0f0f3'},ticks:{callback:v=>fmt(v)}},x:{grid:{display:false}}}}});

  const bk=I.customers.slice(0,8);
  ICH.book=new Chart(document.getElementById('cBook'),{type:'bar',
    data:{labels:bk.map(c=>c.name),datasets:[{data:bk.map(c=>c.booking),
      backgroundColor:AMBER,borderRadius:5,maxBarThickness:22}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmt(c.raw)+' EA'}}},
      scales:{x:{grid:{color:'#f0f0f3'},ticks:{callback:v=>fmt(v)}},y:{grid:{display:false}}}}});

  // 영업실별 재고 비교 — 전체 합계 화면에서만
  const offs=Object.keys(INVS).filter(n=>n!=='전체 합계');
  const grid=document.getElementById('offgrid');
  if(INVK==='전체 합계' && offs.length>1){
    grid.style.display='';
    ICH.off=new Chart(document.getElementById('cOff'),{type:'bar',
      data:{labels:offs,datasets:[
        {label:'가용',data:offs.map(n=>INVS[n].avail_qty),backgroundColor:GREEN,
         borderRadius:6,maxBarThickness:60},
        {label:'예약',data:offs.map(n=>INVS[n].booking_qty),backgroundColor:AMBER,
         borderRadius:6,maxBarThickness:60},
      ]},
      options:{responsive:true,maintainAspectRatio:false,
        onClick:(e,el)=>{ if(el.length) selectInv(offs[el[0].index]); },
        plugins:{legend:{position:'bottom',labels:{usePointStyle:true,boxWidth:8,padding:16}},
          tooltip:{callbacks:{
            afterBody:c=>{const n=offs[c[0].dataIndex];
              return `품목 ${fmt(INVS[n].n_items)}건 · 총 재고 ${fmt(INVS[n].total_qty)} EA`;},
            label:c=>c.dataset.label+': '+fmt(c.raw)+' EA'}}},
        scales:{x:{stacked:true,grid:{display:false}},
                y:{stacked:true,type:'logarithmic',grid:{color:'#f0f0f3'},
                   ticks:{callback:v=>{const s=String(v);
                     return /^[125]0*$/.test(s)?fmt(v):'';}}}}}});
  }else{ grid.style.display='none'; }

  document.getElementById('in-all').textContent=I.items.length;
  document.getElementById('in-old').textContent=I.items.filter(x=>x.old>0).length;
  // 장기재고 탭은 Datecode 가 있는 실(영업1,2실)에서만 보여준다
  document.getElementById('ib-old').style.display=hasDC?'':'none';
  document.getElementById('in-bk').textContent=I.items.filter(x=>x.booking>0).length;
  if(!hasDC && ITAB==='old') ITAB='all';
  document.getElementById('invq').oninput=()=>drawInvTable();
  showInvTab('all');
}

function showInvTab(t){
  ITAB=t;
  ['all','old','bk'].forEach(k=>document.getElementById('ib-'+k).classList.toggle('on',k===t));
  drawInvTable();
}

function drawInvTable(){
  const I=INVS[INVK];
  if(!I) return;
  const q=(document.getElementById('invq').value||'').trim().toUpperCase();
  let rows=I.items;
  if(ITAB==='old') rows=rows.filter(x=>x.old>0);
  if(ITAB==='bk')  rows=rows.filter(x=>x.booking>0);
  if(q) rows=rows.filter(x=>[x.part,x.mobis,x.family,x.sales,x.customer,x.pn]
      .some(v=>String(v||'').toUpperCase().includes(q)));
  rows=[...rows].sort((a,b)=>b.qty-a.qty);

  const el=document.getElementById('invtable');
  if(!rows.length){ el.innerHTML=emptyMsg(); return; }
  // 장기재고/Datecode 칼럼은 데이터가 있는 실에서만 (영업1,2실)
  const dc=I.datecode.some(d=>d.qty>0);
  el.innerHTML=`<table><thead><tr><th>#</th><th>PART#</th><th>MOBIS ID</th><th>FAMILY</th>
    <th>실</th><th class="n">재고</th><th class="n">가용</th><th class="n">예약</th>
    ${dc?'<th class="n">장기재고</th><th>Datecode</th>':''}<th>담당</th></tr></thead><tbody>${
    rows.map((x,i)=>`<tr class="${dc&&x.old>0?'oldrow':''}">
      <td class="n">${i+1}</td>
      <td class="part">${esc(x.part)}</td>
      <td>${esc(x.mobis)||'—'}</td>
      <td>${esc(x.family)||'—'}</td>
      <td>${esc(x.office)||'—'}</td>
      <td class="qty">${fmt(x.qty)}</td>
      <td class="n">${fmt(x.avail)}</td>
      <td class="n">${x.booking?fmt(x.booking):'—'}</td>
      ${dc?`<td class="n ${x.old>0?'old':''}">${x.old?fmt(x.old):'—'}</td>
      <td>${x.oldest?('<span class="pill">'+x.oldest+'~</span>'):'—'}</td>`:''}
      <td>${esc(x.sales)||'—'}</td></tr>`).join('')}</tbody></table>`;
}
</script></body></html>"""


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="일일 입출고 리포트 대시보드")
    ap.add_argument("--port", type=int, default=8780)
    ap.add_argument("--no-open", action="store_true")
    args = ap.parse_args()
    env_port = os.environ.get("PORT")
    if env_port:                                   # 클라우드(Render 등): 0.0.0.0 + $PORT
        host, port = "0.0.0.0", int(env_port)
        args.no_open = True
    else:                                          # 로컬: localhost + 빈 포트 자동
        host, port = "127.0.0.1", find_free_port(args.port)
    url = f"http://localhost:{port}/"
    httpd = ThreadingHTTPServer((host, port), Handler)
    print(f"\n일일 입출고 리포트 → {url} (bind {host}:{port})\n종료: Ctrl+C")
    if not args.no_open:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n종료."); httpd.shutdown()
