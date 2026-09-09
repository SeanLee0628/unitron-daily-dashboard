#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""일일 입출고 리포트 대시보드 (업로드형).

암호화된 'daily material shipping & receiving' 엑셀을 업로드하면 → 일일 입출고 대시보드.

  하루 보기  : 날짜 시트(YYYY-MM-DD, 어제·오늘) → KPI·차트·표
  기간 조회  : 누적 '입고'/'출고' 시트(2019~) → SQLite 에 적재하고 기간·검색으로 조회
  재고 현황  : '~ inventory' 시트

표준 라이브러리만 사용. 엑셀은 브라우저에서 base64로 전송한다.
"""
from __future__ import annotations

import base64
import datetime
import io
import json
import os
import re
import socket
import sqlite3
import math
import threading
import urllib.parse
import webbrowser
from collections import Counter, defaultdict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import msoffcrypto
import openpyxl                                   # Excel 내보내기에만 쓴다 (쓰기)

import fastxl                                     # 읽기 — openpyxl 보다 11배 빠름

HERE = os.path.dirname(os.path.abspath(__file__))
# 엑셀 비밀번호는 서버가 들고 있어서 업로드할 때 칠 필요가 없다.
# 코드에는 두지 않는다 — Render Environment 의 XLSX_PW 에 넣는다 (로컬은 같은 이름의 환경변수).
# 페이지로도 내려보내지 않는다. 공개 URL이라 브라우저에 실으면 그대로 샌다.
DEFAULT_PW = os.environ.get("XLSX_PW", "")
CHART_JS = os.path.join(HERE, "chart.umd.min.js")
DATA_FILE = os.path.join(os.environ.get("DATA_DIR", HERE), "saved_data.json")  # 데이터 저장(공유)
LOG_DB = os.path.join(os.environ.get("DATA_DIR", HERE), "log.db")             # 입출고 이력(기간 조회)
MEMO_VER = 2                    # 재고 메모 저장 형식 (2 = booking 칸 메모만)


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


def memo_rows(ws):
    """시트의 셀 메모를 행 단위로 모은다 → {엑셀 행번호: [(칸번호, 내용)]}.

    메모는 셀 하나에 붙는데 화면은 행 단위라, 같은 행에 여러 개 달려 있으면 이어 붙인다.
    """
    out = defaultdict(list)
    try:
        for (rn, ci), txt in (getattr(ws, "comments", None) or {}).items():
            if txt:
                out[rn].append((ci, txt))
    except Exception as e:                          # 메모 때문에 업로드가 죽으면 안 된다
        print(f"[메모] {getattr(ws, 'title', '?')} 읽기 실패 — 건너뜀: {e}")
    for v in out.values():
        v.sort()
    return out


def memo_text(mrows, rn, hdr=None, only=None):
    """그 행의 메모를 '칸이름: 내용' 으로 붙여 한 줄로. 칸이름을 모르면 내용만.

    only 에 칸번호를 주면 그 칸에 달린 메모만, 칸이름 접두어 없이 낸다.
    재고 시트는 booking 칸(S열) 메모만 쓰기로 했다 — 행 전체를 이어 붙이면
    상관없는 칸 메모까지 같이 나온다.
    """
    got = mrows.get(rn)
    if not got:
        return ""
    if only is not None:
        return " / ".join(t for ci, t in got if ci == only) if only >= 0 else ""
    parts = []
    for ci, txt in got:
        h = ""
        if hdr and ci < len(hdr):
            h = re.sub(r"\s+", " ", clean(hdr[ci])).strip()
        parts.append(f"{h}: {txt}" if h else txt)
    return " / ".join(parts)


def parse_daily(ws):
    """날짜 시트 1개 → (inbound[], outbound[]). 컬럼은 헤더로 동적 인식(실마다 배치 달라도 OK)."""
    mode = None
    cols = {}
    hdr = []
    mrows = memo_rows(ws)
    inbound, outbound = [], []
    # iter_rows 는 빈 행도 채워서 순서대로 내므로 인덱스가 곧 엑셀 행번호다
    for rn, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if not r:
            continue
        c0 = clean(r[0])
        if c0 and not c0.isdigit():
            if c0 == "NO":                          # 헤더행 → 컬럼 위치 매핑
                cols = {_hkey(h): i for i, h in enumerate(r) if clean(h)}
                hdr = list(r)
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
                memo=memo_text(mrows, rn, hdr),
            ))
        elif mode == "out":
            outbound.append(dict(
                customer=col("CUSTOMER"), part=part, qty=qty, sales=sales,
                doc=col("문서번호", "DOC"), remark=col("REMARK"), vendor=col("VENDOR"),
                # 운송장번호는 날짜 시트에만 있다 (누적 '출고' 시트에는 이 칸이 없음)
                waybill=col("운송장번호", "운송장", "WAYBILL", "TRACKINGNUMBER"),
                memo=memo_text(mrows, rn, hdr),
            ))
    return inbound, outbound


def parse_workbook(wb):
    """워크북 → {날짜: (inbound, outbound)} (날짜 시트만)."""
    per = {}
    for s in wb.sheetnames:
        if re.match(r"\d{4}-\d{2}-\d{2}", s):
            per[s] = parse_daily(wb[s])
    return per


# ---------------------------------------------------------------- 입출고 이력 (기간 조회)
# 날짜 시트(YYYY-MM-DD)는 오늘·어제만 남기는 롤링 구조라 기간 조회에 못 쓴다.
# 같은 워크북의 '입고'/'출고' 누적 시트에 2019년부터의 전 이력이 그대로 쌓여 있다 —
# 실당 6만 행이라 브라우저로 통째 내려보낼 수 없어서 SQLite 에 넣고 서버가 걸러 준다.
LOG_DIRS = {"in": ("입고", "INBOUND"), "out": ("출고", "OUTBOUND")}
LOG_MAX_ROWS = 3000              # 화면에 한 번에 내려보낼 행 상한 (집계는 전체 기준으로 낸다)
# 전사 출고 원장(shipping management)은 실 소속이 없다. 실별 이력과 섞이면 같은 출고가
# 두 번 잡히므로 자기만의 slug 를 쓰고, '전체 합계' 집계에서도 빠진다.
LEDGER = "전체 출고 원장"
LEDGER_SLUG = "sm"


def norm_date(v):
    """엑셀 DATE 셀 → 'YYYY-MM-DD'. 못 읽으면 빈 문자열."""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.isoformat()
    m = re.match(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", clean(v))
    return "%04d-%02d-%02d" % tuple(int(x) for x in m.groups()) if m else ""


def log_sheets(wb):
    """누적 '입고'/'출고' 시트 이름 → {'in': 시트명, 'out': 시트명}. 날짜 시트는 뺀다."""
    found = {}
    for s in wb.sheetnames:
        if re.match(r"\d{4}-\d{2}-\d{2}", s):
            continue
        k = clean(s).upper().replace(" ", "")
        for d, names in LOG_DIRS.items():
            if d not in found and any(n in k for n in names):
                found[d] = s
    return found


def parse_log(ws):
    """누적 시트 1개 → [{date, part, qty, customer, sales, doc, mcode, remark, fab}].

    parse_daily 와 같은 헤더 동적 인식이되 DATE 를 함께 읽는다 — 날짜 시트는
    시트 이름이 곧 날짜였지만 여기서는 행마다 날짜가 다르다.
    """
    cols, rows = {}, []
    for r in ws.iter_rows(values_only=True):
        if not r:
            continue
        c0 = clean(r[0])
        if c0 == "NO":                              # 헤더행 → 컬럼 위치 매핑
            cols = {_hkey(h): i for i, h in enumerate(r) if clean(h)}
            continue
        if not cols or not c0 or not c0.isdigit():
            continue

        def raw(*keys):
            for k in keys:
                i = cols.get(k)
                if i is not None and i < len(r):
                    return r[i]
            return None

        def col(*keys):
            for k in keys:
                i = cols.get(k)
                if i is not None and i < len(r):
                    v = clean(r[i])
                    if v:
                        return v
            return ""

        part = col("PART#", "PART", "MPN")
        date = norm_date(raw("DATE", "일자"))
        if not part or not date:                    # 날짜 없는 행은 기간 조회에 못 쓴다
            continue
        rows.append(dict(
            date=date, part=part, qty=to_num(raw("QTY", "QUANTITY")),
            # 입고는 공급처가 CUSTOMER 대신 SR# 칸에 적힌 행이 많다 (parse_daily 와 같은 폴백)
            customer=col("CUSTOMER", "SR#", "공급처"),
            sales=norm_sales(col("담당SALES", "SALES")),
            doc=col("SR#", "문서번호", "DOC"), mcode=col("MATERIALCODE"),
            remark=col("REMARK"), fab=col("FAB"),
        ))
    return rows


def log_conn():
    cx = sqlite3.connect(LOG_DB, timeout=30)
    cx.execute("""CREATE TABLE IF NOT EXISTS log(
        office TEXT, dir TEXT, date TEXT, customer TEXT, part TEXT, qty REAL,
        sales TEXT, doc TEXT, mcode TEXT, remark TEXT, fab TEXT)""")
    # src: 이 행이 어느 시트에서 왔나. 'sm'(shipping management) 이 'sheet'(누적 출고) 를 이긴다.
    # 같은 출고를 두 시트가 다르게 적고 있어서(누적=주문 단위, sm=lot 단위) 우선순위가 없으면
    # 마지막에 올린 파일이 이기는 비결정적 동작이 된다.
    for col in ("src", "lot", "dcode", "waybill", "memo"):
        if col not in {r[1] for r in cx.execute("PRAGMA table_info(log)")}:
            cx.execute(f"ALTER TABLE log ADD COLUMN {col} TEXT")
            if col == "src":
                cx.execute("UPDATE log SET src='sheet' WHERE src IS NULL")
            cx.commit()
    # 실 키를 이름 대신 번호(슬러그)로 쓴다. 파일명이 바뀌어도 같은 실로 인식되어
    # 이름이 달라졌다는 이유로 이력이 두 벌 쌓이는 일이 없다.
    cols = {r[1] for r in cx.execute("PRAGMA table_info(log)")}
    if "slug" not in cols:
        cx.execute("ALTER TABLE log ADD COLUMN slug TEXT")
        for (name,) in cx.execute("SELECT DISTINCT office FROM log").fetchall():
            cx.execute("UPDATE log SET slug=? WHERE office=?", (office_slug(name), name))
        cx.commit()
    cx.execute("CREATE INDEX IF NOT EXISTS ix_log ON log(office, dir, date)")
    cx.execute("CREATE INDEX IF NOT EXISTS ix_log_slug ON log(slug, dir, date)")
    # 예전 판은 재고 파일을 한 번에 여러 개 올리면 전사 출고 원장을 파일 수만큼 중복 저장했다
    # (pick_shipmgmt 주석 참고). 이미 쌓인 중복은 다시 올린 날짜만 교체되므로 저절로 없어지지
    # 않는다 — 한 번만 훑어서 완전히 같은 행이 여러 벌인 것을 한 벌로 줄인다.
    # 같은 날 같은 거래처에 같은 lot·수량을 두 번 출고한 진짜 중복까지 한 벌로 줄어들 수 있으나,
    # lot 단위 원장에서 그런 행은 사실상 없고 수량이 4배로 잡히는 쪽이 훨씬 해롭다.
    cx.execute("CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT)")
    # 선적 리포트(PO# → 발주업체). 주간 파일을 올릴 때마다 누적된다.
    # 같은 운송장·PO·파트·SO 로 수량만 다른 분할 선적이 실제로 있어서 qty 까지 키에 넣는다 —
    # 안 그러면 3,000EA 와 750EA 중 하나가 조용히 사라진다.
    SHIP_DDL = ("CREATE TABLE ship(po TEXT, npo TEXT, part TEXT, customer TEXT, "
                "sales TEXT, qty REAL, so TEXT, ww TEXT, year TEXT, tracking TEXT, "
                "UNIQUE(tracking, po, part, so, qty))")
    old = cx.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='ship'").fetchone()
    if old and "qty)" not in (old[0] or "").split("UNIQUE(")[-1]:
        cx.execute("DROP TABLE ship")                    # 옛 키로 만든 표는 버리고 다시 올린다
        old = None
    if not old:
        cx.execute(SHIP_DDL)
    cx.execute("CREATE INDEX IF NOT EXISTS ix_ship_po ON ship(po)")
    cx.execute("CREATE INDEX IF NOT EXISTS ix_ship_npo ON ship(npo)")
    return cx


def log_store(office, by_dir, src="sheet"):
    """이력을 (실, 방향, 날짜) 단위로만 갈아끼운다 — 대시보드가 거울이 아니라 금고가 되도록.

    올린 파일에 들어 있는 날짜만 지우고 새로 넣는다. 파일에 없는 과거 날짜는 그대로
    둔다. 그래서 회사에서 누적 시트를 정리하거나 새 연도 파일로 갈아타도, 한 실만
    올려도, 이미 쌓인 이력은 남는다. 같은 날짜를 다시 올리면 그 날짜만 교체되므로
    정정분은 그대로 반영된다.

    src 는 이 행이 어느 시트에서 왔는지만 기록한다 ('sheet'=누적 입고/출고, 'sm'=출고 원장).
    원장은 실 소속이 없어 slug 가 달라서, 실별 이력과 서로 덮어쓰지 않는다.
    """
    slug = LEDGER_SLUG if office == LEDGER else office_slug(office)
    cx = log_conn()
    try:
        total = 0
        for d, rs in by_dir.items():
            if not rs:
                continue
            dates = sorted({r["date"] for r in rs})
            for i in range(0, len(dates), 400):         # SQLite 변수 상한(999) 안쪽으로
                chunk = dates[i:i + 400]
                cx.execute("DELETE FROM log WHERE slug=? AND dir=? AND date IN (%s)"
                           % ",".join("?" * len(chunk)), [slug, d] + chunk)
            rows = [(office, slug, d, r["date"], r["customer"], r["part"], r["qty"],
                     r["sales"], r.get("doc", ""), r.get("mcode", ""), r.get("remark", ""),
                     r.get("fab", ""), src, r.get("lot", ""), r.get("dcode", ""),
                     r.get("waybill", ""), r.get("memo", ""))
                    for r in rs]
            cx.executemany(
                "INSERT INTO log(office,slug,dir,date,customer,part,qty,sales,doc,mcode,"
                "remark,fab,src,lot,dcode,waybill,memo) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", rows)
            total += len(rows)
        # 같은 실을 다른 이름으로 올린 적이 있으면 표시 이름만 최신으로 맞춘다
        cx.execute("UPDATE log SET office=? WHERE slug=? AND office<>?", (office, slug, office))
        cx.commit()
        return total
    finally:
        cx.close()


# 검색·정렬에서 쓸 수 있는 칸. 화면 드롭다운과 1:1 이고, 여기 없는 이름은 무시한다
# (정렬 키가 SQL 에 그대로 들어가므로 화이트리스트가 곧 주입 차단이다).
LOG_FIELDS = {"customer": "customer", "part": "part", "sales": "sales",
              "doc": "doc", "mcode": "mcode", "remark": "remark", "date": "date",
              "fab": "fab", "lot": "lot", "dcode": "dcode", "waybill": "waybill",
              "memo": "memo"}
LOG_ALL_COLS = ["customer", "part", "sales", "doc", "mcode", "remark", "lot", "dcode",
                "waybill", "memo"]
LOG_SORTS = dict(LOG_FIELDS, qty="qty")
FIELD_KO = {"customer": "거래처", "part": "PART#", "sales": "담당",
            "doc": "문서번호", "mcode": "Material Code", "remark": "비고", "date": "일자",
            "lot": "lot number", "dcode": "DATECODE", "waybill": "운송장번호",
            "memo": "메모"}


def log_scope(office):
    """실 범위 → (SQL 조각, 인자). 원장은 실 소속이 없어 항상 따로 다룬다."""
    if office == LEDGER:
        return "slug=?", [LEDGER_SLUG]              # 전사 출고 원장만
    if office and office != ALL:
        return "slug=?", [office_slug(office)]      # 그 실만
    return "slug<>?", [LEDGER_SLUG]                 # 전체 합계 = 실 전부, 원장 제외


def _log_where(office, dfrom, dto, conds):
    """조회 조건 → (SQL 조각, 인자).

    conds 는 [(칸, 값)] 이고 전부 AND 다. 칸이 'all'(또는 빈 값)이면 모든 칸을 훑는다.
    한 칸 안에서는 공백으로 나눈 단어를 다시 AND 로 묶는다.
    """
    sc, args = log_scope(office)
    where = [sc]
    if dfrom:
        where.append("date>=?"); args.append(dfrom)
    if dto:
        where.append("date<=?"); args.append(dto)
    for field, val in conds or []:
        col = LOG_FIELDS.get(field)
        for term in clean(val).split():
            like = f"%{term}%"
            if col:
                where.append(f"{col} LIKE ?"); args.append(like)
            else:                                   # 전체 검색
                where.append("(" + " OR ".join(f"{c} LIKE ?" for c in LOG_ALL_COLS) + ")")
                args += [like] * len(LOG_ALL_COLS)
    return (" AND ".join(where) or "1=1"), args


def log_query(office, dfrom, dto, conds, sort="date", direction="desc",
              limit=LOG_MAX_ROWS):
    """기간+검색으로 입고/출고를 각각 집계하고 행을 돌려준다.
    건수·수량은 상한과 무관하게 조건에 맞는 전체 기준이다.
    정렬을 서버가 하는 이유: 상한(3,000행) 때문에 화면에서 정렬하면 그 안에서만 섞인다."""
    sql, args = _log_where(office, dfrom, dto, conds)
    col = LOG_SORTS.get(sort, "date")
    asc = "ASC" if str(direction).lower() == "asc" else "DESC"
    order = f"{col} {asc}, rowid DESC" if col != "date" else f"date {asc}, rowid DESC"
    out = {}
    cx = log_conn()
    try:
        for d in ("out", "in"):
            cnt, qty, custs = cx.execute(
                f"SELECT COUNT(*),COALESCE(SUM(qty),0),COUNT(DISTINCT customer) "
                f"FROM log WHERE dir=? AND {sql}", [d] + args).fetchone()
            rows = [dict(date=a, customer=b, part=c, qty=round(e), sales=f,
                         doc=g, mcode=h, remark=i, fab=j, lot=k or "", dcode=l or "",
                         waybill=m or "", memo=o or "")
                    for a, b, c, e, f, g, h, i, j, k, l, m, o in cx.execute(
                        f"SELECT date,customer,part,qty,sales,doc,mcode,remark,fab,lot,"
                        f"dcode,waybill,memo "
                        f"FROM log WHERE dir=? AND {sql} ORDER BY {order} "
                        f"LIMIT ?", [d] + args + [limit])]
            out[d] = dict(cnt=cnt, qty=round(qty), customers=custs,
                          rows=rows, shown=len(rows), truncated=cnt > len(rows))
    finally:
        cx.close()
    out["net"] = out["in"]["qty"] - out["out"]["qty"]
    return out


def shipmgmt_sheet(wb):
    """재고 파일의 'shipping management' 시트. 이름이 아니라 헤더로 찾는다."""
    for s in wb.sheetnames:
        for r in wb[s].iter_rows(min_row=1, max_row=6, values_only=True):
            if r and {"DATE", "CUSTOMER", "PART#", "QTY"} <= {_hkey(c) for c in r if clean(c)}:
                return wb[s]
    return None


def parse_shipmgmt(ws):
    """shipping management → 출고 행 목록.

    누적 '출고' 시트가 주문 단위인 데 비해 이쪽은 lot 단위로 쪼개져 있고,
    출고 시트에 아예 빠진 거래처(한국전자판매·크래비스 등)까지 들어 있다.
    실측: 겹치는 기간에 출고 시트 4,959행 / 6,104만EA 대 이 시트 89,128행 / 9.7억EA.
    """
    cols, out, hdr = {}, [], []
    mrows = memo_rows(ws)
    for rn, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if not r:
            continue
        keys = {_hkey(c) for c in r if clean(c)}
        if {"DATE", "CUSTOMER", "PART#"} <= keys:       # 헤더행
            cols = {_hkey(h): i for i, h in enumerate(r) if clean(h)}
            hdr = list(r)
            continue
        if not cols:
            continue

        def raw(k):
            i = cols.get(k)
            return r[i] if i is not None and i < len(r) else None

        def col(k):
            return clean(raw(k))

        date, part = norm_date(raw("DATE")), col("PART#")
        if not date or not part:
            continue
        lot, dc = col("LOTNUMBER"), col("DATECODE")
        out.append(dict(
            date=date, part=part, customer=col("CUSTOMER"),
            qty=to_num(raw("QTY")), sales=norm_sales(col("SALES")),
            # '.' 은 '해당 없음' 표기로 쓰이고 있어서 빈 값으로 본다
            lot="" if lot == "." else lot, dcode="" if dc == "." else dc,
            doc="", mcode="", remark="", fab="",
            memo=memo_text(mrows, rn, hdr),
        ))
    return out


def pick_shipmgmt(snaps):
    """파일별 shipping management 스냅샷 → 그중 한 파일만 골라 낸 출고 행.

    재고 파일은 실별로 오지만 그 안의 'shipping management' 는 실 시트가 아니라
    전사 공용 원장의 사본이다. 4개 실 파일을 한 번에 올리면 같은 출고가 4벌 들어와
    원장 수량이 그대로 4배가 된다 (현장 지적: '전체 출고 원장 출고값 중복').

    현장 요청대로 '한 곳의 시트만' 쓴다. 날짜별로 파일을 섞어 고르면, 정정 때문에
    같은 출고가 파일마다 다른 날짜에 적혀 있을 경우 두 벌이 되어 중복이 되살아난다.
    파일 하나로 통일하면 그런 경로 자체가 없다.

    고르는 기준은 행이 가장 많은 파일 — 사본 중 가장 덜 잘린 것이다. 같으면 먼저
    올린 파일. 고른 파일에 없는 날짜는 건드리지 않는다 (log_store 가 올라온 날짜만
    갈아끼우므로, 전에 쌓아 둔 그 날짜 이력은 그대로 남는다).

    반환: (행 목록, 고른 파일명)
    """
    best = 0
    for i in range(1, len(snaps)):
        if len(snaps[i][1]) > len(snaps[best][1]):
            best = i
    name, rows = snaps[best]
    return rows, name


def ledger_dupes(apply=False):
    """이미 쌓인 전사 출고 원장에서 '파일 N벌이 통째로 겹친' 흔적을 찾는다.

    이 판 이전에는 재고 파일 여러 개를 한 번에 올리면 같은 원장이 파일 수만큼 저장됐다.
    지나간 날짜는 다시 올리지 않는 한 부풀어 있는 채로 남는다.

    '완전 동일 행 = 중복' 으로 지우면 안 된다. 원장은 릴 단위라 같은 날 같은 거래처에
    같은 lot·같은 수량을 여러 줄로 적는 게 정상이고, 실측으로 한 파일 안에만 그런 행이
    20,071개(전체의 26.5%, 1.9억 EA) 있다.

    그래서 날짜 단위로 본다. 그 날짜의 행 종류별 개수를 모두 세고 최대공약수를 구한다.
    파일이 N벌 겹쳤다면 모든 종류의 개수가 정확히 N배라 gcd 가 N 이 된다. 겹치지 않은
    날짜는 개수들이 서로 소라 gcd 가 1 이다 (실측: 정상 파일 308일 전부 gcd=1).
    gcd 가 N 인 날짜만 종류별 개수를 1/N 로 줄인다.

    apply=False 면 세기만 한다. 되돌릴 수 없는 삭제라 기본은 미리보기다.
    반환: {dates, rows, qty, detail:[(날짜, 배수, 지울 행 수)]}
    """
    cx = log_conn()
    try:
        rows = cx.execute(
            "SELECT rowid,date,customer,part,qty,sales,doc,mcode,remark,fab,lot,dcode,waybill"
            " FROM log WHERE slug=? AND dir='out' ORDER BY rowid", (LEDGER_SLUG,)).fetchall()
        per = defaultdict(lambda: defaultdict(list))     # 날짜 → 행종류 → [rowid]
        for r in rows:
            per[r[1]][r[2:]].append(r[0])
        kill, detail = [], []
        for d in sorted(per):
            groups = per[d]
            g = 0
            for ids in groups.values():
                g = math.gcd(g, len(ids))
            if g < 2:
                continue
            n = 0
            for ids in groups.values():
                keep = len(ids) // g
                kill.extend(ids[keep:]); n += len(ids) - keep
            detail.append((d, g, n))
        qty = 0.0
        if kill:
            for i in range(0, len(kill), 400):
                chunk = kill[i:i + 400]
                q = cx.execute("SELECT SUM(qty) FROM log WHERE rowid IN (%s)"
                               % ",".join("?" * len(chunk)), chunk).fetchone()[0]
                qty += q or 0.0
            if apply:
                for i in range(0, len(kill), 400):
                    chunk = kill[i:i + 400]
                    cx.execute("DELETE FROM log WHERE rowid IN (%s)"
                               % ",".join("?" * len(chunk)), chunk)
                cx.commit()
        return dict(dates=len(detail), rows=len(kill), qty=qty, detail=detail)
    finally:
        cx.close()


def log_day(office, date):
    """누적 이력에서 하루치를 꺼내 하루 보기와 똑같은 모양으로 만든다.

    업로드한 파일의 날짜 시트는 오늘·어제뿐이라, 그보다 과거를 하루 보기로 열려면
    DB 에서 되살려야 한다. day_block() 을 그대로 써서 KPI·차트·요약이 같은 계산을 탄다.
    """
    sc, sargs = log_scope(office)
    args = [date] + sargs
    sql = "date=? AND " + sc
    cx = log_conn()
    try:
        rows = {"in": [], "out": []}
        for d, cust, part, qty, sales, doc, remark, fab, lot, dc, wb_, mo in cx.execute(
                f"SELECT dir,customer,part,qty,sales,doc,remark,fab,lot,dcode,waybill,memo "
                f"FROM log WHERE {sql}", args):
            rows[d].append(dict(customer=cust, part=part, qty=qty, sales=sales,
                                doc=doc, remark=remark, fab=fab, vendor="",
                                lot=lot or "", dcode=dc or "", waybill=wb_ or "",
                                memo=mo or ""))
        # 비교 기준이 될 직전 영업일 (달력상 전날이 아니라 '자료가 있는 전날')
        prev = cx.execute(f"SELECT MAX(date) FROM log WHERE date<? AND {sc}",
                          [date] + sargs).fetchone()[0]
    finally:
        cx.close()
    return day_block(date, rows["in"], rows["out"]), prev


def log_dates(office, limit=400):
    """그 실에 자료가 있는 날짜 목록 (최근 순). 하루 보기 달력의 선택지."""
    sql, args = log_scope(office)
    cx = log_conn()
    try:
        return [d for (d,) in cx.execute(
            f"SELECT DISTINCT date FROM log WHERE {sql} ORDER BY date DESC LIMIT ?",
            args + [limit])]
    finally:
        cx.close()


def log_span(office=None):
    """그 실의 이력이 언제부터 언제까지 있는지 — 달력 입력의 범위로 쓴다."""
    cx = log_conn()
    try:
        if office == LEDGER:
            row = cx.execute("SELECT MIN(date),MAX(date),COUNT(*) FROM log "
                             "WHERE slug=?", (LEDGER_SLUG,)).fetchone()
        elif office and office != ALL:
            row = cx.execute("SELECT MIN(date),MAX(date),COUNT(*) FROM log "
                             "WHERE slug=?", (office_slug(office),)).fetchone()
        else:
            row = cx.execute("SELECT MIN(date),MAX(date),COUNT(*) FROM log "
                             "WHERE slug<>?", (LEDGER_SLUG,)).fetchone()
    finally:
        cx.close()
    return dict(min=row[0] or "", max=row[1] or "", rows=row[2] or 0)


# ---------------------------------------------------------------- 선적 리포트 (PO# → 발주업체)
# 재고 시트의 'CUSTOMER' 컬럼에는 고객명이 아니라 PO#(26DIT0213N10 꼴)가 들어 있다.
# 선적 리포트의 'PO Number' 가 같은 값이라 이걸로 조인하면 진짜 발주업체명이 붙는다.
# (요청받은 '발주업체+파트' 조인은 두 파일의 거래처 표기가 달라 실측 0건이었다.)


def norm_po(s):
    """분할 발주 꼬리표를 뗀다: 26YS0211N4-5-5 → 26YS0211N4. 정확 일치가 실패할 때만 쓴다."""
    return re.sub(r"-\d+(-\d+)*$", "", clean(s)).upper()


def shipping_sheet(wb):
    """선적 리포트 시트를 찾는다. 시트명이 매주 바뀌므로(ShipRpt_..._2) 헤더로 판별한다."""
    for s in wb.sheetnames:
        for r in wb[s].iter_rows(min_row=1, max_row=6, values_only=True):
            if r and {"PONUMBER", "PARTNUMBER"} <= {_hkey(c) for c in r if clean(c)}:
                return wb[s]
    return None


def parse_shipping(ws):
    """선적 리포트 → [{po, part, customer, sales, qty, so, ww, year, tracking}]."""
    cols, out = {}, []
    for r in ws.iter_rows(values_only=True):
        if not r:
            continue
        keys = {_hkey(c) for c in r if clean(c)}
        if {"PONUMBER", "PARTNUMBER"} <= keys:            # 헤더행
            cols = {_hkey(h): i for i, h in enumerate(r) if clean(h)}
            continue
        if not cols:
            continue

        def col(*names):
            for n in names:
                i = cols.get(n)
                if i is not None and i < len(r):
                    v = clean(r[i])
                    if v:
                        return v
            return ""

        po, part = col("PONUMBER"), col("PARTNUMBER")
        if not po or not part:
            continue
        out.append(dict(
            po=po, part=part, customer=col("CUSTOMER"),
            sales=col("SALES", "담당SALES"), qty=to_num(
                r[cols["TOTALQTY"]] if cols.get("TOTALQTY") is not None
                and cols["TOTALQTY"] < len(r) else 0),
            so=col("SO#", "ORDERNUMBER"), ww=col("선적WW", "WW"), year=col("연도", "YEAR"),
            tracking=col("TRACKINGNUMBER"),
        ))
    return out


def ship_store(rows):
    """선적 행을 누적한다. 같은 선적 라인(운송장+PO+파트+SO)을 다시 올려도 늘어나지 않는다."""
    cx = log_conn()
    try:
        cx.executemany(
            "INSERT OR REPLACE INTO ship VALUES(?,?,?,?,?,?,?,?,?,?)",
            [(r["po"], norm_po(r["po"]), r["part"], r["customer"], r["sales"],
              r["qty"], r["so"], r["ww"], r["year"], r["tracking"]) for r in rows])
        cx.commit()
    finally:
        cx.close()


def po_map():
    """PO# → [발주업체]. 정확 일치용과 꼬리표 제거용 두 벌.

    한 PO#가 두 업체로 선적된 경우가 실제로 있다(26DIT0213N10 → 동일기연 WW15,
    연승일레콤 WW25). 임의로 하나를 고르면 조용히 틀린 회사를 보여주게 되므로
    선적 수량 큰 순서로 전부 담고, 화면에서 '외 N' 으로 알린다.
    """
    cx = log_conn()
    try:
        exact, loose = defaultdict(list), defaultdict(list)
        for key, col in (("po", exact), ("npo", loose)):
            for k, cust, _ in cx.execute(
                    f"SELECT {key}, customer, SUM(qty) q FROM ship "
                    f"WHERE customer<>'' GROUP BY {key}, customer ORDER BY q DESC"):
                col[k.upper()].append(cust)
        n = cx.execute("SELECT COUNT(*) FROM ship").fetchone()[0]
    finally:
        cx.close()
    return dict(exact=dict(exact), loose=dict(loose), rows=n)


def po_lookup(pos, pm=None):
    """재고의 PO# 목록 → {PO#: [발주업체]}. 정확 일치 우선, 없으면 꼬리표 떼고 한 번 더."""
    pm = pm or po_map()
    out = {}
    for p in pos:
        if not p:
            continue
        c = pm["exact"].get(clean(p).upper()) or pm["loose"].get(norm_po(p))
        if c:
            out[p] = c
    return out


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
    # 메모를 붙이려면 엑셀 행번호가 필요하다 — rows[2:] 의 j 번째가 엑셀 j+3 행.
    data2 = [(j + 3, r) for j, r in enumerate(rows[2:])
             if r and i_part < len(r) and clean(r[i_part])]
    data = [r for _, r in data2]
    mrows = memo_rows(ws)

    def idx(*names):
        for n in names:
            for i, h in enumerate(hdr):
                if h.replace(" ", "").upper() == n.replace(" ", "").upper():
                    return i
        return -1

    C = dict(
        office=idx("Sales team"), central=idx("Central"),
        # 시트마다 VENDER/VENDOR/제조사로 적혀 온다 — 하나만 보면 조용히 빈 칸이 된다
        vender=idx("VENDER", "VENDOR", "제조사", "MAKER", "MFR"),
        family=idx("FAMILY"), part=idx("Part#"), mobis=idx("MOBIS ID"),
        pn=idx("품번"), qty=idx("Q'ty"), avail=idx("available Q'ty"),
        booking=idx("booking"), customer=idx("CUSTOMER"), sales=idx("SALES"),
        crd=idx("CRD"),
    )
    if C["part"] < 0 or C["qty"] < 0:
        return None

    # Datecode 연도 컬럼. 헤더는 'Datecode' 다음 줄에 연도가 오는 꼴이라 연도만 떼어 쓴다.
    # 연도 하나에 칼럼 하나여야 한다 — 같은 연도가 두 번 나오면 dict 가 조용히 덮어써서
    # 진짜 그 연도 칼럼이 사라지고 엉뚱한 값이 그 해로 잡힌다. 덮어쓰되 흔적은 남긴다.
    years = {}
    for i, h in enumerate(hdr):
        if "DATECODE" not in h.upper():
            continue
        m = re.fullmatch(r"DATECODE\s*(20\d{2})", clean(h).upper(), re.S)
        if not m:                                   # 'Datecode 2019~2022' 같은 변종
            print(f"[재고] 해석 못 한 Datecode 헤더 (무시): {h!r}")
            continue
        y = int(m.group(1))
        if y in years:
            print(f"[재고] Datecode {y} 칼럼이 두 개다 ({years[y]}, {i}) — 뒤쪽을 쓴다")
        years[y] = i

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

    for rn, r in data2:
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

        old_q, oldest, dq = 0.0, None, {}
        for y, ci in years.items():
            v = to_num(r[ci]) if ci < len(r) else 0.0
            if v > 0:
                dc[y][0] += 1; dc[y][1] += v
                dq[y] = dq.get(y, 0.0) + v          # 화면에 연도별로 펼쳐 보여준다
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
            qty=q, avail=a, booking=b, old=old_q, oldest=oldest, dcq=dq,
            # 재고 시트 메모는 booking(예약) 칸에 단다 — 그 칸 것만 보여준다
            memo=memo_text(mrows, rn, hdr, only=C["booking"]),
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
        years=sorted(years),
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
                      sales=r["sales"], fab=r["fab"], remark=r["remark"],
                      memo=r.get("memo", "")) for r in inbound],
        out_rows=[dict(customer=r["customer"], part=r["part"], qty=round(r["qty"]),
                       sales=r["sales"], doc=r.get("doc", ""), remark=r.get("remark", ""),
                       vendor=r.get("vendor", ""), lot=r.get("lot", ""),
                       dcode=r.get("dcode", ""), waybill=r.get("waybill", ""),
                       memo=r.get("memo", ""))
                  for r in outbound],
    )


# 재고 시트 헤더 (메모 앞에 붙여 저장했던 칸이름) — 옛 저장분 정리에만 쓴다
HDRTAG = re.compile(r"booking|q'?ty|availableq'?ty|datecode\d{4}|\d{1,2}일|mobisid|part#"
                    r"|customer|sales|salesteam|family|vender|vendor|site|central"
                    r"|품번|crd|전월|no|date")


def migrate_saved_memos():
    """저장돼 있는 재고 메모를 booking 칸 것만 남기게 고친다 (한 번만).

    예전엔 행에 달린 메모를 전부 '칸이름: 내용' 으로 이어 붙여 저장했다. 파싱은 고쳤지만
    이미 올라간 데이터는 그대로라, 엑셀을 다시 올리기 전까지 옛 메모가 그대로 보인다.
    """
    try:
        if not os.path.exists(DATA_FILE):
            return
        with open(DATA_FILE, encoding="utf-8") as f:
            data = json.load(f)
        if data.get("memo_ver") == MEMO_VER:
            return
        n = 0
        for inv in (data.get("inventories") or {}).values():
            for it in inv.get("items") or []:
                old = it.get("memo") or ""
                if not old:
                    continue
                new = _booking_part(old)
                if new != old:
                    it["memo"] = new; n += 1
        data["memo_ver"] = MEMO_VER
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        print(f"[메모] 저장된 재고 메모 {n}건을 booking 칸 것만 남기게 정리")
    except Exception as e:
        print(f"[메모] 저장분 정리 실패 — 건너뜀: {e}")


def _booking_part(txt):
    """'Q'ty: ... / booking: ... / 1일: ...' 에서 booking 칸 몫만 떼어 낸다.

    칸이름 없는 조각은 바로 앞 칸에 이어지는 내용이다 (메모 안에 ' / ' 가 들어 있던 경우).
    칸이름이 하나도 없으면 이미 새 형식이므로 그대로 둔다.
    """
    cur, keep, tagged = None, [], False
    for pc in txt.split(" / "):
        m = re.match(r"([^:]{1,20}):\s*(.*)$", pc, re.S)
        # 메모 내용에도 콜론이 들어간다 ('출고: 1,000ea'). 재고 시트에 실제로 있는
        # 칸이름일 때만 칸 표시로 본다 — 아니면 앞 칸 내용이 이어지는 것으로 친다.
        if m and HDRTAG.fullmatch(re.sub(r"\s+", "", m.group(1)).lower()):
            tagged = True
            cur = re.sub(r"\s+", "", m.group(1)).lower()
            pc = m.group(2)
        if cur == "booking" and pc:
            keep.append(pc)
    return txt if not tagged else " / ".join(keep)


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


def fill_days_from_log(per, by_dir):
    """날짜 시트의 입고/출고 섹션이 비어 있으면 같은 워크북의 누적 시트로 채운다.

    현장 운영이 이렇다: 날짜 시트는 아침에 틀만 만들어 두고(NO·DATE 만 채운 빈 행),
    입고 행은 그날 늦게 적는다. 그 사이에 파일을 올리면 하루 보기 입고가 0 으로 뜬다
    (영업5실 2026-09-04 실측: 아침 업로드 시 입고 0행 / 같은 시트 오후 10행).

    누적 '입고'/'출고' 시트에 그 날짜 행이 이미 있으면 그것으로 채운다.
    누적에도 없으면 그대로 0 이다 — 없는 값을 지어내지 않는다.

    반환: (보완된 per, {날짜: ['in','out']})
    """
    idx = {}
    for d, rows in (by_dir or {}).items():
        by_date = defaultdict(list)
        for r in rows:
            if r.get("date"):
                by_date[r["date"]].append(r)
        idx[d] = by_date

    out, filled = {}, defaultdict(list)
    for date, (inb, outb) in per.items():
        if not inb:
            got = idx.get("in", {}).get(date)
            if got:
                inb = got
                filled[date].append("in")
        if not outb:
            got = idx.get("out", {}).get(date)
            if got:
                outb = got
                filled[date].append("out")
        out[date] = (inb, outb)
    return out, dict(filled)


KST = datetime.timezone(datetime.timedelta(hours=9))


def today_kst():
    """오늘 날짜(한국). 서버가 UTC(Render)로 돌아도 현장 날짜를 쓴다."""
    return datetime.datetime.now(KST).strftime("%Y-%m-%d")


def today_index(dates):
    """'오늘' 로 볼 날짜 시트의 위치.

    마지막 시트를 그냥 오늘로 잡으면 안 된다 — 현장에서 다음 날 시트를 빈 틀로 미리
    만들어 두기 때문에 아직 오지 않은 날짜가 오늘이 돼 버린다 (실측: 2026-09-03
    18:17 에 나간 영업5실 메일 제목이 2026-09-04). 한국 날짜로 오늘을 넘지 않는
    마지막 시트를 고른다. 전부 미래면(날짜를 잘못 적은 파일) 마지막 시트를 그대로
    쓴다 — 화면이 비는 것보다 낫다.
    """
    if not dates:
        return 0
    today = today_kst()
    past = [i for i, d in enumerate(dates) if d <= today]
    return past[-1] if past else len(dates) - 1


def office_block(name, per, filled=None):
    """{날짜:(inb,outb)} → 실 1개 대시보드 데이터."""
    dates = sorted(per)
    days = [day_block(d, per[d][0], per[d][1]) for d in dates]
    for d in days:                                  # 누적 시트로 채운 날은 화면에 밝힌다
        d["filled"] = (filled or {}).get(d["date"], [])
    return dict(name=name, dates=dates, today_idx=today_index(dates), days=days)


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
        datecode=dc, years=years,
        month=dict(prev=sum(inv["month"]["prev"] for inv in invs),
                   inbound=sum(inv["month"]["inbound"] for inv in invs),
                   outbound=sum(inv["month"]["outbound"] for inv in invs),
                   daily_in=[dsum("daily_in", i) for i in range(ndays)],
                   daily_out=[dsum("daily_out", i) for i in range(ndays)]),
        customers=[dict(name=n, booking=v)
                   for n, v in sorted(cust.items(), key=lambda x: -x[1])[:12]],
        items=items,
    )


def decode_upload(b64):
    """브라우저 FileReader.readAsDataURL 결과에서 엑셀 바이트를 뽑는다.

    접두사 길이는 브라우저가 붙이는 MIME에 따라 달라진다.
      data:application/octet-stream;base64,...                        → 37자
      data:application/vnd.openxmlformats-...spreadsheetml.sheet;...  → 78자
    base64 알파벳에 콤마는 없으므로, 첫 콤마 앞은 전부 접두사다.
    """
    if b64.startswith("data:"):
        b64 = b64.split(",", 1)[1]
    return base64.b64decode(b64)


def build_payload(files, password):
    """files: [{name, file(base64)}] → {offices:[...], inventories:{실: {...}}}

    파일 종류는 시트 이름으로 자동 판별한다.
      날짜 시트(YYYY-MM-DD) 있음  → 일일 입출고 파일
      '~ inventory' 시트 있음      → 재고 현황 파일
    재고도 실별로 보관한다. (전에는 하나만 남아 마지막 파일이 앞의 것을 덮어썼다)
    """
    raw_offices, inventories = [], {}     # [(실이름, {날짜:(inb,outb)}, 보완표시)]
    logs = {}                             # 실이름 → {'in': [...], 'out': [...]} 누적 이력
    ships = []                            # 선적 리포트 행 (PO# → 발주업체)
    shipmgmt = []                         # 전사 출고 원장 스냅샷 [(파일명, 행들)]
    for f in files:
        name = f.get("name") or "실"
        wb = open_wb(decode_upload(f["file"]), password)
        try:
            per = parse_workbook(wb)
            if per:                                     # 입출고 파일
                # 같은 파일의 누적 시트 = 기간 조회용 이력 (날짜 시트는 2일치뿐이라)
                logs[name] = {d: parse_log(wb[s]) for d, s in log_sheets(wb).items()}
                # 날짜 시트 입고/출고 섹션이 비어 있으면 누적 시트로 채운다
                per, filled = fill_days_from_log(per, logs[name])
                if filled:
                    for dt, dirs in sorted(filled.items()):
                        print(f"[보완] {name} {dt} — 날짜 시트가 비어 누적 시트로 채움: "
                              f"{'·'.join(dirs)}")
                raw_offices.append((name, per, filled))
                continue
            ws = inventory_sheet(wb)                    # 재고 파일
            if ws is not None:
                inv = parse_inventory(ws)
                if inv:
                    inventories[name] = inv
                    # 재고 파일마다 들어 있는 'shipping management' 는 실별 시트가 아니라
                    # 전사 공용 출고 원장의 스냅샷이다 (실측: 파일 간 행 단위 99.4~99.8% 동일,
                    # lot 95.5%가 3개 실에 중복). 여기서 합치면 같은 출고가 파일 수만큼
                    # 곱해지므로 파일별로 따로 들고 있다가 pick_shipmgmt() 로 날짜마다
                    # 한 파일만 고른다.
                    sm = shipmgmt_sheet(wb)
                    if sm is not None:
                        shipmgmt.append((name, parse_shipmgmt(sm)))
                    continue
            ws = shipping_sheet(wb)                     # 선적 리포트 (PO# → 발주업체)
            if ws is not None:
                ships.extend(parse_shipping(ws))
        finally:
            wb.close()

    if ships:
        ship_store(ships)

    if not raw_offices and not inventories and not ships:
        raise ValueError("날짜 시트(YYYY-MM-DD)가 있는 입출고 파일, 'inventory' 시트가 있는 "
                         "재고 파일, 'PO Number' 가 있는 선적 리포트 중 아무것도 찾지 못했습니다.")

    # ---- 거래처 표기 통일 (Mobis / MOBIS → 가장 많이 쓰인 표기 하나로) ----
    # 파일 전체를 본 뒤에야 대표 표기를 고를 수 있어 여기서 일괄 처리한다.
    C = Companies()
    for _, per, _f in raw_offices:
        for inb, outb in per.values():
            for r in inb + outb:
                C.add(r.get("customer"))
    for by_dir in logs.values():
        for rows in by_dir.values():
            for r in rows:
                C.add(r.get("customer"))
    for _, rows in shipmgmt:
        for r in rows:
            C.add(r.get("customer"))
    for inv in inventories.values():
        for it in inv["items"]:
            C.add(it.get("customer"))

    for _, per, _f in raw_offices:
        for inb, outb in per.values():
            for r in inb + outb:
                r["customer"] = C.canon(r.get("customer"))
    for name, by_dir in logs.items():
        for rows in by_dir.values():
            for r in rows:
                r["customer"] = C.canon(r.get("customer"))
        log_store(name, by_dir)
    if shipmgmt:
        sm_rows, sm_file = pick_shipmgmt(shipmgmt)
        for r in sm_rows:
            r["customer"] = C.canon(r.get("customer"))
        log_store(LEDGER, {"out": sm_rows}, src="sm")
        if len(shipmgmt) > 1:
            dropped = sum(len(rows) for _, rows in shipmgmt) - len(sm_rows)
            print(f"[원장] 파일 {len(shipmgmt)}개 중 '{sm_file}' 채택 · "
                  f"{len(sm_rows):,}행 / 중복 {dropped:,}행 제외")
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
    for name, per, filled in raw_offices:
        offices.append(office_block(name, per, filled))
        for d, (inb, outb) in per.items():
            allbydate[d][0].extend(inb); allbydate[d][1].extend(outb)

    if len(offices) > 1:                                # 전체 합계 탭 (맨 앞)
        agg = {d: (allbydate[d][0], allbydate[d][1]) for d in allbydate}
        offices.insert(0, office_block(ALL, agg))
    if len(inventories) > 1:
        inventories[ALL] = merge_inventories(list(inventories.values()))

    # 실별 이력 보유 구간 — 기간 입력의 min/max 로 쓴다 (행 자체는 서버가 들고 있다)
    spans = {o["name"]: log_span(o["name"]) for o in offices}
    spans[LEDGER] = log_span(LEDGER)                # 전사 출고 원장 (실 소속 없음)
    if shipmgmt:
        spans[LEDGER]["src"] = sm_file             # 이번에 채택한 파일
    return dict(offices=offices, inventories=inventories, spans=spans,
                memo_ver=MEMO_VER)


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
        for rc in ([red_col] if isinstance(red_col, int) else (red_col or [])):
            for row in ws.iter_rows(min_row=4, max_row=ws.max_row,
                                    min_col=rc, max_col=rc):
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
        # 재고의 'CUSTOMER' 칸은 실제로 PO# 다. 헤더를 바로잡고, 선적 리포트로
        # 찾아낸 발주업체를 옆에 붙인다 (리포트를 안 올렸으면 빈 칸으로 나간다).
        buyers = po_lookup([i.get("customer") for i in inv["items"]])

        def dq(it, y):
            return to_num((it.get("dcq") or {}).get(str(y)))

        yrs = [y for y in (inv.get("years") or [])
               if any(dq(i, y) > 0 for i in inv["items"])]
        old_y = inv.get("old_year") or OLD_YEAR
        base_h = ["PART#", "PO# / 고객", "발주업체", "MOBIS ID", "FAMILY", "VENDER", "실", "담당",
                  "재고", "가용", "예약"]
        # 화면과 같은 폴백: 연도별 수량이 없는 옛 자료면 예전 장기재고/Datecode 칼럼으로
        def gap(i):
            d = (i.get("dcq") or {})
            if not d:
                return ""
            t = sum(to_num(v) for v in d.values())
            return "" if abs(t - to_num(i.get("qty"))) < 0.5 else f"Datecode 합 {t:,.0f}"

        tail_h = ([f"DC {y}" for y in yrs] + ["확인"]) if yrs else ["장기재고", "Datecode"]
        tail_w = ([11] * len(yrs) + [20]) if yrs else [12, 11]
        tail = ((lambda i: [dq(i, y) or "" for y in yrs] + [gap(i)]) if yrs
                else (lambda i: [i.get("old") or "", i.get("oldest") or ""]))
        red = ([len(base_h) + n for n, y in enumerate(yrs, 1) if y <= old_y] if yrs
               else [len(base_h) + 1])
        sheet("품목", f"품목별 재고 · {office}",
              base_h + tail_h,
              [[i["part"], i["customer"], " / ".join(buyers.get(i["customer"], [])),
                i["mobis"], i["family"], i["vender"], i["office"], i["sales"],
                i["qty"], i["avail"], i["booking"]] + tail(i) for i in
               sorted(inv["items"], key=lambda x: -x["qty"])],
              widths=[26, 18, 24, 16, 16, 12, 10, 10, 12, 12, 12] + tail_w,
              # 장기재고(=old_year 이전) 연도 칼럼만 빨강 — 화면 표와 같은 규칙
              red_col=red)
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
    elif view == "log":
        # 화면은 상한(LOG_MAX_ROWS)까지만 보여주지만 엑셀은 조건에 맞는 전체를 낸다.
        # 그래서 브라우저가 보낸 행을 쓰지 않고 같은 조건으로 서버가 다시 조회한다.
        dfrom, dto = payload.get("from", ""), payload.get("to", "")
        conds = [(c.get("f", ""), c.get("v", "")) for c in payload.get("conds") or []
                 if c.get("v")]
        res = log_query(office, dfrom, dto, conds, payload.get("sort") or "date",
                        payload.get("dir") or "desc", limit=1000000)
        span = f"{dfrom or '처음'} ~ {dto or '끝'}"
        if conds:
            span += " · " + " AND ".join(
                f"{FIELD_KO.get(f, '전체')}='{v}'" for f, v in conds)
        sheet("요약", f"입출고 이력 · {office} · {span}",
              ["항목", "값"],
              [["기간", span], ["입고 건수", res["in"]["cnt"]], ["입고 수량", res["in"]["qty"]],
               ["출고 건수", res["out"]["cnt"]], ["출고 수량", res["out"]["qty"]],
               ["순물동(입-출)", res["net"]], ["출고 거래처", res["out"]["customers"]]],
              widths=[20, 30])
        sheet("출고", f"출고 이력 · {office} · {span}",
              ["#", "일자", "거래처", "PART#", "수량", "담당", "lot number", "DATECODE",
               "운송장번호", "문서번호", "비고", "메모"],
              [[i + 1, r["date"], r["customer"], r["part"], r["qty"],
                r["sales"], r["lot"], r["dcode"], r["waybill"], r["doc"], r["remark"],
                r.get("memo", "")]
               for i, r in enumerate(res["out"]["rows"])],
              widths=[6, 13, 22, 26, 12, 10, 22, 12, 18, 16, 22, 40])
        sheet("입고", f"입고 이력 · {office} · {span}",
              ["#", "일자", "거래처/공급", "PART#", "수량", "담당", "SR#", "FAB", "비고",
               "메모"],
              [[i + 1, r["date"], r["customer"], r["part"], r["qty"],
                r["sales"], r["doc"], r["fab"], r["remark"], r.get("memo", "")]
               for i, r in enumerate(res["in"]["rows"])],
              widths=[6, 13, 22, 26, 12, 10, 16, 10, 22, 40])
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
              ["#", "거래처", "PART#", "수량", "담당", "lot number", "DATECODE",
               "운송장번호", "문서번호", "비고", "메모"],
              [[i + 1, r.get("customer"), r.get("part"), r.get("qty"),
                r.get("sales"), r.get("lot", ""), r.get("dcode", ""),
                r.get("waybill", ""), r.get("doc"), r.get("remark"),
                r.get("memo", "")]
               for i, r in enumerate(day.get("out_rows", []))],
              widths=[6, 22, 26, 12, 10, 22, 12, 18, 16, 22, 40])
        sheet("입고", f"입고 내역 · {office} · {day.get('date','')}",
              ["#", "거래처/공급", "PART#", "수량", "담당", "FAB", "비고", "메모"],
              [[i + 1, r.get("customer"), r.get("part"), r.get("qty"),
                r.get("sales"), r.get("fab"), r.get("remark"), r.get("memo", "")]
               for i, r in enumerate(day.get("in_rows", []))],
              widths=[6, 22, 26, 12, 10, 10, 22, 40])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- 이메일 (Outlook)
def _e(s):
    return (str(s) if s is not None else "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# 메일 하단 문의처. 답장(Reply-To)은 첫 번째 사람에게 간다.
CONTACTS = [
    ("자재관리팀", "안성우 책임", "sw.ahn@unitrontech.com"),
    ("경영기획팀", "이희서 매니저", "seanlee@unitrontech.com"),
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

    # 버튼 아래에 URL 을 그대로 한 번 더 노출하던 블록은 뺐다. 같은 링크가 본문에
    # 두 번 나오는 형태는 스팸 필터가 싫어한다. 텍스트 버전에는 URL 이 그대로 남아 있다.

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


def _send_smtp(emails, subject, body, cfg, text=None, cc=None):
    """SMTP 발송. cfg = {host, port, user, pw, sender}. text = 대체 텍스트 본문. cc = 참조."""
    import smtplib, ssl
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.utils import formataddr, formatdate, make_msgid
    host, port = cfg["host"], cfg["port"]
    user, pw, sender = cfg["user"], cfg["pw"], cfg["sender"]

    # 같은 사람이 받는사람과 참조에 동시에 들어가면 메일이 두 번 간다.
    to_keys = {e.strip().lower() for e in emails}
    cc = [c.strip() for c in (cc or []) if c.strip() and c.strip().lower() not in to_keys]

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
    if cc:
        msg["Cc"] = ", ".join(cc)
    # 답장은 발송용 계정이 아니라 문의처로 가야 한다
    msg["Reply-To"] = formataddr((CONTACT_NAME, CONTACT_MAIL))
    # Date/Message-ID 가 없는 메일은 스팸 점수가 올라간다. 지메일 릴레이는 없으면 채워주지만
    # 회사 M365(smtp.office365.com) 로 바꾸면 안 채워준다 — 직접 넣어둔다.
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=sender.split("@")[-1] if "@" in sender else None)
    # 자동 발송 메일임을 명시 → 부재중 자동응답이 되돌아오지 않는다 (RFC 3834 / M365)
    msg["Auto-Submitted"] = "auto-generated"
    msg["X-Auto-Response-Suppress"] = "OOF, AutoReply"
    with smtplib.SMTP(host, port, timeout=25) as s:
        s.ehlo()
        try:
            s.starttls(context=ssl.create_default_context()); s.ehlo()
        except Exception:
            pass
        if user:
            s.login(user, pw)
        # Cc 헤더는 표시용일 뿐이다. 실제 전달은 봉투(envelope)에 넣어야 간다.
        s.sendmail(sender, list(emails) + cc, msg.as_string())
    return {"ok": True, "sent": True, "via": "smtp", "n": len(emails) + len(cc), "cc": len(cc)}


def _send_outlook(emails, subject, body, send, cc=None):
    """로컬 PC Outlook 발송(폴백)."""
    to_keys = {e.strip().lower() for e in emails}
    cc = [c.strip() for c in (cc or []) if c.strip() and c.strip().lower() not in to_keys]
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
        if cc:
            mail.CC = "; ".join(cc)
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
    cc = [e for e in (payload.get("cc") or []) if e]
    office = payload.get("office", "")
    date = payload.get("date", "")
    day = payload.get("day") or {}
    link = payload.get("link", "")
    # 제목이 '['로 시작하면 스팸 필터가 광고성 머리말로 본다. 회사명 + 실 + 날짜로 푼다.
    subject = " ".join(x for x in ("유니트론텍 입출고 및 재고현황", office, date) if x)
    body = compose_email_html(office, date, day, link)
    text = compose_email_text(office, date, link)
    cfg = resolve_smtp(payload.get("smtp"))
    if cfg:                                             # SMTP (요청 설정 또는 환경변수)
        if not emails:
            raise ValueError("받는 사람 이메일을 입력하세요.")
        return _send_smtp(emails, subject, body, cfg, text=text, cc=cc)
    return _send_outlook(emails, subject, body, payload.get("send"), cc=cc)  # 폴백: 로컬 Outlook


# ---------------------------------------------------------------- 자동 발송 (서버)
# 발송은 업로드를 처리하는 서버 안에서 끝난다. 예전엔 업로더의 브라우저가 /send 를
# 실별로 불렀는데, 업로드하고 탭을 닫으면 메일이 한 통도 안 나가고 기록도 안 남았다.

# 고정 수신자. 키는 실 이름에서 뽑은 숫자 (영업5실 → '5', 영업1,2실 → '12').
# 여기 없는 실은 DEFAULT_EMAILS 로 간다.
# 스팸함 통과가 아직 확인 안 된 사람들이 남아 있다 — 확인되면 4·5실 수신자에 아래를 더한다:
#   "frankie@unitrontech.com", "mh.choi@unitrontech.com", "yj.park@unitrontech.com"
DEFAULT_EMAILS = [
    "seanlee@unitrontech.com",
]

# 전 실 참조(Cc). 2026-08-06 — 받는 사람이 아니라 참조로 간다.
# gy.choi 는 이전까지 받는 사람이었고, sw.ahn 은 4·5실만 받았다. 둘 다 여기로 옮겼다.
CC_EMAILS = [
    "sw.ahn@unitrontech.com",
    "gy.choi@unitrontech.com",
    "bjsoh@unitrontech.com",
]

# 실 담당자. 2026-08-06 추가 — 각 실이 자기 실 자료를 직접 받는다.
# 위 DEFAULT_EMAILS(전 실 공통 수신)는 그대로 유지되고 여기에 더해진다.
SALES_EMAILS = {
    "12": [
        "sales1@unitrontech.com",
        "sh.hong@unitrontech.com",
        "davidpark@unitrontech.com",
        "ys.jung@unitrontech.com",
        "trevis@unitrontech.com",
        "royola@unitrontech.com",
        # 2026-08-25 추가.
        "henry.jeong@unitrontech.com",
    ],
    "3": ["sales3@unitrontech.com"],
    # 2026-09-09 — 그룹 주소 sales1team@ 을 빼고 개인 주소로 전부 교체했다.
    # 그룹으로 보내면 4실이 아닌 사람에게도 퍼지는데 그걸 앱에서 막을 방법이 없다.
    "4": [
        "hw.kim@unitrontech.com",
        "yk.kwon@unitrontech.com",
        "hjpark@unitrontech.com",
        "hjgo@unitrontech.com",
        "harold@unitrontech.com",
        "jysong@unitrontech.com",
        "sccho@unitrontech.com",
        "sw.lee@unitrontech.com",
    ],
    # 2026-09-09 — 4실과 같은 이유로 그룹 주소 sales3team@ 을 빼고 개인 주소로 교체했다.
    "5": [
        "hw.kim@unitrontech.com",
        "cj.lim@unitrontech.com",
        "clark@unitrontech.com",
        "jacob@unitrontech.com",
        "hskang@unitrontech.com",
        "martin@unitrontech.com",
        "sangil@unitrontech.com",
        "hs.yang@unitrontech.com",
        # jyson(손진영). 4실 jysong 과 철자가 한 글자 다르지만 오타 아님.
        "jyson@unitrontech.com",
        "linday@unitrontech.com",
        "boeun.kim@unitrontech.com",
        "sw.lee@unitrontech.com",
        "GBC118@unitrontech.com",
    ],
}

def _office_recipients(slug):
    """받는 사람 = 공통 수신자 + 실 담당자. 중복 주소는 한 번만 남기고 순서는 유지한다.

    자재관리팀(sw.ahn)은 예전에 4·5실 받는 사람이었지만 지금은 전 실 참조(CC_EMAILS)다.
    """
    plan = list(DEFAULT_EMAILS) + SALES_EMAILS.get(slug, [])

    seen, out = set(), []
    for mail in plan:
        key = mail.strip().lower()
        if key and key not in seen:
            seen.add(key)
            out.append(mail.strip())
    return out


OFFICE_EMAILS = {slug: _office_recipients(slug) for slug in ("12", "3", "4", "5")}


def office_slug(name):
    """실 이름 → URL 슬러그. 숫자만 뽑는다. 숫자가 없으면 'all' (= 전체 합계)."""
    d = "".join(re.findall(r"\d", str(name)))
    return d or "all"


def smtp_ready():
    """환경변수만으로 실제 발송이 가능한 상태인가.

    host 만 보면 안 된다 — render.yaml 이 SMTP_HOST 를 하드코딩하고 있어서 계정이
    비어 있어도 host 는 항상 잡힌다. 계정 없이 보내면 Gmail 이 거부한다.
    """
    cfg = resolve_smtp(None)
    return bool(cfg and cfg["user"] and cfg["pw"])


def send_office_emails(result, origin):
    """업로드된 실마다 각자의 링크가 담긴 메일을 보낸다. '전체 합계'는 실이 아니므로 제외."""
    out = {"ok": [], "fail": [], "error": ""}
    offices = [o for o in (result.get("offices") or []) if o.get("name") != "전체 합계"]
    if not offices:
        return out
    if not resolve_smtp(None):
        out["error"] = "서버에 메일 계정(SMTP)이 설정돼 있지 않습니다."
        return out

    for o in offices:
        name = o.get("name", "")
        try:
            days = o.get("days") or []
            day = days[o.get("today_idx", 0)] if days else {}
            send_email({
                "office": name,
                "date": day.get("date", ""),
                "day": day,
                "link": f"{origin.rstrip('/')}/{office_slug(name)}",
                "emails": OFFICE_EMAILS.get(office_slug(name), DEFAULT_EMAILS),
                "cc": CC_EMAILS,
            })
            out["ok"].append(name)
        except Exception as e:                              # noqa: BLE001 — 한 실이 죽어도 나머지는 보낸다
            out["fail"].append(f"{name} ({e})")
            print(f"[mail] 발송 실패 {name}: {type(e).__name__}: {e}", flush=True)
    print(f"[mail] 성공 {out['ok']} / 실패 {out['fail']}", flush=True)
    return out


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
        if path == "/mailcfg":                           # 서버 설정 상태 (메일계정 / 엑셀 비밀번호)
            # 비밀번호 자체는 절대 안 내려보낸다. 있는지 없는지만.
            self._send(200, json.dumps({"configured": smtp_ready(),
                                        "pw": bool(DEFAULT_PW)}))
        elif path == "/data":                            # 저장된 데이터(있으면)
            if os.path.exists(DATA_FILE):
                with open(DATA_FILE, encoding="utf-8") as f:
                    self._send(200, f.read())
            else:
                self._send(200, json.dumps({"offices": []}))
        elif path == "/po":                              # PO# → 발주업체 (선적 리포트 누적분)
            pm = po_map()
            self._send(200, json.dumps(
                {"map": pm["exact"], "loose": pm["loose"], "rows": pm["rows"]},
                ensure_ascii=False))
        elif path == "/health":                          # 저장소 진단 (디스크가 실제로 붙었나)
            # 데이터가 재시작마다 사라질 때 원인을 눈으로 확인하려고 둔다.
            # 비밀번호·계정 같은 건 절대 싣지 않는다. 경로와 참/거짓만.
            d = os.path.dirname(LOG_DB)
            info = {"DATA_DIR": os.environ.get("DATA_DIR") or "(미설정 → 앱 폴더 사용)",
                    "db_path": LOG_DB, "dir_exists": os.path.isdir(d)}
            try:
                probe = os.path.join(d, ".write_test")
                with open(probe, "w") as f:
                    f.write("x")
                os.remove(probe)
                info["writable"] = True
            except Exception as e:  # noqa: BLE001
                info["writable"] = False
                info["write_error"] = str(e)
            try:
                # 마운트된 디스크는 앱 폴더와 device 번호가 다르다 — 이게 결정적 단서다
                info["separate_mount"] = os.stat(d).st_dev != os.stat(HERE).st_dev
            except Exception:
                info["separate_mount"] = None
            info["db_bytes"] = os.path.getsize(LOG_DB) if os.path.exists(LOG_DB) else 0
            info["saved_data_exists"] = os.path.exists(DATA_FILE)
            try:
                cx = log_conn()
                info["log_rows"] = cx.execute("SELECT COUNT(*) FROM log").fetchone()[0]
                info["ship_rows"] = cx.execute("SELECT COUNT(*) FROM ship").fetchone()[0]
                cx.close()
            except Exception as e:  # noqa: BLE001
                info["db_error"] = str(e)
            try:
                # 옛 판이 남긴 원장 중복 (지우지 않는다. 세어서 보여만 준다 —
                # 정리는 `python app.py --dedupe-ledger --apply` 로 사람이 돌린다)
                dup = ledger_dupes()
                info["ledger_dupe_dates"] = dup["dates"]
                info["ledger_dupe_rows"] = dup["rows"]
                info["ledger_dupe_qty"] = round(dup["qty"])
                if dup["rows"]:
                    info["ledger_dupe_hint"] = "python app.py --dedupe-ledger 로 확인 후 --apply"
            except Exception as e:  # noqa: BLE001
                info["ledger_dupe_error"] = str(e)
            self._send(200, json.dumps(info, ensure_ascii=False, indent=1))
        elif path == "/day":                             # 누적 이력에서 하루치 되살리기
            q = urllib.parse.parse_qs(self.path.partition("?")[2])
            g = lambda k: (q.get(k) or [""])[0].strip()   # noqa: E731
            try:
                day, prev_date = log_day(g("office"), g("date"))
                prev = log_day(g("office"), prev_date)[0] if prev_date else None
                self._send(200, json.dumps({"day": day, "prev": prev},
                                           ensure_ascii=False))
            except Exception as e:  # noqa: BLE001
                self._send(500, json.dumps({"error": f"{type(e).__name__}: {e}"},
                                           ensure_ascii=False))
        elif path == "/days":                            # 자료가 있는 날짜 목록
            q = urllib.parse.parse_qs(self.path.partition("?")[2])
            self._send(200, json.dumps(
                log_dates((q.get("office") or [""])[0].strip()), ensure_ascii=False))
        elif path == "/log":                             # 기간 조회 (입출고 이력)
            q = urllib.parse.parse_qs(self.path.partition("?")[2])
            g = lambda k: (q.get(k) or [""])[0].strip()   # noqa: E731
            conds = [(g("f1"), g("v1")), (g("f2"), g("v2"))]
            conds = [(f, v) for f, v in conds if v]
            try:
                self._send(200, json.dumps(
                    log_query(g("office"), g("from"), g("to"), conds,
                              g("sort") or "date", g("dir") or "desc"),
                    ensure_ascii=False))
            except Exception as e:  # noqa: BLE001
                self._send(500, json.dumps({"error": f"{type(e).__name__}: {e}"},
                                           ensure_ascii=False))
        else:                                            # / , /12 , /3 , /4 , /5 ... 모두 같은 페이지(클라이언트 라우팅)
            self._send(200, PAGE.replace("<!--CHARTJS-->", chart_js()), "text/html; charset=utf-8")

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            payload = json.loads(self.rfile.read(length))
            if self.path == "/build":
                files = payload.get("files") or [{"name": "", "file": payload["file"]}]
                result = build_payload(files, payload.get("password") or DEFAULT_PW)
                # 선적 리포트만 올린 경우 offices/inventories 가 비어 있다. 그대로 저장하면
                # 대시보드가 통째로 지워지므로, 그때는 저장돼 있던 것을 그대로 살린다.
                if not result["offices"] and not result["inventories"]:
                    try:
                        with open(DATA_FILE, encoding="utf-8") as f:
                            prev = json.load(f)
                        # 이력이 늘었을 수 있으니 보유 구간은 다시 계산해 준다
                        prev["spans"] = {o["name"]: log_span(o["name"])
                                         for o in prev.get("offices", [])}
                        result = {**prev, "ship_only": True}
                        with open(DATA_FILE, "w", encoding="utf-8") as f:
                            json.dump({k: v for k, v in result.items()
                                       if k != "ship_only"}, f, ensure_ascii=False)
                    except Exception:
                        pass
                else:
                    try:                                 # 업로드 결과 저장 → 링크 연 사람 모두 공유
                        with open(DATA_FILE, "w", encoding="utf-8") as f:
                            json.dump(result, f, ensure_ascii=False)
                    except Exception:
                        pass
                # 발송은 여기서 끝낸다. 업로더가 탭을 닫아도 메일은 나간다.
                origin = payload.get("origin") or f"http://{self.headers.get('Host', 'localhost')}"
                result["mail"] = send_office_emails(result, origin)
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
.pwtoggle{margin-top:16px;color:#8b8b96;font-size:12px;cursor:pointer;text-decoration:underline;text-underline-offset:3px;}
.pwtoggle:hover{color:#ccc;}
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
.daypickwrap{display:inline-flex;align-items:center;gap:5px;padding:0 8px 0 10px;font-size:12px;color:var(--mut);}
.daypickwrap input[type=date]{font-family:inherit;font-size:12.5px;font-weight:700;color:var(--ink);
  border:1.5px solid #dcdce3;background:#fff;border-radius:8px;padding:6px 8px;outline:none;cursor:pointer;}
.daypickwrap input[type=date]:focus{border-color:var(--red);}
/* 기간 조회 바 */
.logbar{display:flex;flex-wrap:wrap;align-items:center;gap:8px;background:#fff;border:1.5px solid var(--line);
  border-radius:14px;padding:12px 14px;margin-bottom:22px;box-shadow:0 1px 2px rgba(0,0,0,.04);}
.logbar .lab{font-size:11.5px;color:var(--mut);font-weight:700;}
.logbar input[type=date]{font-family:inherit;font-size:13px;font-weight:600;color:var(--ink);
  border:1.5px solid var(--line);border-radius:9px;padding:7px 10px;outline:none;}
.logbar input[type=date]:focus{border-color:var(--red);}
.logbar .dpair{display:inline-flex;align-items:center;gap:4px;}  /* 좁은 화면에서 날짜 한 쌍이 갈라지지 않게 */
.logbar .tilde{color:var(--mut);font-weight:700;margin:0 2px;}
.logbar .pre{border:1.5px solid var(--line);background:#fff;font-family:inherit;font-size:12px;font-weight:700;
  color:#777;padding:7px 12px;border-radius:9px;cursor:pointer;transition:.15s;}
.logbar .pre:hover{border-color:#cfcfd6;color:var(--ink);}
.logbar .pre.on{background:var(--ink);color:#fff;border-color:var(--ink);}
.logbar .q{flex:1;min-width:240px;font-family:inherit;font-size:13.5px;padding:8px 13px;
  border:1.5px solid var(--line);border-radius:10px;outline:none;}
.logbar .q:focus{border-color:var(--red);}
.logbar .go{border:none;background:var(--red);color:#fff;font-family:inherit;font-size:13px;font-weight:800;
  padding:9px 18px;border-radius:10px;cursor:pointer;}
.logbar .go:hover{background:var(--red2);}
.logbar .off{border:1.5px solid var(--line);background:#fff;font-family:inherit;font-size:12.5px;font-weight:700;
  color:#777;padding:8px 14px;border-radius:10px;cursor:pointer;}
.logbar .off:hover{border-color:#cfcfd6;color:var(--ink);}
.logbar .sep{width:1px;height:22px;background:var(--line);margin:0 2px;}
.logbar .note{font-size:11.5px;color:var(--mut);width:100%;padding-top:2px;}
.trunc{padding:10px 16px;background:#fff8e6;color:#8a6d1f;font-size:12px;border-bottom:1px solid var(--line);}
/* 칸 지정 2개 AND 검색줄 (입출고 표 / 재고 표 공용) */
.srch{display:flex;flex-wrap:wrap;align-items:center;gap:7px;padding:12px 16px;border-bottom:1px solid var(--line);}
.srch .fsel{font-family:inherit;font-size:12.5px;font-weight:700;color:var(--ink);background:#fff;
  border:1.5px solid var(--line);border-radius:9px;padding:8px 9px;outline:none;cursor:pointer;}
.srch .fval{font-family:inherit;font-size:13px;padding:8px 12px;border:1.5px solid var(--line);
  border-radius:9px;outline:none;min-width:130px;flex:1 1 130px;max-width:230px;}
.srch .fval:focus,.srch .fsel:focus{border-color:var(--red);}
.srch .andlab{font-size:11px;font-weight:800;color:var(--mut);letter-spacing:.5px;}
.srch .go{border:none;background:var(--red);color:#fff;font-family:inherit;font-size:12.5px;
  font-weight:800;padding:9px 16px;border-radius:9px;cursor:pointer;}
.srch .go:hover{background:var(--red2);}
.srch .off{border:1.5px solid var(--line);background:#fff;font-family:inherit;font-size:12.5px;
  font-weight:700;color:#777;padding:8px 13px;border-radius:9px;cursor:pointer;}
.srch .off:hover{border-color:#cfcfd6;color:var(--ink);}
.srch .hit{font-size:12px;color:var(--mut);font-weight:600;}
/* 정렬 가능한 헤더 */
th.s{cursor:pointer;user-select:none;white-space:nowrap;}
th.s:hover{color:var(--red);}
th.s .ar{font-size:9px;color:#c4c4cc;margin-left:3px;}
th.s.on{color:var(--red);}
th.s.on .ar{color:var(--red);}
td.dt{font-variant-numeric:tabular-nums;color:#666;white-space:nowrap;}
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
th.old{color:var(--red);}                     /* 장기재고에 해당하는 Datecode 연도 칼럼 */
td.dcy,th.dcy{white-space:nowrap;padding-left:9px;padding-right:9px;}
/* 메모(엑셀 셀 메모)는 datecode 나열처럼 긴 게 섞여 있다 — 줄여 보여주고 전체는 툴팁 */
.fillnote{margin:10px 0 0;padding:9px 14px;border-radius:8px;background:#f4f6f7;
  border-left:3px solid #b9bec3;color:#5f666d;font-size:12px;line-height:1.5;}
.fillnote b{color:#343a40;}
td.memo{max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
  color:#8a6d3b;background:#fffdf5;cursor:help;}
td.nw{white-space:nowrap;}                    /* 칼럼이 늘어난 재고 표에서 줄바꿈 방지 */
.warn{color:var(--amber);font-weight:900;margin-left:4px;cursor:help;}
/* 재고 표는 Datecode 연도 칼럼이 붙어 18칸이 된다 — 1280px 안에서는 뒤쪽 연도가 잘려서
   화면 폭이 허락하는 만큼 이 화면만 넓게 쓴다 (좁은 노트북에서는 가로 스크롤로 떨어진다). */
#invview{max-width:min(1700px,98vw);}
body.wideview #viewnav{max-width:min(1700px,98vw);}   /* 표만 넓어져 탭이 어긋나지 않게 */
/* 칼럼이 많아 가로 스크롤이 생기면 # 과 PART# 는 왼쪽에 붙여 둔다 — 2026 칸까지
   밀어놓고 보면 어느 품목 줄인지 알 수 없어진다. --c1 은 그릴 때마다 실측해 넣는다. */
#invtable td.stk1,#invtable th.stk1{position:sticky;left:0;z-index:2;background:#fff;}
#invtable td.stk2,#invtable th.stk2{position:sticky;left:var(--c1,44px);z-index:2;
  background:#fff;box-shadow:1px 0 0 var(--line);}
#invtable thead th.stk1,#invtable thead th.stk2{z-index:4;background:#fafafb;}
#invtable tr.oldrow td.stk1,#invtable tr.oldrow td.stk2{background:#fdf1f1;}
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
.tablewrap{background:#fff;border-radius:0 16px 16px 16px;box-shadow:0 1px 2px rgba(0,0,0,.04),0 6px 22px rgba(0,0,0,.05);overflow:hidden;margin-bottom:26px;}
.scroll{max-height:520px;overflow:auto;}
table{width:100%;border-collapse:collapse;font-size:12.5px;}
thead th{position:sticky;top:0;background:#fafafb;color:var(--mut);font-weight:700;font-size:10.5px;text-transform:uppercase;letter-spacing:.4px;text-align:left;padding:11px 14px;border-bottom:1.5px solid var(--line);z-index:1;}
tbody td{padding:10px 14px;border-bottom:1px solid var(--line);}
tbody tr:hover{background:#fbfafc;}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;}
td.part{font-family:ui-monospace,Consolas,monospace;font-size:11.5px;}
td.po{font-family:ui-monospace,Consolas,monospace;font-size:11.5px;color:#555;white-space:nowrap;}
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
      <div class="h" id="pwhint">.xlsx · 비밀번호는 서버에 저장돼 있어 입력하지 않아도 됩니다 · 파일명의 (영업N실)로 실 구분</div>
      <input id="file" type="file" accept=".xlsx" multiple style="display:none">
    </div>
    <!-- 비밀번호는 서버(DEFAULT_PW)에 있다. 다른 비밀번호를 쓰는 파일일 때만 펼쳐서 입력한다. -->
    <div class="pwrow" id="pwrow" style="display:none">
      <span>🔒 비밀번호</span><input id="pw" type="text" placeholder="다른 비밀번호를 쓸 때만"></div>
    <div class="pwtoggle" id="pwtoggle"
         onclick="document.getElementById('pwrow').style.display='flex';this.style.display='none';document.getElementById('pw').focus()">
      비밀번호가 다른 파일인가요?</div>
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
        <button id="vw-sm" onclick="showView('sm')">📋 전체 출고 원장</button>
      </div>
      <button id="btn-xls" class="xlsbtn" onclick="exportXlsx()">⬇ Excel 내보내기</button>
    </div>
  </div>

  <div class="wrap" id="ioview">
    <div class="offseg" id="offseg"></div>
    <div class="dayseg" id="dayseg"></div>

    <!-- 기간 조회 — 날짜 시트는 오늘·어제뿐이라 누적 시트를 서버가 걸러 준다 -->
    <div class="logbar" id="logbar" style="display:none">
      <span class="lab">기간</span>
      <span class="dpair"><input type="date" id="lg-from"><span class="tilde">~</span><input type="date" id="lg-to"></span>
      <button class="pre" data-k="m"   onclick="logPreset('m')">이번달</button>
      <button class="pre" data-k="m3"  onclick="logPreset('m3')">3개월</button>
      <button class="pre" data-k="y"   onclick="logPreset('y')">올해</button>
      <button class="pre" data-k="all" onclick="logPreset('all')">전체</button>
      <button class="off" id="lg-off" onclick="exitLog()" style="display:none">← 하루 보기</button>
      <div class="note" id="lg-note"></div>
    </div>

    <div class="kpis" id="kpis"></div>
    <div class="hl" id="hl" style="display:none"></div>
    <div class="fillnote" id="fillnote" style="display:none"></div>
    <div class="tabs">
      <button class="tabbtn on" id="tb-out" onclick="showTab('out')">출고 내역<span class="n" id="n-out"></span></button>
      <button class="tabbtn" id="tb-in" onclick="showTab('in')">입고 내역<span class="n" id="n-in"></span></button>
    </div>
    <div class="tablewrap">
      <!-- 칸 지정 2개 AND. 하루 보기와 기간 조회 양쪽에 그대로 적용된다 -->
      <div class="srch" id="io-srch">
        <select class="fsel" id="io-f1"></select>
        <input class="fval" id="io-v1" placeholder="값 입력…">
        <span class="andlab">AND</span>
        <select class="fsel" id="io-f2"></select>
        <input class="fval" id="io-v2" placeholder="값 입력…">
        <button class="go" onclick="applyIoSearch()">조회</button>
        <button class="off" onclick="clearIoSearch()">초기화</button>
        <span class="hit" id="io-hit"></span>
      </div>
      <div class="scroll" id="tablearea"></div>
    </div>
    <div class="grid" id="iogrid1">
      <div class="card"><h2>어제 vs 오늘 물동</h2><p class="desc">입고·출고 수량 비교</p><div class="cbox"><canvas id="cCompare"></canvas></div></div>
      <div class="card"><h2>오늘 출고 Top 거래처</h2><p class="desc">수량 기준 상위</p><div class="cbox"><canvas id="cCust"></canvas></div></div>
    </div>
    <div class="grid" id="iogrid2">
      <div class="card"><h2>담당자별 처리 건수</h2><p class="desc">오늘 입고+출고</p><div class="cbox"><canvas id="cSales"></canvas></div></div>
      <div class="card" id="splitcard"><h2>오늘 요약</h2><p class="desc">한눈에</p><div id="summary"></div></div>
    </div>

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

    <!-- 품목 표를 차트보다 먼저 — 입출고 화면과 같은 순서(KPI → 표 → 차트) -->
    <div class="tabs">
      <button class="tabbtn on" id="ib-all" onclick="showInvTab('all')">전체 품목<span class="n" id="in-all"></span></button>
      <button class="tabbtn" id="ib-old" onclick="showInvTab('old')">장기재고<span class="n" id="in-old"></span></button>
      <button class="tabbtn" id="ib-bk" onclick="showInvTab('bk')">예약분<span class="n" id="in-bk"></span></button>
    </div>
    <div class="tablewrap">
      <div class="srch" id="inv-srch">
        <select class="fsel" id="inv-f1"></select>
        <input class="fval" id="inv-v1" placeholder="값 입력…">
        <span class="andlab">AND</span>
        <select class="fsel" id="inv-f2"></select>
        <input class="fval" id="inv-v2" placeholder="값 입력…">
        <button class="go" onclick="drawInvTable()">조회</button>
        <button class="off" onclick="clearInvSearch()">초기화</button>
        <span class="hit" id="inv-hit"></span>
      </div>
      <div class="scroll" id="invtable"></div>
    </div>

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
    <div class="foot" id="inv-foot"></div>
  </div>

  <!-- ============ 전체 출고 원장 (shipping management) ============
       전사 공용 원장이라 실 구분이 없다. 실별 입출고와 섞지 않고 따로 본다. -->
  <div class="wrap" id="smview" style="display:none">
    <div class="logbar" id="sm-bar">
      <span class="lab">기간</span>
      <span class="dpair"><input type="date" id="sm-from"><span class="tilde">~</span><input type="date" id="sm-to"></span>
      <button class="pre" data-k="m"   onclick="smPreset('m')">이번달</button>
      <button class="pre" data-k="m3"  onclick="smPreset('m3')">3개월</button>
      <button class="pre" data-k="y"   onclick="smPreset('y')">올해</button>
      <button class="pre on" data-k="all" onclick="smPreset('all')">전체</button>
      <div class="note" id="sm-note"></div>
    </div>
    <div class="kpis" id="sm-kpis"></div>
    <div class="tablewrap">
      <div class="srch" id="sm-srch">
        <select class="fsel" id="sm-f1"></select>
        <input class="fval" id="sm-v1" placeholder="값 입력…">
        <span class="andlab">AND</span>
        <select class="fsel" id="sm-f2"></select>
        <input class="fval" id="sm-v2" placeholder="값 입력…">
        <button class="go" onclick="runSm()">조회</button>
        <button class="off" onclick="clearSmSearch()">초기화</button>
        <span class="hit" id="sm-hit"></span>
      </div>
      <div class="scroll" id="sm-table"></div>
    </div>
    <div class="foot" id="sm-foot"></div>
  </div>
</div>

<script>
const fmt=n=>(n==null?'—':Number(n).toLocaleString('ko-KR'));
const esc=s=>String(s==null?'':s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const RED="#c43a3a",BLUE="#3a6ea5",GREEN="#3f9d6b",AMBER="#e0a93a",INK="#23232b";
if(window.Chart){Chart.defaults.font.family="'Pretendard',system-ui,sans-serif";Chart.defaults.color="#6b6b74";Chart.defaults.font.size=12;}
let DATA=null, O=null, OFFI=0;
let INVS={}, INVK=null, INVNAMES=[], VIEW='io';   // 실별 재고 / 선택된 실 / 볼 수 있는 실 / 현재 뷰

// 각 실 페이지(/12 /3 /4 /5)는 자기 실 것만 본다. 재고도 마찬가지.
// selectOffice 가 주소를 바꾸므로 시작할 때 한 번 붙잡아 둔다.
const URLSLUG=location.pathname.replace(/\//g,'');

// 수신자 목록과 발송은 서버(app.py 의 OFFICE_EMAILS / send_office_emails)가 들고 있다.
// 브라우저는 결과만 받아 보여준다 — 업로드하고 탭을 닫아도 메일은 나가야 하니까.
// 서버에 엑셀 비밀번호(XLSX_PW)가 없으면 비밀번호 칸을 펼쳐 둔다 — 안 그러면 업로드가 그냥 실패한다.
fetch('/mailcfg').then(r=>r.json()).then(d=>{
  if(!d.pw){
    document.getElementById('pwrow').style.display='flex';
    document.getElementById('pwtoggle').style.display='none';
    document.getElementById('pwhint').textContent=
      '.xlsx · 비밀번호 보호 지원 · 파일명의 (영업N실)로 실 구분';
  }
}).catch(()=>{});

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
      body:JSON.stringify({files:arr, password:document.getElementById('pw').value,
                           origin:location.origin})})
    .then(r=>r.json()).then(res=>{
      document.getElementById('loading').style.display='none';
      if(res.error){ const e=document.getElementById('errmsg'); e.textContent='오류: '+res.error; e.style.display='block'; return; }
      DATA=res; renderApp(); showMailResult(res.mail);
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
  document.getElementById('foot').textContent='자료: 사내 일일 입출고 엑셀 · 하루 보기는 날짜 시트, 기간 조회는 누적 입고/출고 시트 · 파일명 (영업N실)로 실 구분';

  INVS=DATA.inventories||{};
  const hasIO=DATA.offices && DATA.offices.length;

  // 입출고에서 볼 실을 먼저 정한다 (URL 슬러그 → 없으면 첫 실)
  if(hasIO){
    let idx=0;
    if(URLSLUG){ const i=DATA.offices.findIndex(o=>officeSlug(o.name)===URLSLUG); if(i>=0) idx=i; }
    selectOffice(idx);
  }

  // 재고는 지금 보고 있는 그 실 것만 본다. 다른 실 재고는 목록에도 없고 버튼도 없다.
  const slug=hasIO?officeSlug(O.name):URLSLUG;
  INVNAMES=Object.keys(INVS).filter(n=>officeSlug(n)===slug);
  if(!INVNAMES.length && !slug) INVNAMES=Object.keys(INVS).slice(0,1);   // 재고만 올린 경우
  const hasInv=INVNAMES.length>0;

  // 전사 출고 원장은 재고 파일을 올렸을 때만 생긴다 (shipping management 시트)
  SMSPAN=(DATA.spans||{})[LEDGER]||{min:'',max:'',rows:0};
  const hasSm=SMSPAN.rows>0;
  if(hasSm) initSmBar();

  document.getElementById('viewnav').style.display=(hasIO||hasInv||hasSm)?'block':'none';
  document.getElementById('vw-io').style.display=hasIO?'':'none';
  document.getElementById('vw-inv').style.display=hasInv?'':'none';
  document.getElementById('vw-sm').style.display=hasSm?'':'none';

  // 발주업체 표를 먼저 받아 두고 재고를 그린다 (없으면 PO# 칼럼만 나오고 그대로 동작)
  if(hasInv){ INVK=INVNAMES[0]; loadPoMap().then(renderInventory); }
  showView(hasIO?'io':(hasInv?'inv':'io'));
}

function showView(v){
  VIEW=v;
  document.getElementById('ioview').style.display =v==='io' ?'block':'none';
  document.getElementById('invview').style.display=v==='inv'?'block':'none';
  document.getElementById('smview').style.display =v==='sm' ?'block':'none';
  ['io','inv','sm'].forEach(k=>document.getElementById('vw-'+k).classList.toggle('on',k===v));
  document.body.classList.toggle('wideview', v==='inv');   // 재고 표만 화면을 넓게 쓴다
  if(v==='inv') fixSticky();
  if(v==='inv'){
    const I=INVS[INVK];
    if(!I) return;
    document.getElementById('h-date').innerHTML=`재고 현황 <span class="d">·</span> ${esc(INVK)}`;
    document.getElementById('h-meta').textContent=
      `품목 ${fmt(I.n_items)}건 · 총 재고 ${fmt(I.total_qty)} EA`;
  }else if(v==='sm'){
    SMRES ? renderSm() : runSm();
  }else if(O){ LOGMODE&&LOGRES ? renderLog() : showDay(CUR); }   // 기간 조회 중이었으면 그대로 복귀
}

// ── Excel 내보내기 (지금 보고 있는 화면 그대로) ──
function exportXlsx(){
  const b=document.getElementById('btn-xls');
  const body = VIEW==='inv'
    ? {office:INVK, view:'inv', inventory:INVS[INVK]}
    : VIEW==='sm'
      ? {office:LEDGER, view:'log', ...SMARGS}
      : LOGMODE                                 // 기간 조회는 화면 상한과 무관하게 전체가 나간다
        ? {office:O.name, view:'log', ...LOGARGS}
        : {office:O.name, view:'io', day:CURDAY||O.days[CUR]};
  const label = VIEW==='inv' ? `재고_${INVK}`
    : VIEW==='sm' ? `출고원장_${SMARGS.from||'처음'}_${SMARGS.to||'끝'}`
    : LOGMODE ? `입출고이력_${O.name}_${LOGARGS.from||'처음'}_${LOGARGS.to||'끝'}`
    : `입출고_${O.name}_${(CURDAY||O.days[CUR]).date}`;
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
  // 파일에 있는 날짜(오늘·어제)는 버튼으로, 그 이전은 달력으로 — 누적 이력에서 되살린다
  const sp=logSpan();
  document.getElementById('dayseg').innerHTML=O.dates.map((d,i)=>{
    const lab=i===O.today_idx?'오늘':(i===O.today_idx-1?'어제':'');
    return `<button onclick="showDay(${i})">${esc(d)}${lab?'<span class="tg">'+lab+'</span>':''}</button>`;
  }).join('')+(sp.rows?
    `<span class="daypickwrap">📅 <input type="date" id="daypick" title="과거 날짜 보기"
       min="${sp.min}" max="${sp.max}" onchange="showPickedDay(this.value)"></span>`:'');
  drawCompare();
  initLogbar();
  showDay(O.today_idx);
}

function officeSlug(name){ const d=(String(name).match(/\d/g)||[]).join(''); return d||'all'; }
function officeLink(){ return location.origin+'/'+officeSlug(O.name); }

// ── 자동 발송 결과 표시 ──
// 발송 자체는 서버가 /build 안에서 이미 끝냈다. 여기서는 결과만 보여준다.
// 그래서 이 배너를 못 보고 탭을 닫아도 메일은 이미 나간 뒤다.
function showMailResult(m){
  const box=document.getElementById('mailnotice');
  const st=document.getElementById('emailstatus');
  if(!box||!st||!m) return;
  const ok=m.ok||[], fail=m.fail||[];
  if(!ok.length && !fail.length && !m.error) return;
  st.innerHTML=
    (m.error?`⚠️ 자동 발송 안 됨 — ${esc(m.error)}`:'')+
    (ok.length?`✅ 자동 발송 완료 — <b>${esc(ok.join(', '))}</b>`:'')+
    (fail.length?`${ok.length?' · ':''}❌ 실패: ${esc(fail.join(', '))}`:'');
  box.style.display='flex';
}

function showUpload(){
  document.getElementById('dash').style.display='none';
  document.getElementById('land').style.display='flex';
}
// 저장된 데이터가 있으면 업로드 없이 바로 보여줌 (링크 공유용)
fetch('/data').then(r=>r.json()).then(d=>{
  if(d && d.offices && d.offices.length){ DATA=d; renderApp(); }
}).catch(()=>{});

// 업로드한 파일에 있는 날짜(오늘·어제)는 그대로 쓰고,
// 그보다 과거는 누적 이력에서 서버가 되살려 준다 — 같은 renderDay 를 탄다.
function showDay(idx){
  CUR=idx; PICKED='';
  const el=document.getElementById('daypick'); if(el) el.value='';
  renderDay(O.days[idx], idx>0?O.days[idx-1]:null,
            idx===O.today_idx?'오늘':(idx===O.today_idx-1?'어제':'선택일'), idx);
}

let PICKED='';                              // 달력으로 고른 과거 날짜 ('' 면 파일의 날짜를 보는 중)
function showPickedDay(date){
  if(!date) return;
  PICKED=date;
  const pk=document.getElementById('daypick');
  if(pk) pk.value=date;                     // 코드로 불렀을 때도 달력이 그 날짜를 가리키게
  const i=O.dates.indexOf(date);
  if(i>=0){ showDay(i); if(pk) pk.value=date; PICKED=date; return; }
  document.getElementById('tablearea').innerHTML=
    '<div style="padding:40px;text-align:center;color:#aaa;font-size:13px">불러오는 중…</div>';
  fetch(`/day?office=${encodeURIComponent(O.name)}&date=${date}`)
    .then(r=>r.json()).then(res=>{
      if(res.error) throw new Error(res.error);
      if(!res.day || (!res.day.in_rows.length && !res.day.out_rows.length)){
        document.getElementById('tablearea').innerHTML=
          `<div style="padding:40px;text-align:center;color:#aaa;font-size:13px">${esc(date)} 자료가 없습니다 (휴일이거나 미기재)</div>`;
        return;
      }
      renderDay(res.day, res.prev, '선택일', -1);
    }).catch(e=>{
      document.getElementById('tablearea').innerHTML=
        `<div style="padding:40px;text-align:center;color:var(--red);font-size:13px">불러오기 실패 — ${esc(e.message)}</div>`;
    });
}

function renderDay(day, prev, tag, idx){
  exitLogUI();
  const k=day.kpi, pk=prev?prev.kpi:null;
  document.querySelectorAll('#dayseg button').forEach((b,i)=>b.classList.toggle('on',i===idx));
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

  // 날짜 시트가 비어 누적 시트로 채운 날은 숫자의 출처를 밝힌다
  const fnEl=document.getElementById('fillnote');
  const fl=(day.filled||[]);
  if(fl.length){
    const ko={in:'입고',out:'출고'};
    fnEl.style.display='block';
    fnEl.innerHTML=`이 날짜의 <b>${fl.map(x=>ko[x]||x).join(' · ')}</b> 는 날짜 시트가 비어 있어 `
      +`같은 파일의 <b>누적 ${fl.map(x=>ko[x]||x).join('/')} 시트</b>에서 가져왔습니다.`;
  } else fnEl.style.display='none';

  CURDAY=day;                               // 정렬·검색이 다시 그릴 때 쓸 현재 하루
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
function noHit(){return '<div style="padding:40px;text-align:center;color:#aaa;font-size:13px">조건에 맞는 내역이 없습니다</div>';}

const IO_ALL=r=>[r.customer,r.part,r.sales,r.doc,r.remark,r.fab,r.lot,r.dcode,r.waybill,r.memo];
// lot·DATECODE 는 shipping management 에서만, 운송장번호는 날짜 시트에서만 온다.
// 값이 있는 표에서만 칼럼을 띄운다 — 빈 칸만 늘어놓지 않도록.
const hasLot=rows=>rows.some(r=>r.lot||r.dcode);
const hasWb =rows=>rows.some(r=>r.waybill);
// 메모 칸은 내용이 하나라도 있을 때만 낸다 (메모 안 쓰는 실에서 빈 칸이 늘지 않게)
const hasMemo=rows=>rows.some(r=>r.memo);
const memoTd=v=>'<td class="memo" title="'+esc(v||'')+'">'+(esc(v||'')||'&mdash;')+'</td>';
let DSORT={k:'',d:0};                      // 하루 보기 정렬 (기본 = 엑셀 순서)
let CURDAY=null;                           // 지금 보고 있는 하루 (파일의 날짜일 수도, 달력으로 고른 과거일 수도)

function sortDay(key){ nextDir(DSORT,key); buildTables(CURDAY); showTab(TBcur); }

function buildTables(day){
  const conds=condsOf('io');
  let outR=applyConds(day.out_rows, conds, IO_ALL);
  let inR =applyConds(day.in_rows,  conds, IO_ALL);
  if(DSORT.d!==0){ outR=[...outR].sort(cmpBy(DSORT.k,DSORT.d)); inR=[...inR].sort(cmpBy(DSORT.k,DSORT.d)); }
  document.getElementById('n-out').textContent=outR.length;
  document.getElementById('n-in').textContent=inR.length;
  document.getElementById('io-hit').textContent=
    conds.length? `출고 ${hitText(outR.length,day.out_rows.length)} · 입고 ${hitText(inR.length,day.in_rows.length)}` : '';

  const S=(l,k,c)=>th(l,k,DSORT,'sortDay',c);
  const empty=day.out_rows.length||day.in_rows.length?noHit():emptyMsg();
  const L=hasLot(day.out_rows), W=hasWb(day.out_rows);
  const MO=hasMemo(day.out_rows), MI=hasMemo(day.in_rows);
  const out=!outR.length?empty:`<table><thead><tr><th>#</th>${S('거래처','customer')}${S('PART#','part')}${S('수량','qty','n')}${S('담당','sales')}${L?S('lot number','lot')+S('DATECODE','dcode'):''}${W?S('운송장번호','waybill'):''}${L?'':S('문서번호','doc')}${S('비고','remark')}${MO?S('메모','memo'):''}</tr></thead><tbody>${
    outR.map((r,i)=>`<tr><td class="n">${i+1}</td><td><b>${esc(r.customer)||'—'}</b></td><td class="part">${esc(r.part)}</td>
      <td class="qty">${fmt(r.qty)}</td><td>${esc(r.sales)}</td>
      ${L?`<td class="po">${esc(r.lot)||'—'}</td><td class="dt">${esc(r.dcode)||'—'}</td>`:''}
      ${W?`<td class="po">${esc(r.waybill)||'—'}</td>`:''}
      ${L?'':`<td>${esc(r.doc)||'—'}</td>`}<td style="color:#888">${esc(r.remark)||''}</td>${MO?memoTd(r.memo):''}</tr>`).join('')}</tbody></table>`;
  const inn=!inR.length?empty:`<table><thead><tr><th>#</th>${S('거래처/공급','customer')}${S('PART#','part')}${S('수량','qty','n')}${S('담당','sales')}${S('FAB','fab')}${S('비고','remark')}${MI?S('메모','memo'):''}</tr></thead><tbody>${
    inR.map((r,i)=>`<tr><td class="n">${i+1}</td><td><b>${esc(r.customer)||'—'}</b></td><td class="part">${esc(r.part)}</td>
      <td class="qty">${fmt(r.qty)}</td><td>${esc(r.sales)}</td><td>${esc(r.fab)?'<span class="pill">FAB '+esc(r.fab)+'</span>':'—'}</td>
      <td style="color:#888">${esc(r.remark)||''}</td>${MI?memoTd(r.memo):''}</tr>`).join('')}</tbody></table>`;
  TB={out,in:inn};
}

// 검색줄은 하루 보기와 기간 조회가 같이 쓴다 — 어느 쪽이 떠 있느냐에 따라 갈린다
function applyIoSearch(){
  if(LOGMODE) runLog();
  else if(CURDAY){ buildTables(CURDAY); showTab(TBcur); }
}
function clearIoSearch(){ clearFields('io'); applyIoSearch(); }
function showTab(which){
  TBcur=which;
  document.getElementById('tb-out').classList.toggle('on',which==='out');
  document.getElementById('tb-in').classList.toggle('on',which==='in');
  document.getElementById('tablearea').innerHTML=TB[which]||'';
}

// ================= 칸 지정 검색 + 헤더 정렬 (표 3개 공용) =================
// 검색은 [칸][값] AND [칸][값] 두 줄. '전체' 는 모든 칸을 훑는다.
// 정렬은 헤더 클릭으로 오름 → 내림 → 기본 순환. 기간 조회만 서버가 정렬하는데,
// 3,000행 상한이 걸린 뒤 화면에서 정렬하면 그 3,000행 안에서만 섞이기 때문이다.
const F_IO =[['all','전체'],['customer','거래처'],['part','PART#'],
             ['sales','담당'],['waybill','운송장번호'],['doc','문서번호'],
             ['lot','lot number'],['dcode','DATECODE'],['remark','비고'],['memo','메모']];
const F_INV=[['all','전체'],['part','PART#'],['customer','PO# / 고객'],['buyer','발주업체'],
             ['mobis','MOBIS ID'],['family','FAMILY'],['vender','VENDER'],
             ['sales','담당'],['office','실'],['memo','메모']];

function fillFields(prefix, fields, d1, d2){
  const opt=f=>fields.map(([v,l])=>`<option value="${v}"${v===f?' selected':''}>${l}</option>`).join('');
  document.getElementById(prefix+'-f1').innerHTML=opt(d1);
  document.getElementById(prefix+'-f2').innerHTML=opt(d2);
}
function condsOf(prefix){
  const out=[];
  for(const n of ['1','2']){
    const v=(document.getElementById(prefix+'-v'+n).value||'').trim();
    if(v) out.push({f:document.getElementById(prefix+'-f'+n).value, v});
  }
  return out;
}
function clearFields(prefix){
  document.getElementById(prefix+'-v1').value='';
  document.getElementById(prefix+'-v2').value='';
}
// 한 조건 안의 띄어쓰기는 다시 AND. 칸이 'all' 이면 준 값들 전체를 훑는다.
function condMatch(row, c, allOf){
  const hay=(c.f==='all'? allOf(row) : [row[c.f]]).map(v=>String(v==null?'':v).toUpperCase());
  return c.v.trim().split(/\s+/).every(t=>hay.some(h=>h.includes(t.toUpperCase())));
}
function applyConds(rows, conds, allOf){
  return conds.length? rows.filter(r=>conds.every(c=>condMatch(r,c,allOf))) : rows;
}

const NUMCOL={qty:1,avail:1,booking:1,old:1};
const isNumCol=k=>NUMCOL[k]||/^y20\d\d$/.test(k);      // y2019… = Datecode 연도별 수량
function cmpBy(k,d){
  return (a,b)=>{
    if(isNumCol(k)) return ((Number(a[k])||0)-(Number(b[k])||0))*d;
    return String(a[k]==null?'':a[k]).localeCompare(String(b[k]==null?'':b[k]),'ko',{numeric:true})*d;
  };
}
function nextDir(st,key){
  if(st.k!==key){ st.k=key; st.d=1; }
  else if(st.d===1) st.d=-1;
  else { st.k=''; st.d=0; }              // 세 번째 클릭 = 기본 정렬로 복귀
}
function th(label, key, st, fn, cls){
  const on=st.k===key && st.d!==0;
  return `<th class="s${on?' on':''}${cls?' '+cls:''}" onclick="${fn}('${key}')">`+
         `${label}<span class="ar">${!on?'⇅':(st.d>0?'▲':'▼')}</span></th>`;
}
function hitText(shown, total){
  return shown===total ? `${fmt(total)}건` : `${fmt(shown)}건 / 전체 ${fmt(total)}건`;
}

// ================= 기간 조회 (입출고 이력) =================
// 실당 6만 행이라 브라우저로 다 못 내린다. 조건만 서버에 보내고 걸러진 것만 받는다.
// 하루 보기(날짜 버튼)와 같은 표 자리를 쓰되, 하루짜리가 아닌 기간 집계를 보여준다.
let LOGMODE=false, LOGRES=null, LOGARGS={from:'',to:'',conds:[],sort:'date',dir:'desc'}, LOGTMR=null;

const iso=d=>new Date(d.getTime()-d.getTimezoneOffset()*6e4).toISOString().slice(0,10);
function isoAdd(s,days){ const d=new Date(s+'T00:00:00'); d.setDate(d.getDate()+days); return iso(d); }
function logSpan(){ return (DATA.spans||{})[O.name]||{min:'',max:'',rows:0}; }

function initLogbar(){
  const sp=logSpan(), bar=document.getElementById('logbar');
  if(!sp.rows){ bar.style.display='none'; return; }      // 이력 없는 실은 기간 조회 자체를 숨긴다
  bar.style.display='flex';
  const f=document.getElementById('lg-from'), t=document.getElementById('lg-to');
  f.min=t.min=sp.min; f.max=t.max=sp.max;
  document.getElementById('lg-note').textContent=
    `이력 ${sp.min} ~ ${sp.max} · ${fmt(sp.rows)}행 (누적 입고/출고 시트) · 기간을 고르면 그 구간 전체를 합산해 보여줍니다`;
  logPreset('m', true);
}

// 기준점은 오늘이 아니라 '이력의 마지막 날'이다 — 누적 시트는 하루 이틀 늦게 채워진다.
function logPreset(kind, quiet){
  const sp=logSpan(), end=sp.max||iso(new Date());
  let start=end;
  if(kind==='m')   start=end.slice(0,8)+'01';
  if(kind==='m3')  start=isoAdd(end,-90);
  if(kind==='y')   start=end.slice(0,4)+'-01-01';
  if(kind==='all') start=sp.min||end;
  if(sp.min && start<sp.min) start=sp.min;
  document.getElementById('lg-from').value=start;
  document.getElementById('lg-to').value=end;
  document.querySelectorAll('#logbar .pre').forEach(b=>b.classList.toggle('on',b.dataset.k===kind));
  if(!quiet) runLog();
}

let LSORT={k:'',d:0};                      // 기간 조회 정렬 (기본 = 일자 내림차순, 서버가 처리)
function sortLog(key){ nextDir(LSORT,key); runLog(); }

function runLog(){
  const from=document.getElementById('lg-from').value,
        to  =document.getElementById('lg-to').value,
        conds=condsOf('io');
  if(from && to && from>to){
    document.getElementById('tablearea').innerHTML=
      '<div style="padding:40px;text-align:center;color:var(--red);font-size:13px">시작일이 종료일보다 뒤입니다</div>';
    return;
  }
  const sort=LSORT.d?LSORT.k:'date', dir=LSORT.d>0?'asc':'desc';
  LOGARGS={from,to,conds,sort,dir};
  document.getElementById('tablearea').innerHTML=
    '<div style="padding:40px;text-align:center;color:#aaa;font-size:13px">조회 중…</div>';
  const p=new URLSearchParams({office:O.name, from, to, sort, dir});
  conds.forEach((c,i)=>{ p.set('f'+(i+1),c.f); p.set('v'+(i+1),c.v); });
  fetch('/log?'+p.toString())
    .then(r=>r.json()).then(res=>{
      if(res.error) throw new Error(res.error);
      LOGRES=res; LOGMODE=true; renderLog();
    }).catch(e=>{
      document.getElementById('tablearea').innerHTML=
        `<div style="padding:40px;text-align:center;color:var(--red);font-size:13px">조회 실패 — ${esc(e.message)}</div>`;
    });
}

function renderLog(){
  const r=LOGRES, {from,to,conds}=LOGARGS;
  const q=conds.map(c=>`${(F_IO.find(f=>f[0]===c.f)||['','전체'])[1]}='${c.v}'`).join(' AND ');
  document.querySelectorAll('#dayseg button').forEach(b=>b.classList.remove('on'));
  document.getElementById('lg-off').style.display='';
  document.getElementById('hl').style.display='none';
  document.getElementById('iogrid1').style.display='none';   // 차트는 하루 기준이라 기간에선 의미가 없다
  document.getElementById('iogrid2').style.display='none';

  document.getElementById('h-date').innerHTML=
    `${esc(O.name)} <span class="d">·</span> ${esc(from||'처음')} ~ ${esc(to||'끝')}`;
  document.getElementById('h-meta').textContent=
    `기간 조회${q?` · 검색 "${q}"`:''} · 입고 ${fmt(r.in.cnt)}건 / 출고 ${fmt(r.out.cnt)}건`;
  document.getElementById('kpis').innerHTML=[
    ['in','입고 건수',r.in.cnt,'건'],['in','입고 수량',r.in.qty,'EA'],
    ['out','출고 건수',r.out.cnt,'건'],['out','출고 수량',r.out.qty,'EA'],
    ['net','순물동(입-출)',r.net,'EA'],['cu','거래처(출고)',r.out.customers,'곳'],
  ].map(([c,l,v,u])=>`<div class="kpi ${c}"><div class="l">${l}</div>
     <div class="v tab">${fmt(v)}<span class="u">${u}</span></div><span class="d fl">&nbsp;</span></div>`).join('');

  document.getElementById('n-out').textContent=r.out.cnt;
  document.getElementById('n-in').textContent=r.in.cnt;
  document.getElementById('io-hit').textContent=`출고 ${fmt(r.out.cnt)}건 · 입고 ${fmt(r.in.cnt)}건`;

  const none=`<div style="padding:40px;text-align:center;color:#aaa;font-size:13px">조건에 맞는 내역이 없습니다</div>`;
  const cap=s=>s.truncated?`<div class="trunc">전체 ${fmt(s.cnt)}건 중 ${fmt(s.shown)}건만 표시합니다 —
    기간을 좁히거나 검색 조건을 넣어 보세요. <b>⬇ Excel 내보내기는 전체가 나갑니다.</b></div>`:'';
  const S=(l,k,c)=>th(l,k,LSORT,'sortLog',c);
  const L=hasLot(r.out.rows), W=hasWb(r.out.rows);
  const MO=hasMemo(r.out.rows), MI=hasMemo(r.in.rows);

  TB={
    out: !r.out.rows.length?none:cap(r.out)+`<table><thead><tr><th>#</th>${S('일자','date')}${S('거래처','customer')}${S('PART#','part')}${S('수량','qty','n')}${S('담당','sales')}${L?S('lot number','lot')+S('DATECODE','dcode'):''}${W?S('운송장번호','waybill'):''}${L?'':S('문서번호','doc')}${S('비고','remark')}${MO?S('메모','memo'):''}</tr></thead><tbody>${
      r.out.rows.map((x,i)=>`<tr><td class="n">${i+1}</td><td class="dt">${esc(x.date)}</td>
        <td><b>${esc(x.customer)||'—'}</b></td><td class="part">${esc(x.part)}</td>
        <td class="qty">${fmt(x.qty)}</td><td>${esc(x.sales)}</td>
        ${L?`<td class="po">${esc(x.lot)||'—'}</td><td class="dt">${esc(x.dcode)||'—'}</td>`:''}
        ${W?`<td class="po">${esc(x.waybill)||'—'}</td>`:''}
        ${L?'':`<td>${esc(x.doc)||'—'}</td>`}<td style="color:#888">${esc(x.remark)||''}</td>${MO?memoTd(x.memo):''}</tr>`).join('')}</tbody></table>`,
    in: !r.in.rows.length?none:cap(r.in)+`<table><thead><tr><th>#</th>${S('일자','date')}${S('거래처/공급','customer')}${S('PART#','part')}${S('수량','qty','n')}${S('담당','sales')}${S('FAB','fab')}${S('비고','remark')}${MI?S('메모','memo'):''}</tr></thead><tbody>${
      r.in.rows.map((x,i)=>`<tr><td class="n">${i+1}</td><td class="dt">${esc(x.date)}</td>
        <td><b>${esc(x.customer)||'—'}</b></td><td class="part">${esc(x.part)}</td>
        <td class="qty">${fmt(x.qty)}</td><td>${esc(x.sales)}</td>
        <td>${esc(x.fab)?'<span class="pill">FAB '+esc(x.fab)+'</span>':'—'}</td>
        <td style="color:#888">${esc(x.remark)||''}</td>${MI?memoTd(x.memo):''}</tr>`).join('')}</tbody></table>`,
  };
  showTab(TBcur);
}

// 하루 보기로 돌아갈 때 기간 조회의 흔적만 걷어낸다 (showDay 가 나머지를 다시 그린다)
function exitLogUI(){
  LOGMODE=false;
  const off=document.getElementById('lg-off');
  if(off) off.style.display='none';
  const g1=document.getElementById('iogrid1'), g2=document.getElementById('iogrid2');
  if(g1) g1.style.display=''; if(g2) g2.style.display='';
}
function exitLog(){ LSORT={k:'',d:0}; showDay(CUR); }

document.addEventListener('DOMContentLoaded',()=>{
  fillFields('io', F_IO, 'all', 'part');
  fillFields('inv', F_INV, 'all', 'part');
  fillFields('sm', F_IO, 'all', 'part');
  ['sm-v1','sm-v2'].forEach(id=>{
    const el=document.getElementById(id);
    if(!el) return;
    el.addEventListener('input',()=>{ clearTimeout(SMTMR); SMTMR=setTimeout(runSm,350); });
    el.addEventListener('keydown',e=>{ if(e.key==='Enter'){ clearTimeout(SMTMR); runSm(); } });
  });
  ['sm-f1','sm-f2'].forEach(id=>{
    const el=document.getElementById(id);
    if(el) el.addEventListener('change',runSm);
  });
  ['sm-from','sm-to'].forEach(id=>{
    const el=document.getElementById(id);
    if(el) el.addEventListener('change',()=>{
      document.querySelectorAll('#sm-bar .pre').forEach(b=>b.classList.remove('on'));
      runSm();
    });
  });
  // 타이핑이 멎으면 자동 조회 — 조회 버튼을 누르지 않아도 되게
  ['io-v1','io-v2'].forEach(id=>{
    const el=document.getElementById(id);
    if(!el) return;
    el.addEventListener('input',()=>{ clearTimeout(LOGTMR); LOGTMR=setTimeout(applyIoSearch,350); });
    el.addEventListener('keydown',e=>{ if(e.key==='Enter'){ clearTimeout(LOGTMR); applyIoSearch(); } });
  });
  ['io-f1','io-f2'].forEach(id=>{
    const el=document.getElementById(id);
    if(el) el.addEventListener('change',applyIoSearch);
  });
  ['inv-v1','inv-v2'].forEach(id=>{
    const el=document.getElementById(id);
    if(!el) return;
    el.addEventListener('input',()=>{ clearTimeout(LOGTMR); LOGTMR=setTimeout(drawInvTable,250); });
    el.addEventListener('keydown',e=>{ if(e.key==='Enter'){ clearTimeout(LOGTMR); drawInvTable(); } });
  });
  ['inv-f1','inv-f2'].forEach(id=>{
    const el=document.getElementById(id);
    if(el) el.addEventListener('change',drawInvTable);
  });
  ['lg-from','lg-to'].forEach(id=>{
    const el=document.getElementById(id);
    if(el) el.addEventListener('change',()=>{
      document.querySelectorAll('#logbar .pre').forEach(b=>b.classList.remove('on'));
      runLog();
    });
  });
});

// ================= 전체 출고 원장 (shipping management) =================
// 전사 공용 원장이라 실 소속이 없다. 서버는 office=LEDGER 로 이 범위만 조회한다.
const LEDGER='전체 출고 원장';
let SMSORT={k:'',d:0}, SMSPAN={min:'',max:'',rows:0}, SMRES=null, SMARGS=null, SMTMR=null;

function sortSm(key){ nextDir(SMSORT,key); runSm(); }
function clearSmSearch(){ clearFields('sm'); runSm(); }

function initSmBar(){
  const f=document.getElementById('sm-from'), t=document.getElementById('sm-to');
  f.min=t.min=SMSPAN.min; f.max=t.max=SMSPAN.max;
  document.getElementById('sm-note').textContent=
    `원장 ${SMSPAN.min} ~ ${SMSPAN.max} · ${fmt(SMSPAN.rows)}행 · lot 단위 출고 내역 (재고 파일의 shipping management 시트)`+
    (SMSPAN.src?` · 이 시트는 여러 실 파일에 같은 내용으로 들어 있어 '${SMSPAN.src}' 파일 것만 씁니다`:'');
  smPreset('all', true);
}
function smPreset(kind, quiet){
  const end=SMSPAN.max||'', start=
    kind==='m' ? end.slice(0,8)+'01' :
    kind==='m3'? isoAdd(end,-90) :
    kind==='y' ? end.slice(0,4)+'-01-01' : (SMSPAN.min||end);
  document.getElementById('sm-from').value=(SMSPAN.min&&start<SMSPAN.min)?SMSPAN.min:start;
  document.getElementById('sm-to').value=end;
  document.querySelectorAll('#sm-bar .pre').forEach(b=>b.classList.toggle('on',b.dataset.k===kind));
  if(!quiet) runSm();
}

function runSm(){
  const from=document.getElementById('sm-from').value,
        to  =document.getElementById('sm-to').value,
        conds=condsOf('sm');
  if(from && to && from>to){
    document.getElementById('sm-table').innerHTML=
      '<div style="padding:40px;text-align:center;color:var(--red);font-size:13px">시작일이 종료일보다 뒤입니다</div>';
    return;
  }
  const sort=SMSORT.d?SMSORT.k:'date', dir=SMSORT.d>0?'asc':'desc';
  SMARGS={from,to,conds,sort,dir};
  document.getElementById('sm-table').innerHTML=
    '<div style="padding:40px;text-align:center;color:#aaa;font-size:13px">조회 중…</div>';
  const p=new URLSearchParams({office:LEDGER, from, to, sort, dir});
  conds.forEach((c,i)=>{ p.set('f'+(i+1),c.f); p.set('v'+(i+1),c.v); });
  fetch('/log?'+p.toString()).then(r=>r.json()).then(res=>{
    if(res.error) throw new Error(res.error);
    SMRES=res; renderSm();
  }).catch(e=>{
    document.getElementById('sm-table').innerHTML=
      `<div style="padding:40px;text-align:center;color:var(--red);font-size:13px">조회 실패 — ${esc(e.message)}</div>`;
  });
}

function renderSm(){
  const o=SMRES.out, {from,to,conds}=SMARGS;
  const q=conds.map(c=>`${(F_IO.find(f=>f[0]===c.f)||['','전체'])[1]}='${c.v}'`).join(' AND ');
  document.getElementById('h-date').innerHTML=
    `전체 출고 원장 <span class="d">·</span> ${esc(from||'처음')} ~ ${esc(to||'끝')}`;
  document.getElementById('h-meta').textContent=
    `${q?q+' · ':''}출고 ${fmt(o.cnt)}건 / ${fmt(o.qty)} EA · 거래처 ${fmt(o.customers)}곳`;
  document.getElementById('sm-kpis').innerHTML=[
    ['out','출고 건수',o.cnt,'건'],['out','출고 수량',o.qty,'EA'],['cu','거래처',o.customers,'곳'],
  ].map(([c,l,v,u])=>`<div class="kpi ${c}"><div class="l">${l}</div>
     <div class="v tab">${fmt(v)}<span class="u">${u}</span></div><span class="d fl">&nbsp;</span></div>`).join('');
  document.getElementById('sm-hit').textContent=conds.length?`${fmt(o.cnt)}건`:'';

  const S=(l,k,c)=>th(l,k,SMSORT,'sortSm',c);
  const cap=o.truncated?`<div class="trunc">전체 ${fmt(o.cnt)}건 중 ${fmt(o.shown)}건만 표시합니다 —
    기간을 좁히거나 검색 조건을 넣어 보세요. <b>⬇ Excel 내보내기는 전체가 나갑니다.</b></div>`:'';
  document.getElementById('sm-table').innerHTML = !o.rows.length
    ? '<div style="padding:40px;text-align:center;color:#aaa;font-size:13px">조건에 맞는 내역이 없습니다</div>'
    : cap+`<table><thead><tr><th>#</th>${S('일자','date')}${S('거래처','customer')}${S('PART#','part')}${S('수량','qty','n')}${S('담당','sales')}${S('lot number','lot')}${S('DATECODE','dcode')}</tr></thead><tbody>${
      o.rows.map((x,i)=>`<tr><td class="n">${i+1}</td><td class="dt">${esc(x.date)}</td>
        <td><b>${esc(x.customer)||'—'}</b></td><td class="part">${esc(x.part)}</td>
        <td class="qty">${fmt(x.qty)}</td><td>${esc(x.sales)||'—'}</td>
        <td class="po">${esc(x.lot)||'—'}</td><td class="dt">${esc(x.dcode)||'—'}</td></tr>`).join('')}</tbody></table>`;
  document.getElementById('sm-foot').textContent=
    '자료: 재고 파일의 shipping management 시트 · 전사 공용 원장이라 영업실 구분이 없습니다';
}

// ================= 재고 현황 =================
let ICH={}, ITAB='all';

// PO# → 발주업체. 재고 시트의 CUSTOMER 칸은 고객명이 아니라 PO# 라서,
// 선적 리포트에서 받아온 이 표로 진짜 발주업체를 붙인다.
// 서버에서 따로 받는 이유: 재고와 선적 리포트를 다른 날 올려도 항상 최신으로 맞추기 위해서다.
let POMAP={}, POLOOSE={}, POROWS=0;
const normPo=s=>String(s||'').replace(/-\d+(-\d+)*$/,'').toUpperCase();   // 26YS0211N4-5-5 → 26YS0211N4
const POLIKE=/^\d{2}[A-Z]{2,6}\d{3,4}[A-Z]?\d*(-\d+)*$/i;                // 26DIT0213N10 꼴

// 한 PO#가 두 업체로 선적된 경우가 있어 배열로 온다. 첫째만 쓰고 나머지는 '외 N' 으로 알린다.
function buyersOf(x){
  const p=String(x.customer||'').toUpperCase();
  return POMAP[p] || POLOOSE[normPo(x.customer)] || [];
}
function buyerOf(x){ return (buyersOf(x)[0])||''; }
function buyerCell(x){
  const b=buyersOf(x);
  if(!b.length) return '<span style="color:#c8c8d0">—</span>';
  return esc(b[0])+(b.length>1?` <span class="pill" title="${esc(b.slice(1).join(', '))}">외 ${b.length-1}</span>`:'');
}

// CUSTOMER 칸의 의미가 실마다 다르다. 영업5실은 PO#(69%)와 고객명이 섞여 있고,
// 영업1,2실·4실은 거의 고객명이다. 한쪽으로 확실할 때만 그 이름을 쓰고,
// 섞여 있으면 섞였다고 적는다 — 'PO#' 로 단정하면 나머지 행에서 거짓말이 된다.
function poHeader(items){
  const v=items.map(x=>x.customer).filter(Boolean);
  if(!v.length) return 'PO# / 고객';
  const r=v.filter(s=>POLIKE.test(s)).length/v.length;
  return r>=0.9 ? 'PO#' : (r<=0.1 ? '고객' : 'PO# / 고객');
}
function loadPoMap(){
  return fetch('/po').then(r=>r.json()).then(d=>{
    POMAP=d.map||{}; POLOOSE=d.loose||{}; POROWS=d.rows||0;
  }).catch(()=>{});
}
const pct=(a,b)=>b?Math.round(a/b*1000)/10:0;

// 볼 수 있는 실만 연다 (실 페이지에서는 자기 실뿐 — 차트 드릴다운으로도 못 넘어간다)
function selectInv(name){
  if(!INVNAMES.includes(name)) return;
  INVK=name; renderInventory(); showView('inv');
}

function renderInventory(){
  const I=INVS[INVK];
  if(!I) return;

  // 보고 있는 실 하나만 — 다른 실로 가는 버튼은 없다 (입출고 쪽과 동일)
  document.getElementById('inv-offseg').innerHTML=
    `<span class="lab">영업실</span>`+
    `<button class="on" style="cursor:default" disabled>${esc(INVK)}</button>`;

  const dcOn=I.items.some(x=>x.dcq&&Object.keys(x.dcq).length);
  document.getElementById('inv-foot').textContent=
    `자료: 재고 엑셀의 '${I.sheet}' 시트 · 재고 수량이 있는 품목만 집계 (합계 행 제외) · `+
    (dcOn?`Datecode 가 적힌 품목은 DC 연도별 수량의 합 = 그 품목의 '재고' 수량 `+
          `(원본이 안 맞으면 ⚠, Datecode 가 아예 없는 품목은 전부 '—') · `:'')+
    `장기재고 = Datecode ${I.old_year}년 이전 (빨간 칼럼)`;

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
  showInvTab('all');
}

function showInvTab(t){
  ITAB=t;
  ['all','old','bk'].forEach(k=>document.getElementById('ib-'+k).classList.toggle('on',k===t));
  drawInvTable();
}

let ISORT={k:'',d:0};                      // 재고 정렬 (기본 = 재고 수량 내림차순)
function sortInv(key){ nextDir(ISORT,key); drawInvTable(); }
function clearInvSearch(){ clearFields('inv'); drawInvTable(); }

// '발주업체' 로 검색하려면 조인 결과가 행에 있어야 한다 — 필터 직전에 붙여 둔다
const INV_ALL=x=>[x.part,x.mobis,x.family,x.vender,x.sales,x.customer,x.pn,x.office,x.buyer,x.memo];

function drawInvTable(){
  const I=INVS[INVK];
  if(!I) return;
  // Datecode 연도별 수량을 y2019… 칸으로 펼쳐 둔다 — 검색·정렬이 다른 칸과 똑같이 돌게.
  let base=I.items.map(x=>{
    const o={...x, buyer:buyerOf(x)};
    let sum=0, n=0;
    for(const y in (x.dcq||{})){ o['y'+y]=x.dcq[y]; sum+=x.dcq[y]; n++; }
    // 연도별 합이 재고 수량과 안 맞는 품목이 실제로 있다 (원본 시트가 그렇다).
    // 조용히 덮지 않고 재고 칸에 표시해 둔다 — 어느 쪽이 맞는지는 사람이 판단할 몫.
    o.dcgap = (n && Math.abs(sum-x.qty)>0.5) ? sum-x.qty : 0;
    return o;
  });
  if(ITAB==='old') base=base.filter(x=>x.old>0);
  if(ITAB==='bk')  base=base.filter(x=>x.booking>0);
  const conds=condsOf('inv');
  let rows=applyConds(base, conds, INV_ALL);
  document.getElementById('inv-hit').textContent=conds.length?hitText(rows.length,base.length):'';
  // 장기재고 탭에서는 장기 수량이 많은 순 — 어디부터 손댈지 고르는 화면이라.
  const DEF=ITAB==='old'?(a,b)=>b.old-a.old:(a,b)=>b.qty-a.qty;
  rows=[...rows].sort(ISORT.d?cmpBy(ISORT.k,ISORT.d):DEF);

  const el=document.getElementById('invtable');
  if(!rows.length){ el.innerHTML=base.length?noHit():emptyMsg(); return; }
  // Datecode 칼럼은 데이터가 있는 실에서만. '장기재고' 합계 한 칸 대신 연도별로 쪼갠다 —
  // 합계만 놓으면 '재고 6,000 + 장기 3,136 = 9,136?' 처럼 읽힌다는 현장 지적이 있었다.
  // 칼럼 구성은 탭·검색과 무관하게 실 단위로 고정한다 (필터마다 칼럼이 들락거리면 읽기 힘들다).
  const OY=I.old_year;
  const ALLY=(I.years&&I.years.length)?I.years
    :[...new Set(I.items.flatMap(x=>Object.keys(x.dcq||{}).map(Number)))].sort((a,b)=>a-b);
  const YRS=ALLY.filter(y=>I.items.some(x=>x.dcq&&x.dcq[y]>0));
  // 배포 직후에는 이전 판이 저장해 둔 자료가 그대로 올라온다 — 거기엔 연도별 수량이 없다.
  // 재고 파일을 다시 올리기 전까지는 예전 '장기재고/Datecode' 칼럼으로 버틴다.
  const legacy=!YRS.length && I.datecode.some(d=>d.qty>0);
  const dc=YRS.length>0||legacy;
  // VENDER 검색은 항상 열려 있고, 칼럼만 여러 곳이 섞인 실에서 보여준다 (영업4실)
  const hasVen=new Set(I.items.map(x=>x.vender).filter(Boolean)).size>1;
  // 발주업체 칼럼은 선적 리포트를 올렸고 실제로 붙는 게 있을 때만 (영업1,2실은 전부 빈칸이라 뺀다)
  const hasBuyer=POROWS>0 && base.some(x=>x.buyer);
  const MM=hasMemo(base);
  const S=(l,k,c)=>th(l,k,ISORT,'sortInv',c);
  const YH=YRS.map(y=>S('DC '+y,'y'+y,'n dcy'+(y<=OY?' old':''))).join('');
  const YC=x=>YRS.map(y=>`<td class="n dcy${y<=OY&&x['y'+y]?' old':''}">${x['y'+y]?fmt(x['y'+y]):'—'}</td>`).join('');
  el.innerHTML=`<table><thead><tr><th class="stk1">#</th>${S('PART#','part','stk2')}${S(poHeader(I.items),'customer')}${hasBuyer?S('발주업체','buyer'):''}
    ${S('MOBIS ID','mobis')}${S('FAMILY','family')}${hasVen?S('VENDER','vender'):''}
    ${S('실','office')}${S('재고','qty','n')}${S('가용','avail','n')}${S('예약','booking','n')}${MM?S('메모','memo'):''}
    ${YH}${legacy?S('장기재고','old','n')+S('Datecode','oldest'):''}${S('담당','sales')}</tr></thead><tbody>${
    rows.map((x,i)=>`<tr class="${dc&&x.old>0?'oldrow':''}">
      <td class="n stk1">${i+1}</td>
      <td class="part stk2">${esc(x.part)}</td>
      <td class="po">${esc(x.customer)||'—'}</td>
      ${hasBuyer?`<td>${buyerCell(x)}</td>`:''}
      <td class="nw">${esc(x.mobis)||'—'}</td>
      <td>${esc(x.family)||'—'}</td>
      ${hasVen?`<td class="nw">${esc(x.vender)||'—'}</td>`:''}
      <td class="nw">${esc(x.office)||'—'}</td>
      <td class="qty">${fmt(x.qty)}${x.dcgap?`<span class="warn" title="Datecode 연도별 합계는 ${fmt(x.qty+x.dcgap)} EA 로 재고 수량과 ${fmt(Math.abs(x.dcgap))} EA 차이납니다 (원본 시트 그대로)">⚠</span>`:''}</td>
      <td class="n">${fmt(x.avail)}</td>
      <td class="n">${x.booking?fmt(x.booking):'—'}</td>
      ${MM?memoTd(x.memo):''}
      ${YC(x)}
      ${legacy?`<td class="n ${x.old>0?'old':''}">${x.old?fmt(x.old):'—'}</td>
      <td>${x.oldest?('<span class="pill">'+x.oldest+'~</span>'):'—'}</td>`:''}
      <td>${esc(x.sales)||'—'}</td></tr>`).join('')}</tbody></table>`;
  fixSticky();
}

// PART# 를 왼쪽에 붙여 둘 위치 = 실제로 그려진 '#' 칸 너비 (자리수에 따라 달라진다).
// 화면이 숨어 있을 때 재면 0 이 나오므로 뷰를 켤 때도 한 번 더 잰다.
function fixSticky(){
  const el=document.getElementById('invtable');
  const c1=el&&el.querySelector('tbody td');
  if(c1&&c1.offsetWidth) el.style.setProperty('--c1', c1.offsetWidth+'px');
}
</script></body></html>"""


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="일일 입출고 리포트 대시보드")
    ap.add_argument("--port", type=int, default=8780)
    ap.add_argument("--no-open", action="store_true")
    ap.add_argument("--dedupe-ledger", action="store_true",
                    help="전사 출고 원장에 남은 옛 중복을 세어 본다 (기본은 미리보기)")
    ap.add_argument("--apply", action="store_true",
                    help="--dedupe-ledger 와 함께: 실제로 지운다 (되돌릴 수 없음)")
    args = ap.parse_args()

    if args.dedupe_ledger:
        r = ledger_dupes(apply=args.apply)
        if not r["rows"]:
            print("전사 출고 원장: 파일 여러 벌이 겹친 흔적 없음.")
            raise SystemExit(0)
        print(f"{'지웠습니다' if args.apply else '미리보기 (아직 안 지움)'} — "
              f"{r['dates']}일 · {r['rows']:,}행 · {r['qty']:,.0f} EA")
        for d, g, n in r["detail"][:20]:
            print(f"  {d}  {g}벌 겹침 → {n:,}행")
        if len(r["detail"]) > 20:
            print(f"  … 외 {len(r['detail']) - 20}일")
        if not args.apply:
            print()
            print("실제로 지우려면: python app.py --dedupe-ledger --apply")
        raise SystemExit(0)
    env_port = os.environ.get("PORT")
    if env_port:                                   # 클라우드(Render 등): 0.0.0.0 + $PORT
        host, port = "0.0.0.0", int(env_port)
        args.no_open = True
    else:                                          # 로컬: localhost + 빈 포트 자동
        host, port = "127.0.0.1", find_free_port(args.port)
    url = f"http://localhost:{port}/"
    migrate_saved_memos()                          # 저장돼 있던 옛 재고 메모 정리
    httpd = ThreadingHTTPServer((host, port), Handler)
    print(f"\n일일 입출고 리포트 → {url} (bind {host}:{port})\n종료: Ctrl+C")
    if not args.no_open:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n종료."); httpd.shutdown()
