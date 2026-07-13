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
from collections import defaultdict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import msoffcrypto
import openpyxl

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


def norm_sales(s):
    if not s:
        return "(미지정)"
    return s.split("/")[0].strip() or "(미지정)"


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
    try:                                                 # 1) 평문 xlsx
        return openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    except Exception:
        pass
    try:                                                 # 2) 암호화 → 복호화
        off = msoffcrypto.OfficeFile(io.BytesIO(raw))
        off.load_key(password=password or "")
        dec = io.BytesIO(); off.decrypt(dec)
        return openpyxl.load_workbook(dec, read_only=True, data_only=True)
    except Exception:
        raise ValueError("엑셀을 열 수 없습니다. 비밀번호가 비었거나 틀렸을 수 있어요 — 업로드 화면의 '비밀번호' 칸에 입력하세요 (예: 9178).")


def office_block(name, per):
    """{날짜:(inb,outb)} → 실 1개 대시보드 데이터."""
    dates = sorted(per)
    days = [day_block(d, per[d][0], per[d][1]) for d in dates]
    return dict(name=name, dates=dates, today_idx=len(days) - 1, days=days)


def build_payload(files, password):
    """files: [{name, file(base64)}] → {offices:[...], inventory:{...}}

    파일 종류는 시트 이름으로 자동 판별한다.
      날짜 시트(YYYY-MM-DD) 있음  → 일일 입출고 파일
      '~ inventory' 시트 있음      → 재고 현황 파일
    """
    offices, inventory = [], None
    allbydate = defaultdict(lambda: ([], []))
    for f in files:
        b64 = f["file"]
        if "," in b64[:64]:
            b64 = b64.split(",", 1)[1]
        wb = open_wb(base64.b64decode(b64), password)
        try:
            per = parse_workbook(wb)
            if per:                                     # 입출고 파일
                offices.append(office_block(f.get("name") or "실", per))
                for d, (inb, outb) in per.items():
                    allbydate[d][0].extend(inb); allbydate[d][1].extend(outb)
                continue
            ws = inventory_sheet(wb)                    # 재고 파일
            if ws is not None:
                inv = parse_inventory(ws)
                if inv:
                    inventory = inv
        finally:
            wb.close()

    if not offices and not inventory:
        raise ValueError("날짜 시트(YYYY-MM-DD)가 있는 입출고 파일이나 "
                         "'inventory' 시트가 있는 재고 파일을 찾지 못했습니다.")
    if len(offices) > 1:                            # 전체 합계 탭 (맨 앞)
        agg = {d: (allbydate[d][0], allbydate[d][1]) for d in allbydate}
        offices.insert(0, office_block("전체 합계", agg))
    return dict(offices=offices, inventory=inventory)


# ---------------------------------------------------------------- 이메일 (Outlook)
def _e(s):
    return (str(s) if s is not None else "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


CONTACT_NAME = "유니트론텍 안성우 책임"
CONTACT_MAIL = "sw.ahn@unitrontech.com"


def compose_email_html(office, date, day=None, link=""):
    """메일 본문 — 대시보드 링크와 문의처만. (요약·내역표는 대시보드에서 본다)"""
    btn = (f'<a href="{_e(link)}" style="display:inline-block;background:#c43a3a;color:#fff;'
           f'text-decoration:none;font-weight:800;font-size:15px;padding:14px 28px;'
           f'border-radius:10px">입출고 및 재고현황 조회 →</a>'
           f'<div style="font-size:12px;color:#999;margin-top:12px;word-break:break-all">{_e(link)}</div>'
           ) if link else '<div style="color:#999">링크 없음</div>'

    return f"""<div style="font-family:'Malgun Gothic',sans-serif;color:#222;max-width:560px">
  {btn}
  <div style="margin-top:30px;padding-top:14px;border-top:1px solid #e6e6ea;font-size:13px;color:#555">
    문의: <b>{_e(CONTACT_NAME)}</b>
    (<a href="mailto:{_e(CONTACT_MAIL)}" style="color:#3a6ea5;text-decoration:none">{_e(CONTACT_MAIL)}</a>)
  </div>
</div>"""


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


def _send_smtp(emails, subject, body, cfg):
    """SMTP 발송. cfg = {host, port, user, pw, sender}."""
    import smtplib, ssl
    from email.mime.text import MIMEText
    host, port = cfg["host"], cfg["port"]
    user, pw, sender = cfg["user"], cfg["pw"], cfg["sender"]
    msg = MIMEText(body, "html", "utf-8")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(emails)
    # 답장은 발송용 계정이 아니라 문의처로 가야 한다
    msg["Reply-To"] = CONTACT_MAIL
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
    subject = f"[입출고 대시보드] {office} · {date}"
    body = compose_email_html(office, date, day, link)
    cfg = resolve_smtp(payload.get("smtp"))
    if cfg:                                             # SMTP (요청 설정 또는 환경변수)
        if not emails:
            raise ValueError("받는 사람 이메일을 입력하세요.")
        return _send_smtp(emails, subject, body, cfg)
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
    <div class="viewseg">
      <button class="on" id="vw-io" onclick="showView('io')">📦 일일 입출고</button>
      <button id="vw-inv" onclick="showView('inv')">📊 재고 현황</button>
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

    <div class="card" id="smscard" style="margin-top:22px">
      <h2>📱 대시보드 링크 문자 발송 <span style="font-size:12px;color:var(--mut);font-weight:500">— 각 직원에게 자동으로 링크 전송</span></h2>
      <p class="desc">받는 사람 휴대폰 번호를 넣고 발송하면, <b>지금 보고 있는 실의 링크</b>가 문자로 각 사람에게 자동 전송됩니다. (실마다 링크·번호가 따로 — 줄바꿈·콤마·세미콜론 구분 · 저장됨)</p>
      <div style="font-size:12.5px;color:#555;margin:0 0 12px;word-break:break-all;background:#f6f6f8;border-radius:9px;padding:9px 13px">
        🔗 <b id="off-name"></b> 링크: <span id="officelink" style="color:var(--blue);font-weight:600"></span>
        <button id="btn-copy" style="margin-left:6px;font-size:11.5px;border:1px solid var(--line);background:#fff;border-radius:7px;padding:3px 9px;cursor:pointer;font-family:inherit">복사</button>
      </div>
      <textarea id="phonebox" rows="3" placeholder="010-1234-5678, 010-2222-3333, ..." style="width:100%;font-size:14px;padding:12px 14px;border:1.5px solid var(--line);border-radius:11px;font-family:inherit;outline:none;resize:vertical"></textarea>
      <div style="margin-top:14px;display:flex;gap:12px;align-items:center;flex-wrap:wrap">
        <button id="btn-sms" class="confirmbtn">📱 링크 문자 발송</button>
        <span id="smsstatus" style="font-size:12.5px;color:var(--mut)"></span>
      </div>
    </div>

    <div class="card" id="emailcard" style="margin-top:22px">
      <h2>📧 대시보드 링크 이메일 발송 <span style="font-size:12px;color:var(--mut);font-weight:500">— 각 실 담당자에게 그 실 링크 전송</span>
        <button id="btn-smtp" style="float:right;font-size:12px;border:1px solid var(--line);background:#fff;border-radius:8px;padding:6px 12px;cursor:pointer;font-family:inherit;font-weight:600">⚙️ 메일 설정</button></h2>
      <p class="desc">받는 담당자 이메일을 넣고 발송하면 <b>지금 보고 있는 실의 링크</b>가 이메일로 전송됩니다. (실마다 따로 저장) · <b>전체 실 일괄 발송</b>은 모든 실에 각자 링크를 한 번에 보냅니다. · 보내는 계정은 <b>서버에 미리 설정</b>돼 있어 별도 설정 없이 발송됩니다. (⚙️는 로컬 테스트용)</p>
      <div style="font-size:12.5px;color:#555;margin:0 0 12px;word-break:break-all;background:#f6f6f8;border-radius:9px;padding:9px 13px">
        🔗 <b id="off-name2"></b> 링크: <span id="officelink2" style="color:var(--blue);font-weight:600"></span></div>
      <textarea id="emailbox" rows="3" placeholder="seanlee@unitrontech.com" style="width:100%;font-size:14px;padding:12px 14px;border:1.5px solid var(--line);border-radius:11px;font-family:inherit;outline:none;resize:vertical"></textarea>
      <div style="margin-top:14px;display:flex;gap:12px;align-items:center;flex-wrap:wrap">
        <button id="btn-email" class="confirmbtn">📧 이 실 링크 이메일 발송</button>
        <button id="btn-email-all" class="confirmbtn" style="background:var(--ink)">📮 전체 실 일괄 발송</button>
        <span id="emailstatus" style="font-size:12.5px;color:var(--mut)"></span>
      </div>
    </div>

    <div class="foot" id="foot"></div>
  </div>

  <!-- ================= 재고 현황 ================= -->
  <div class="wrap" id="invview" style="display:none">
    <div class="kpis" id="inv-kpis"></div>
    <div class="hl" id="inv-hl" style="display:none"></div>
    <div class="grid">
      <div class="card"><h2>재고 노후화 (Datecode 연도별)</h2>
        <p class="desc"><span id="inv-oldlabel">—</span>년 이전 = 장기재고 (빨강) · 연도별 편차가 커서 <b>로그 스케일</b> — 막대 길이를 그대로 비교하지 마세요</p>
        <div class="cbox"><canvas id="cAge"></canvas></div></div>
      <div class="card"><h2>FAMILY별 재고</h2><p class="desc">수량 기준 상위</p>
        <div class="cbox"><canvas id="cFam"></canvas></div></div>
    </div>
    <div class="grid">
      <div class="card"><h2>당월 일별 입출고</h2><p class="desc">재고 시트 기준</p>
        <div class="cbox"><canvas id="cMon"></canvas></div></div>
      <div class="card"><h2>고객사 예약(booking)</h2><p class="desc">예약 수량 상위</p>
        <div class="cbox"><canvas id="cBook"></canvas></div></div>
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

<div id="smtpmodal" style="display:none;position:fixed;inset:0;background:rgba(20,20,26,.55);z-index:200;align-items:center;justify-content:center;padding:20px">
  <div style="background:#fff;border-radius:16px;max-width:440px;width:100%;padding:24px 26px;box-shadow:0 20px 60px rgba(0,0,0,.4)">
    <h2 style="margin:0 0 4px;font-size:17px;font-weight:800">⚙️ 메일 발송 설정 (SMTP)</h2>
    <p style="font-size:12px;color:var(--mut);margin:0 0 16px;line-height:1.6">이메일 보내는 계정을 1회 등록합니다. (이 브라우저에만 저장) · Gmail: <b>smtp.gmail.com</b>/587 + <b>앱 비밀번호</b> · Office365: <b>smtp.office365.com</b>/587</p>
    <div style="display:flex;flex-direction:column;gap:10px">
      <label style="font-size:12px;color:#555;font-weight:600">SMTP 호스트<input id="s-host" type="text" placeholder="smtp.gmail.com" style="width:100%;margin-top:4px;padding:9px 11px;border:1.5px solid var(--line);border-radius:9px;font-family:inherit;font-size:13px"></label>
      <div style="display:flex;gap:10px">
        <label style="font-size:12px;color:#555;font-weight:600;width:90px">포트<input id="s-port" type="text" value="587" style="width:100%;margin-top:4px;padding:9px 11px;border:1.5px solid var(--line);border-radius:9px;font-family:inherit;font-size:13px"></label>
        <label style="font-size:12px;color:#555;font-weight:600;flex:1">보내는 주소(From)<input id="s-from" type="text" placeholder="비우면 계정과 동일" style="width:100%;margin-top:4px;padding:9px 11px;border:1.5px solid var(--line);border-radius:9px;font-family:inherit;font-size:13px"></label>
      </div>
      <label style="font-size:12px;color:#555;font-weight:600">계정(아이디)<input id="s-user" type="text" placeholder="you@company.com" style="width:100%;margin-top:4px;padding:9px 11px;border:1.5px solid var(--line);border-radius:9px;font-family:inherit;font-size:13px"></label>
      <label style="font-size:12px;color:#555;font-weight:600">비밀번호(앱 비밀번호)<input id="s-pass" type="password" style="width:100%;margin-top:4px;padding:9px 11px;border:1.5px solid var(--line);border-radius:9px;font-family:inherit;font-size:13px"></label>
    </div>
    <div style="margin-top:18px;display:flex;gap:10px;justify-content:flex-end">
      <button id="smtp-cancel" style="background:#eee;border:none;font-family:inherit;font-size:13px;font-weight:700;padding:10px 18px;border-radius:9px;cursor:pointer;color:#555">닫기</button>
      <button id="smtp-save" class="confirmbtn" style="padding:10px 20px">저장</button>
    </div>
  </div>
</div>

<script>
const fmt=n=>(n==null?'—':Number(n).toLocaleString('ko-KR'));
const esc=s=>String(s==null?'':s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const RED="#c43a3a",BLUE="#3a6ea5",GREEN="#3f9d6b",AMBER="#e0a93a",INK="#23232b";
if(window.Chart){Chart.defaults.font.family="'Pretendard',system-ui,sans-serif";Chart.defaults.color="#6b6b74";Chart.defaults.font.size=12;}
let DATA=null, O=null, OFFI=0, SERVER_MAIL=false;

// ── 고정 수신자 ──────────────────────────────────────────────
// 실 이름에 숫자가 들어가면 여기서 잡는다 (예: '영업4실', '영업1,2실').
// 여기 없는 실은 DEFAULT_EMAILS 로 간다.
const OFFICE_EMAILS={
  '4': 'sean94kr@gmail.com',            // 영업4실
  '5': 'seanlee@unitrontech.com',       // 영업5실
};
const DEFAULT_EMAILS='seanlee@unitrontech.com';
function emailsFor(name){
  return parseEmails(OFFICE_EMAILS[officeSlug(name)]||DEFAULT_EMAILS);
}
// 서버에 보내는 메일계정(SMTP)이 설정돼 있으면, 사용자는 아무 설정 없이 발송 가능
fetch('/mailcfg').then(r=>r.json()).then(d=>{ SERVER_MAIL=!!d.configured;
  const b=document.getElementById('btn-smtp');
  if(SERVER_MAIL && b){ b.textContent='✅ 메일 발송 준비됨'; b.title='서버에 발송 계정이 설정돼 있습니다'; }
}).catch(()=>{});

// 업로드 (여러 파일)
const drop=document.getElementById('drop'),file=document.getElementById('file');
drop.onclick=()=>file.click();
['dragover','dragenter'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('hot');}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('hot');}));
drop.addEventListener('drop',ev=>{if(ev.dataTransfer.files.length)go(ev.dataTransfer.files);});
file.addEventListener('change',ev=>{if(ev.target.files.length)go(ev.target.files);});

function officeName(fn){ const m=fn.match(/\(([^)]+)\)/); return m?m[1]:fn.replace(/\.xlsx$/i,''); }
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

  const hasIO=DATA.offices && DATA.offices.length, hasInv=!!DATA.inventory;
  // 두 종류가 다 올라왔을 때만 전환 버튼을 보여준다
  document.getElementById('viewnav').style.display=(hasIO&&hasInv)?'block':'none';

  if(hasIO){
    const slug=location.pathname.replace(/\//g,'');
    let idx=0;
    if(slug){ const i=DATA.offices.findIndex(o=>officeSlug(o.name)===slug); if(i>=0) idx=i; }
    selectOffice(idx);
  }
  if(hasInv) renderInventory();
  showView(hasIO?'io':'inv');
}

function showView(v){
  const io=v==='io';
  document.getElementById('ioview').style.display=io?'block':'none';
  document.getElementById('invview').style.display=io?'none':'block';
  document.getElementById('vw-io').classList.toggle('on',io);
  document.getElementById('vw-inv').classList.toggle('on',!io);
  if(!io){
    const I=DATA.inventory;
    document.getElementById('h-date').innerHTML=`재고 현황 <span class="d">·</span> ${esc(I.sheet)}`;
    document.getElementById('h-meta').textContent=`품목 ${fmt(I.n_items)}건 · 총 재고 ${fmt(I.total_qty)} EA`;
  }else if(O){ showDay(CUR); }
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
  buildEmailCard();
}

function officeSlug(name){ const d=(String(name).match(/\d/g)||[]).join(''); return d||'all'; }
function officeLink(){ return location.origin+'/'+officeSlug(O.name); }
function buildEmailCard(){
  document.getElementById('off-name').textContent=O.name;
  document.getElementById('officelink').textContent=officeLink();
  document.getElementById('phonebox').value = localStorage.getItem('phones_'+O.name)||'';
  document.getElementById('smsstatus').textContent='';
  document.getElementById('off-name2').textContent=O.name;
  document.getElementById('officelink2').textContent=officeLink();
  document.getElementById('emailbox').value =
    localStorage.getItem('emails_'+O.name)||emailsFor(O.name).join(', ');
  document.getElementById('emailstatus').textContent='';
}
document.getElementById('btn-copy').onclick=()=>{
  navigator.clipboard.writeText(officeLink()).then(()=>{
    document.getElementById('smsstatus').textContent='✅ 링크 복사됨 — 카톡 등에 붙여넣기 가능';
  });
};
document.getElementById('btn-sms').onclick=()=>{
  const st=document.getElementById('smsstatus');
  const raw=document.getElementById('phonebox').value.trim();
  localStorage.setItem('phones_'+O.name, raw);
  const numbers=[...new Set(raw.split(/[;,\s]+/).map(s=>s.trim()).filter(Boolean))];
  if(!numbers.length){ st.textContent='⚠️ 받는 휴대폰 번호를 입력하세요.'; return; }
  const day=O.days[CUR];
  const text=`[입출고 대시보드] ${O.name} (${day.date})\n${officeLink()}`;
  st.textContent=`문자 발송 중… (${numbers.length}명)`;
  fetch('/sms',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({numbers, text})})
  .then(r=>r.json()).then(res=>{
    st.textContent = res.error ? ('오류: '+res.error)
      : `✅ ${O.name} 링크를 ${numbers.length}명에게 문자 발송했습니다.`;
  }).catch(e=>{ st.textContent='전송 오류: '+e; });
};

// ── SMTP 설정 (이 브라우저에 저장) ──
function getSmtp(){ try{return JSON.parse(localStorage.getItem('smtp_config')||'{}');}catch(e){return {};} }
const smtpModal=document.getElementById('smtpmodal');
document.getElementById('btn-smtp').onclick=()=>{
  const c=getSmtp();
  document.getElementById('s-host').value=c.host||'';
  document.getElementById('s-port').value=c.port||'587';
  document.getElementById('s-from').value=c.from||'';
  document.getElementById('s-user').value=c.user||'';
  document.getElementById('s-pass').value=c.pass||'';
  smtpModal.style.display='flex';
};
document.getElementById('smtp-cancel').onclick=()=>smtpModal.style.display='none';
smtpModal.onclick=e=>{ if(e.target===smtpModal) smtpModal.style.display='none'; };
document.getElementById('smtp-save').onclick=()=>{
  const c={host:document.getElementById('s-host').value.trim(),
    port:document.getElementById('s-port').value.trim()||'587',
    from:document.getElementById('s-from').value.trim(),
    user:document.getElementById('s-user').value.trim(),
    pass:document.getElementById('s-pass').value};
  if(!c.host){ alert('SMTP 호스트를 입력하세요.'); return; }
  localStorage.setItem('smtp_config', JSON.stringify(c));
  smtpModal.style.display='none';
  document.getElementById('emailstatus').textContent='✅ 메일 설정 저장됨 — 이제 발송할 수 있어요.';
};

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
async function autoSend(){
  const st=document.getElementById('emailstatus');
  const targets=(DATA.offices||[])
    .filter(o=>o.name!=='전체 합계')
    .map(o=>({o, emails:emailsFor(o.name)}))
    .filter(t=>t.emails.length);
  if(!targets.length) return;

  if(!SERVER_MAIL && !getSmtp().host){
    if(st) st.textContent='⚠️ 자동 발송 안 됨 — 서버에 메일 계정(SMTP)이 설정돼 있지 않습니다.';
    return;
  }
  if(st) st.textContent=`업로드 완료 — 자동 발송 중… (${targets.length}개 실)`;

  const ok=[], fail=[];
  for(const t of targets){
    try{
      const r=await sendOfficeEmail(t.o, t.emails);
      (r && r.ok && r.sent ? ok : fail).push(t.o.name+(r&&r.error?(' ('+r.error+')'):''));
    }catch(e){ fail.push(t.o.name+' ('+e+')'); }
  }
  if(st) st.textContent=
    (ok.length?`✅ 자동 발송 완료: ${ok.join(', ')}`:'')+
    (fail.length?`${ok.length?' · ':''}❌ 실패: ${fail.join(', ')}`:'');
}

document.getElementById('btn-email').onclick=()=>{
  const st=document.getElementById('emailstatus');
  if(!SERVER_MAIL && !getSmtp().host){ st.textContent='⚠️ 메일 발송 계정이 없습니다. (배포 시엔 서버에 설정되어 자동, 로컬 테스트는 ⚙️ 메일 설정)'; return; }
  const raw=document.getElementById('emailbox').value.trim();
  localStorage.setItem('emails_'+O.name, raw);
  const emails=parseEmails(raw);
  if(!emails.length){ st.textContent='⚠️ 받는 이메일을 입력하세요.'; return; }
  st.textContent=`발송 중… (${emails.length}명)`;
  sendOfficeEmail(O, emails).then(res=>{
    st.textContent = res.error ? ('오류: '+res.error) : `✅ ${O.name} 링크를 ${emails.length}명에게 이메일 발송했습니다.`;
  }).catch(e=>{ st.textContent='전송 오류: '+e; });
};
document.getElementById('btn-email-all').onclick=async()=>{
  const st=document.getElementById('emailstatus');
  if(!SERVER_MAIL && !getSmtp().host){ st.textContent='⚠️ 메일 발송 계정이 없습니다. (배포 시엔 서버에 설정되어 자동)'; return; }
  localStorage.setItem('emails_'+O.name, document.getElementById('emailbox').value.trim());
  const targets=DATA.offices.map(o=>({o, emails:parseEmails(localStorage.getItem('emails_'+o.name)||'')})).filter(t=>t.emails.length);
  if(!targets.length){ st.textContent='⚠️ 저장된 실별 이메일이 없습니다. 각 실(/12 /3 /4 /5)에서 이메일을 입력·저장하세요.'; return; }
  st.textContent=`전체 발송 중… (${targets.length}개 실)`;
  let ok=0; const fail=[];
  for(const t of targets){
    try{ const res=await sendOfficeEmail(t.o, t.emails); if(res.error) fail.push(t.o.name+': '+res.error); else ok++; }
    catch(e){ fail.push(t.o.name+': '+e); }
  }
  st.textContent=`✅ ${ok}개 실 발송 완료`+(fail.length?` · ⚠️ 실패 ${fail.length}건: `+fail.join(' / '):'');
};

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

  const net=k.net, netcol=net>=0?GREEN:RED;
  document.getElementById('summary').innerHTML=`
    <div style="display:flex;gap:18px;flex-wrap:wrap;font-size:13px;line-height:1.9">
      <div>📦 <b>입고</b> ${fmt(k.in_cnt)}건 · ${fmt(k.in_qty)} EA</div>
      <div>🚚 <b>출고</b> ${fmt(k.out_cnt)}건 · ${fmt(k.out_qty)} EA</div>
      <div>⚖️ <b>순물동</b> <span style="color:${netcol};font-weight:800">${net>=0?'+':''}${fmt(net)} EA</span></div>
      <div>🏢 <b>출고 거래처</b> ${fmt(k.customers)}곳</div>
    </div>`;

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

function renderInventory(){
  const I=DATA.inventory;
  document.getElementById('inv-oldlabel').textContent=I.old_year;
  document.getElementById('inv-foot').textContent=
    `자료: 재고 엑셀의 '${I.sheet}' 시트 · 재고 수량이 있는 품목만 집계 (합계 행 제외) · 장기재고 = Datecode ${I.old_year}년 이전`;

  document.getElementById('inv-kpis').innerHTML=[
    ['it','재고 품목',I.n_items,'건'],
    ['in','총 재고',I.total_qty,'EA'],
    ['av','가용 재고',I.avail_qty,'EA'],
    ['bk','예약(booking)',I.booking_qty,'EA'],
    ['old',`장기재고 (${I.old_year}년 이전)`,I.old_qty,'EA'],
    ['cu','당월 출고',I.month.outbound,'EA'],
  ].map(([c,l,v,u])=>`<div class="kpi ${c}"><div class="l">${l}</div>
     <div class="v tab">${fmt(v)}<span class="u">${u}</span></div></div>`).join('');

  const hl=document.getElementById('inv-hl');
  if(I.old_qty>0){
    hl.style.display='flex';
    hl.innerHTML=`<span class="tag">장기재고</span><div class="txt">
      Datecode <b>${I.old_year}년 이전</b> 재고가 <span class="q">${fmt(I.old_qty)} EA</span>
      — 전체 재고의 <b>${pct(I.old_qty,I.total_qty)}%</b></div>`;
  }else hl.style.display='none';

  Object.values(ICH).forEach(c=>c&&c.destroy()); ICH={};

  // 재고 노후화 — 장기재고는 빨강
  // 2026년 재고가 2019년의 50배라 선형 축에서는 장기재고 막대가 안 보인다 → 로그 스케일
  const dc=I.datecode.filter(d=>d.qty>0);
  ICH.age=new Chart(document.getElementById('cAge'),{type:'bar',
    data:{labels:dc.map(d=>d.year),datasets:[{data:dc.map(d=>d.qty),
      backgroundColor:dc.map(d=>d.year<=I.old_year?RED:BLUE),borderRadius:6,maxBarThickness:52}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{
        label:c=>fmt(c.raw)+' EA ('+dc[c.dataIndex].items+'품목) · 전체의 '
                 +pct(c.raw,I.total_qty)+'%'}}},
      scales:{y:{type:'logarithmic',grid:{color:'#f0f0f3'},
                 ticks:{callback:v=>{const s=String(v);
                   return /^[125]0*$/.test(s)?fmt(v):'';}}},
              x:{grid:{display:false}}}}});

  const fam=I.by_family.slice(0,8);
  ICH.fam=new Chart(document.getElementById('cFam'),{type:'bar',
    data:{labels:fam.map(f=>f.name||'(미지정)'),datasets:[{data:fam.map(f=>f.qty),
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

  document.getElementById('in-all').textContent=I.items.length;
  document.getElementById('in-old').textContent=I.items.filter(x=>x.old>0).length;
  document.getElementById('in-bk').textContent=I.items.filter(x=>x.booking>0).length;
  document.getElementById('invq').oninput=()=>drawInvTable();
  showInvTab('all');
}

function showInvTab(t){
  ITAB=t;
  ['all','old','bk'].forEach(k=>document.getElementById('ib-'+k).classList.toggle('on',k===t));
  drawInvTable();
}

function drawInvTable(){
  const I=DATA.inventory;
  const q=(document.getElementById('invq').value||'').trim().toUpperCase();
  let rows=I.items;
  if(ITAB==='old') rows=rows.filter(x=>x.old>0);
  if(ITAB==='bk')  rows=rows.filter(x=>x.booking>0);
  if(q) rows=rows.filter(x=>[x.part,x.mobis,x.family,x.sales,x.customer,x.pn]
      .some(v=>String(v||'').toUpperCase().includes(q)));
  rows=[...rows].sort((a,b)=>b.qty-a.qty);

  const el=document.getElementById('invtable');
  if(!rows.length){ el.innerHTML=emptyMsg(); return; }
  el.innerHTML=`<table><thead><tr><th>#</th><th>PART#</th><th>MOBIS ID</th><th>FAMILY</th>
    <th>실</th><th class="n">재고</th><th class="n">가용</th><th class="n">예약</th>
    <th class="n">장기재고</th><th>Datecode</th><th>담당</th></tr></thead><tbody>${
    rows.map((x,i)=>`<tr class="${x.old>0?'oldrow':''}">
      <td class="n">${i+1}</td>
      <td class="part">${esc(x.part)}</td>
      <td>${esc(x.mobis)||'—'}</td>
      <td>${esc(x.family)||'—'}</td>
      <td>${esc(x.office)||'—'}</td>
      <td class="qty">${fmt(x.qty)}</td>
      <td class="n">${fmt(x.avail)}</td>
      <td class="n">${x.booking?fmt(x.booking):'—'}</td>
      <td class="n ${x.old>0?'old':''}">${x.old?fmt(x.old):'—'}</td>
      <td>${x.oldest?('<span class="pill">'+x.oldest+'~</span>'):'—'}</td>
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
