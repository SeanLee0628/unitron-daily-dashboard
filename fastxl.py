# -*- coding: utf-8 -*-
"""값만 읽는 최소 xlsx 리더.

openpyxl 은 워크북을 열 때 styles.xml 을 전부 파이썬 객체로 만든다.
이 파일들은 styles.xml 이 14MB(셀 서식 수십만 개)라 그 과정만 파일당 5~8초가 걸린다.
우리는 값만 필요하므로 서식 객체를 만들 이유가 없다.

여기서는 zip 에서 필요한 시트 XML 만 스트리밍으로 읽는다.
날짜는 numFmt 만 훑어 판별한다 (XML 파싱 자체는 0.2초로 싸다).

실측: 파일 8개 열기 47.8초 → 2.5초.
"""
from __future__ import annotations

import datetime
import io
import re
import zipfile
import xml.etree.ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

_COL = re.compile(r"([A-Z]+)")
_EPOCH = datetime.datetime(1899, 12, 30)          # 엑셀 1900 날짜 체계

# 엑셀 내장 날짜/시간 서식 번호
_BUILTIN_DATE = set(range(14, 23)) | set(range(45, 48)) | {27, 30, 36, 50, 57}
_DATE_CHARS = re.compile(r"[ymdhs]")
_ESCAPED = re.compile(r'\[[^\]]*\]|"[^"]*"|\\.')  # 서식의 리터럴 구간


def _col_idx(ref):
    m = _COL.match(ref or "")
    if not m:
        return 0
    n = 0
    for ch in m.group(1):
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _is_date_fmt(code):
    """서식 문자열이 날짜/시간인가. 리터럴('원', [red] 등)은 빼고 본다."""
    return bool(_DATE_CHARS.search(_ESCAPED.sub("", (code or "").lower())))


class FastWorkbook:
    """openpyxl.Workbook 대체 (읽기 전용, 값만)."""

    def __init__(self, fileobj):
        self.z = zipfile.ZipFile(fileobj)
        self._strings = None
        self._date_xf = None

        wb = ET.fromstring(self.z.read("xl/workbook.xml"))
        rels = ET.fromstring(self.z.read("xl/_rels/workbook.xml.rels"))
        tgt = {r.get("Id"): r.get("Target") for r in rels}
        self._paths = {}
        self.sheetnames = []
        for sh in wb.iter(NS + "sheet"):
            t = tgt.get(sh.get(RNS + "id"), "")
            if t.startswith("/"):
                t = t[1:]
            elif not t.startswith("xl/"):
                t = "xl/" + t
            name = sh.get("name")
            self.sheetnames.append(name)
            self._paths[name] = t

    # ---- 공유 문자열 ----
    @property
    def strings(self):
        if self._strings is None:
            out = []
            try:
                data = self.z.read("xl/sharedStrings.xml")
            except KeyError:
                self._strings = out
                return out
            for _, si in ET.iterparse(io.BytesIO(data), events=("end",)):
                if si.tag == NS + "si":
                    out.append("".join(t.text or "" for t in si.iter(NS + "t")))
                    si.clear()
            self._strings = out
        return self._strings

    # ---- 날짜 서식인 스타일 인덱스 ----
    @property
    def date_xf(self):
        if self._date_xf is None:
            xf = set()
            try:
                root = ET.fromstring(self.z.read("xl/styles.xml"))
            except KeyError:
                self._date_xf = xf
                return xf
            custom = {}
            for nf in root.iter(NS + "numFmt"):
                try:
                    custom[int(nf.get("numFmtId"))] = nf.get("formatCode") or ""
                except (TypeError, ValueError):
                    pass
            cellxfs = root.find(NS + "cellXfs")
            if cellxfs is not None:
                for i, x in enumerate(cellxfs.findall(NS + "xf")):
                    try:
                        fid = int(x.get("numFmtId") or 0)
                    except ValueError:
                        continue
                    if fid in _BUILTIN_DATE or _is_date_fmt(custom.get(fid)):
                        xf.add(i)
            self._date_xf = xf
        return self._date_xf

    def __getitem__(self, name):
        return FastSheet(self, name)

    def close(self):
        self.z.close()


class FastSheet:
    def __init__(self, wb, name):
        self.wb = wb
        self.name = name
        self.title = name

    def iter_rows(self, min_row=1, max_row=None, values_only=True):
        """openpyxl 과 같은 모양. 빈 행도 건너뛰지 않고 채워서 낸다
        (헤더를 행 위치로 찾는 코드가 있어 행이 밀리면 안 된다)."""
        path = self.wb._paths.get(self.name)
        if not path:
            return
        ss, dxf = self.wb.strings, self.wb.date_xf
        expect = 1                                  # 다음에 나와야 할 행 번호
        with self.wb.z.open(path) as fp:
            for _, el in ET.iterparse(fp, events=("end",)):
                if el.tag != NS + "row":
                    continue
                try:
                    rn = int(el.get("r") or expect)
                except ValueError:
                    rn = expect
                while expect < rn:                  # 생략된 빈 행 채우기
                    if max_row and expect > max_row:
                        el.clear(); return
                    if expect >= min_row:
                        yield ()
                    expect += 1

                row = []
                for c in el.iterfind(NS + "c"):
                    i = _col_idx(c.get("r"))
                    t = c.get("t")
                    if t == "inlineStr":
                        isel = c.find(NS + "is")
                        val = ("".join(x.text or "" for x in isel.iter(NS + "t"))
                               if isel is not None else None)
                    else:
                        v = c.find(NS + "v")
                        if v is None or v.text is None:
                            val = None
                        elif t == "s":
                            k = int(v.text)
                            val = ss[k] if 0 <= k < len(ss) else None
                        elif t == "b":
                            val = v.text == "1"
                        elif t in ("str", "e"):
                            val = v.text
                        else:
                            s = v.text
                            try:
                                val = (int(s) if ("." not in s and "e" not in s.lower())
                                       else float(s))
                            except ValueError:
                                val = s
                            # 날짜 서식이면 datetime 으로 (openpyxl 과 동일하게)
                            if isinstance(val, (int, float)):
                                try:
                                    if int(c.get("s") or -1) in dxf and val > 0:
                                        val = _EPOCH + datetime.timedelta(days=float(val))
                                except (ValueError, OverflowError):
                                    pass
                    while len(row) <= i:
                        row.append(None)
                    row[i] = val
                el.clear()

                if max_row and rn > max_row:
                    return
                if rn >= min_row:
                    yield tuple(row)
                expect = rn + 1


def load(fileobj):
    return FastWorkbook(fileobj)
