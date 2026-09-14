#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""헤더 정규화(_hkey) 회귀 테스트.

원래 버그: app.py 에 _hkey 가 두 번 정의돼 있었다(117행, 928행). 파이썬은 나중
정의가 이기므로 아포스트로피·슬래시를 안 떼는 쪽이 모듈 전체를 덮어썼고,
엑셀 헤더 "Q'ty" 가 "QTY" 로 안 풀려 입출고 수량이 전부 0 으로 떴다.
건수·거래처는 멀쩡해서 화면상으로는 "수량만 0" 으로 보였다.
"""
import re
import unittest

import app


class TestHkey(unittest.TestCase):
    def test_strips_apostrophe_and_slash(self):
        # 현장 파일 헤더가 Q'ty 다 — 이게 QTY 로 안 풀리면 수량이 통째로 0 이 된다
        self.assertEqual(app._hkey("Q'ty"), "QTY")
        self.assertEqual(app._hkey("available Q'ty"), "AVAILABLEQTY")
        self.assertEqual(app._hkey(" in / out "), "INOUT")
        self.assertEqual(app._hkey("PART #"), "PART#")

    def test_defined_only_once(self):
        """중복 정의가 다시 생기면 조용히 같은 버그가 재발한다."""
        with open("app.py", encoding="utf-8") as f:
            n = len(re.findall(r"^def _hkey\(", f.read(), re.M))
        self.assertEqual(n, 1, f"_hkey 정의가 {n} 개다 — 나중 정의가 앞을 덮어쓴다")


class FakeSheet:
    """parse_daily 가 쓰는 것만 흉내: iter_rows(values_only=True) + comments."""

    comments = {}

    def __init__(self, rows):
        self._rows = rows

    def iter_rows(self, values_only=True, **kw):
        return iter(self._rows)


class TestParseDailyQty(unittest.TestCase):
    ROWS = [
        ("입고", None, None, None),
        ("NO", "CUSTOMER", "PART#", "Q'ty"),
        ("1", "GEMSTONE Korea", "MTSD064AMC8MS-2WT", 1500),
        ("2", "YURA", "MAYA-W166-00B", "2,000"),
        ("출고", None, None, None),
        ("NO", "CUSTOMER", "PART#", "Q'ty"),
        ("1", "METACOM", "MT28EW01GABA1HPC-0SIT", 340),
    ]

    def test_qty_column_is_read(self):
        inbound, outbound = app.parse_daily(FakeSheet(self.ROWS))
        self.assertEqual([r["qty"] for r in inbound], [1500.0, 2000.0])
        self.assertEqual([r["qty"] for r in outbound], [340.0])
        self.assertEqual(outbound[0]["customer"], "METACOM")


if __name__ == "__main__":
    unittest.main(verbosity=2)
