#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""거래처코드 마스터(품목코드 생성기) 회귀 테스트.

출고현황의 CUSTOMER# 가 비어 있으면 마스터의 거래처명/약칭으로 코드를 찾는다.
시트에 원래 있던 코드는 건드리지 않고, 한 이름에 코드가 둘 이상이면 비워 둔다.
"""
import os
import tempfile
import unittest

import app
from test_hkey import FakeSheet


class TestKey(unittest.TestCase):
    def test_corp_forms_match(self):
        k = app.cmaster_key("주식회사 선민테크")
        for s in ("(주)선민테크", "㈜ 선민테크", "선민테크(주)", "선민테크 주식회사", " 선민 테크 "):
            self.assertEqual(app.cmaster_key(s), k, s)

    def test_english_suffix_only_at_end(self):
        self.assertEqual(app.cmaster_key("Abc Co., Ltd."), app.cmaster_key("ABC"))
        self.assertEqual(app.cmaster_key("Lincoln"), "LINCOLN")      # 중간 INC 는 그대로


class TestMaster(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.old = app.LOG_DB
        app.LOG_DB = os.path.join(self.tmp, "log.db")
        ws = FakeSheet([
            ("코드", "거래처명", "거래처약칭", "사용여부"),
            ("138692", "주식회사 선민테크", "주식회사 선민테크", "사용"),
            ("100001", "(주)중복", "", "사용"),
            ("100002", "중복 주식회사", "", "사용"),
            ("100003", "(주)옛코드", "", "미사용"),
            ("100004", "(주)옛코드", "", "사용"),
            (None, None, None, None),
        ])
        app.cmaster_store(app.parse_cmaster(ws))

    def tearDown(self):
        app.LOG_DB = self.old

    def test_map(self):
        cm = app.cmaster_map()
        self.assertEqual(cm[app.cmaster_key("선민테크")], "138692")
        self.assertNotIn(app.cmaster_key("중복"), cm)                  # 둘 다 사용 → 고르지 않음
        self.assertEqual(cm[app.cmaster_key("옛코드")], "100004")      # 사용 쪽만

    def test_apply_keeps_sheet_code_and_refreshes(self):
        row = lambda c, cc: dict(date="2026-10-23", customer=c, part="P", qty=1,  # noqa: E731
                                 sales="", ccode=cc)
        app.log_store("영업5실", {"out": [row("주식회사 선민테크", ""),
                                          row("(주)옛코드", "999999"),
                                          row("모르는회사", "")]})
        self.assertEqual(app.cmaster_apply(), 1)
        got = lambda: dict(app.log_conn().execute(                     # noqa: E731
            "SELECT customer, ccode FROM log").fetchall())
        self.assertEqual(got()["주식회사 선민테크"], "138692")
        self.assertEqual(got()["(주)옛코드"], "999999")                # 시트 값이 이긴다
        self.assertEqual(got()["모르는회사"], "")
        # 마스터가 바뀌면 마스터로 채운 값만 다시 계산된다
        app.cmaster_store([("200000", "주식회사 선민테크", "", "사용")])
        app.cmaster_apply()
        self.assertEqual(got()["주식회사 선민테크"], "200000")
        self.assertEqual(got()["(주)옛코드"], "999999")

    def test_fill_rows(self):
        rows = [dict(customer="선민테크", ccode=""), dict(customer="선민테크", ccode="1")]
        app.cmaster_fill(rows, app.cmaster_map())
        self.assertEqual([r["ccode"] for r in rows], ["138692", "1"])


if __name__ == "__main__":
    unittest.main()
