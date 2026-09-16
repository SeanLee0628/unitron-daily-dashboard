#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""VENDER(제조사) 칸 회귀 테스트.

출고 시트 K열에 KEC/Delta/Micron/ublox 같은 제조사가 적혀 있는데 헤더 행에는
칸 이름이 없다 (날짜 시트·누적 '출고' 시트 모두). 그래서 이름으로 못 찾으면
표준 배치의 K열(10번)을 쓰는데, 다른 칸이 제자리에 있을 때만 그렇게 한다 —
배치가 바뀐 파일에서 엉뚱한 값을 VENDER 로 읽는 것보다 비우는 게 낫다.
"""
import unittest

import app
from test_hkey import FakeSheet

# 출고 표준 배치: NO,DATE,CUSTOMER,Material Code,PART#,Q'ty,담당 SALES,REMARK,운송장번호,문서번호,VENDER
OUT_HDR = ("NO", "DATE", "CUSTOMER", "Material Code", "PART#", "Q'ty",
           "담당 SALES", "REMARK", "운송장번호", "문서번호")
IN_HDR = ("NO", "DATE", "SR#", "Material Code", "PART#", "Q'ty",
          "담당 SALES", "CUSTOMER", "Customer SRD", "REMARK", "FAB")


class TestVendorIdx(unittest.TestCase):
    def test_unnamed_k_column_when_layout_is_standard(self):
        cols = {app._hkey(h): i for i, h in enumerate(OUT_HDR)}
        self.assertEqual(app.vendor_idx(cols), 10)

    def test_named_header_wins(self):
        cols = {app._hkey(h): i for i, h in enumerate(OUT_HDR + ("VENDER",))}
        self.assertEqual(app.vendor_idx(cols), 10)
        cols = {app._hkey(h): i for i, h in enumerate(("VENDER",) + OUT_HDR)}
        self.assertEqual(app.vendor_idx(cols), 0)

    def test_accumulated_sheet_missing_middle_headers(self):
        """누적 '출고' 시트는 Material Code·운송장번호·문서번호 이름이 비어 있다."""
        hdr = ("NO", "DATE", "CUSTOMER", "", "PART#", "Q'ty", "담당 SALES", "REMARK", "")
        cols = {app._hkey(h): i for i, h in enumerate(hdr) if app.clean(h)}
        self.assertEqual(app.vendor_idx(cols), 10)

    def test_shifted_layout_gives_up(self):
        """칸이 하나 밀리면 위치 추정을 포기한다 (빈 값 > 엉뚱한 값)."""
        cols = {app._hkey(h): i for i, h in enumerate(("실",) + OUT_HDR)}
        self.assertIsNone(app.vendor_idx(cols))

    def test_inbound_layout_has_no_vendor(self):
        cols = {app._hkey(h): i for i, h in enumerate(IN_HDR)}
        self.assertIsNone(app.vendor_idx(cols))


class TestParseVendor(unittest.TestCase):
    ROWS = [
        ("입고", None),
        IN_HDR,
        ("1", "2026-09-15", "45HX260318-14", "UT-MR03183-000", "MT41K256M16TW-093:P",
         24000, "MARTIN", "HUMAX NETWORKS", "FCST", "202632(2,000)", "6"),
        ("출고", None),
        OUT_HDR,
        ("1", "2026-09-15", "Mobis(문백)", "UT-MR03184-000", "MT41K256M16TW-107 AA",
         64000, "조치현 책임매니저/화물발송", "M3203-001333", "화물", "2026-00134", "Micron"),
        # K열이 비어 있는 행도 있다 (예전 행·미기재)
        ("2", "2026-09-15", "Texon", "UT-DT00203-000", "THB2024HSKQT",
         95, "양현석 책임매니저/화물발송", "", "469818906", "2026-00176"),
    ]

    def test_outbound_reads_vendor_inbound_does_not(self):
        inbound, outbound = app.parse_daily(FakeSheet(self.ROWS))
        self.assertEqual([r["vendor"] for r in outbound], ["Micron", ""])
        self.assertEqual(inbound[0].get("vendor", ""), "")
        # 옆 칸을 밀어 읽지 않았는지 — 문서번호가 VENDER 로 새면 바로 티가 난다
        self.assertEqual(outbound[0]["doc"], "2026-00134")

    def test_log_sheet_reads_vendor(self):
        rows = [("NO", "DATE", "CUSTOMER", "", "PART#", "Q'ty", "담당 SALES", "REMARK", ""),
                ("1", "2026-09-14", "PAPYLUS", "UT-MR03819-000", "MT47H64M16NF-25E IT:",
                 4000, "김태헌 책임매니저/퀵 발송", "", "469840527", "2026-00143", "Micron")]
        got = app.parse_log(FakeSheet(rows))
        self.assertEqual([r["vendor"] for r in got], ["Micron"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
