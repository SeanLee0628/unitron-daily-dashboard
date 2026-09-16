# -*- coding: utf-8 -*-
"""당일 내역이 기간 조회에 잡히는지 — daily_log_rows().

자재팀이 18시에 날짜 시트를 갱신해도 누적 '입고'/'출고' 시트는 하루 늦게 따라온다.
누적 시트만 이력에 넣던 때는 기간 조회 달력이 어제까지만 열렸다.
"""
import datetime
import os
import tempfile
import unittest

os.environ["DATA_DIR"] = tempfile.mkdtemp()     # 실제 log.db 를 건드리지 않는다
import app                                      # noqa: E402


def shift(date, days):
    return (datetime.datetime.strptime(date, "%Y-%m-%d")
            + datetime.timedelta(days=days)).strftime("%Y-%m-%d")


class TestDailyLogRows(unittest.TestCase):
    def setUp(self):
        self.today = app.today_kst()
        self.yday = shift(self.today, -1)
        self.tmrw = shift(self.today, 1)
        self.per = {
            self.yday: ([dict(part="A1", qty=10, customer="MOBIS", sales="김")],
                        [dict(part="B1", qty=5, customer="LGE", sales="박")]),
            self.today: ([dict(part="A2", qty=20, customer="MOBIS", sales="김")],
                         [dict(part="B2", qty=7, customer="LGE", sales="박")]),
            self.tmrw: ([dict(part="A9", qty=99, customer="X", sales="김")], []),
        }
        self.by_dir = {
            "in": [dict(date=self.yday, part="A1", qty=10, customer="MOBIS",
                        sales="김", doc="SR-1", mcode="M1")],
            "out": [dict(date=self.yday, part="B1", qty=5, customer="LGE",
                         sales="박", doc="DO-1", mcode="M2")],
        }

    def test_takes_only_the_day_the_log_sheets_missed(self):
        """누적 시트에 없는 날짜(당일)만 가져온다."""
        got = app.daily_log_rows(self.per, self.by_dir)
        self.assertEqual([r["date"] for r in got["in"]], [self.today])
        self.assertEqual([r["date"] for r in got["out"]], [self.today])

    def test_skips_future_sheets(self):
        """빈 틀로 미리 만들어 둔 내일 시트는 이력에 넣지 않는다."""
        got = app.daily_log_rows(self.per, self.by_dir)
        self.assertNotIn(self.tmrw, [r["date"] for rs in got.values() for r in rs])

    def test_skips_direction_already_in_log_sheets(self):
        """방향별로 따로 판단한다 — 출고만 누적에 있으면 입고만 가져온다."""
        got = app.daily_log_rows(self.per, {"out": [dict(date=self.today, part="B2",
                                                        qty=7, customer="LGE",
                                                        sales="박")]})
        self.assertEqual([r["date"] for r in got["in"]], [self.yday, self.today])
        self.assertEqual([r["date"] for r in got["out"]], [self.yday])

    def test_empty_when_log_sheets_are_current(self):
        """누적 시트가 당일까지 반영돼 있으면 넣을 게 없다."""
        cur = {d: rows + [dict(r, date=self.today) for r in rows]
               for d, rows in self.by_dir.items()}
        self.assertEqual(app.daily_log_rows(self.per, cur), {})

    def test_no_log_sheets_at_all(self):
        """누적 시트가 아예 없는 파일도 당일·어제는 살린다."""
        got = app.daily_log_rows(self.per, None)
        self.assertEqual([r["date"] for r in got["in"]], [self.yday, self.today])


class TestDailyLogStore(unittest.TestCase):
    """이력 DB 왕복 — 달력 상한이 당일까지 열리고, 다음 날 정본으로 교체된다."""

    def setUp(self):
        self.today = app.today_kst()
        self.yday = shift(self.today, -1)
        self.office = "테스트실"
        cx = app.log_conn()
        cx.execute("DELETE FROM log WHERE slug=?", (app.office_slug(self.office),))
        cx.commit()
        cx.close()

    def test_span_opens_to_today_then_log_sheet_wins(self):
        by_dir = {"out": [dict(date=self.yday, part="B1", qty=5, customer="LGE",
                               sales="박", doc="DO-1")]}
        per = {self.today: ([], [dict(part="B2", qty=7, customer="LGE", sales="박")])}

        app.log_store(self.office, by_dir)
        self.assertEqual(app.log_span(self.office)["max"], self.yday)

        app.log_store(self.office, app.daily_log_rows(per, by_dir), src="daily")
        self.assertEqual(app.log_span(self.office)["max"], self.today)

        # 다음 날 누적 시트가 따라오면 그 날짜만 교체된다 (중복 누적 없음)
        by_dir["out"].append(dict(date=self.today, part="B2", qty=7, customer="LGE",
                                  sales="박", doc="DO-2"))
        app.log_store(self.office, by_dir)
        cx = app.log_conn()
        rows = cx.execute("SELECT part,doc,src FROM log WHERE slug=? AND date=?",
                          (app.office_slug(self.office), self.today)).fetchall()
        cx.close()
        self.assertEqual(rows, [("B2", "DO-2", "sheet")])


if __name__ == "__main__":
    unittest.main(verbosity=2)
