#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""서버측 자동 발송 테스트.

원래 버그: 메일 발송이 업로더의 브라우저(autoSend)에서 돌았다. 업로드 직후 탭을
닫으면 서버는 데이터만 저장하고 메일은 한 통도 안 나갔고, 기록도 안 남았다.
발송은 /build 를 처리하는 서버 안에서 끝나야 한다.
"""
import unittest
from unittest import mock

import app


SAMPLE = {
    "offices": [
        {"name": "영업5실", "today_idx": 1,
         "days": [{"date": "2026-07-13", "kpi": {}}, {"date": "2026-07-14", "kpi": {}}]},
        {"name": "영업1,2실", "today_idx": 0,
         "days": [{"date": "2026-07-14", "kpi": {}}]},
        {"name": "전체 합계", "today_idx": 0,
         "days": [{"date": "2026-07-14", "kpi": {}}]},
    ]
}


class TestOfficeSlug(unittest.TestCase):
    def test_slug_pulls_digits(self):
        self.assertEqual(app.office_slug("영업5실"), "5")
        self.assertEqual(app.office_slug("영업1,2실"), "12")
        self.assertEqual(app.office_slug("전체 합계"), "all")


class TestSendOfficeEmails(unittest.TestCase):
    def setUp(self):
        self.sent = []
        patcher = mock.patch.object(
            app, "_send_smtp",
            # 참조(cc)가 나중에 붙었는데 이 가짜 함수가 안 받아서 세 건이
            # 조용히 실패하고 있었다. 인자를 열어 두면 다음에 뭐가 붙어도
            # 테스트가 그 이유로 깨지지는 않는다.
            side_effect=lambda emails, subject, body, cfg, text=None, **kw: (
                self.sent.append({"emails": emails, "subject": subject, "cc": kw.get("cc")}),
                {"ok": True, "sent": True, "via": "smtp", "n": len(emails)})[1])
        self.smtp = patcher.start()
        self.addCleanup(patcher.stop)
        cfg = mock.patch.object(app, "resolve_smtp", return_value={
            "host": "smtp.gmail.com", "port": 587,
            "user": "bot@x.com", "pw": "pw", "sender": "bot@x.com"})
        cfg.start()
        self.addCleanup(cfg.stop)

    def test_sends_one_email_per_office_skipping_total(self):
        res = app.send_office_emails(SAMPLE, "https://dash.example.com")
        self.assertEqual(len(self.sent), 2, "실별로 한 통씩 — '전체 합계'는 제외")
        self.assertEqual(res["ok"], ["영업5실", "영업1,2실"])
        self.assertEqual(res["fail"], [])

    def test_recipients_follow_office_map(self):
        """받는 사람은 실별로 다르고, 자재관리팀은 전 실 참조다.

        예전에는 sw.ahn 이 4·5실 '받는 사람' 이었고 이 테스트가 그걸 지켰다.
        2026-08-06 에 참조로 옮겼는데 테스트는 안 따라와서, 그 뒤로 계속
        실패하고 있었다. 지금 동작에 맞춘다.
        """
        app.send_office_emails(SAMPLE, "https://dash.example.com")
        for sent in self.sent:
            self.assertIn("seanlee@unitrontech.com", sent["emails"], "공통 수신자")
            self.assertNotIn("sw.ahn@unitrontech.com", sent["emails"],
                             "자재관리팀은 받는 사람이 아니라 참조다")
            self.assertIn("sw.ahn@unitrontech.com", sent["cc"] or [], "전 실 참조")

    def test_한_실_담당자는_다른_실에_새지_않는다(self):
        """한 실만 맡은 사람은 그 실 메일에만 들어간다.

        원래 이 테스트는 jysong·sccho 를 4실 전용으로 못박았는데, 2026-09-09 에
        두 사람이 1,2실·5실 담당자로도 들어가면서 계속 실패했다. 겸직은 실제
        상황이니 테스트가 따라간다 — 대신 지키려던 것(자기 실 밖으로 안 샌다)은
        한 실만 맡은 사람으로 검사한다.
        """
        # 2026-09-11 — 4실 대표를 sdpark → yk.kwon 으로 바꿨다. sdpark 이 5실
        # 담당자로도 들어가면서 '한 실만 맡은 사람' 이 아니게 됐다 (겸직 테스트로 옮김).
        전용 = {"12": "jini@unitrontech.com", "3": "bh.hwang@unitrontech.com",
               "4": "yk.kwon@unitrontech.com", "5": "cj.lim@unitrontech.com"}
        for slug, mail in 전용.items():
            self.assertIn(mail, app.OFFICE_EMAILS[slug], f"{slug}실 담당자")
            for other in set(app.OFFICE_EMAILS) - {slug}:
                self.assertNotIn(mail, app.OFFICE_EMAILS[other],
                                 f"{mail} 이 {other}실로 샜다")

    def test_겸직자는_맡은_실_전부에_들어간다(self):
        """여러 실을 맡은 사람은 그 실들에 모두 들어간다 (2026-09-09~10 명단)."""
        for mail, slugs in {
            "jysong@unitrontech.com": ("12", "4", "5"),
            "sccho@unitrontech.com": ("12", "4", "5"),
            "harold@unitrontech.com": ("12", "4", "5"),
            "lindsay@unitrontech.com": ("3", "4", "5"),
            "sdpark@unitrontech.com": ("4", "5"),          # 2026-09-11 5실 추가
        }.items():
            for slug in app.OFFICE_EMAILS:
                (self.assertIn if slug in slugs else self.assertNotIn)(
                    mail, app.OFFICE_EMAILS[slug], f"{mail} / {slug}실")

    def test_그룹_주소는_쓰지_않는다(self):
        """sales1@·sales3team@ 같은 그룹 주소로 보내면 그 실 밖으로 퍼지는데
        앱에서 막을 방법이 없다. 2026-09-09 에 전부 개인 주소로 바꿨다."""
        모든주소 = (app.DEFAULT_EMAILS + app.CC_EMAILS
                 + [m for v in app.SALES_EMAILS.values() for m in v])
        for mail in 모든주소:
            local = mail.split("@")[0].lower()
            self.assertFalse(local.startswith("sales") or local.endswith("team"),
                             f"그룹 주소로 보이는 수신자: {mail}")

    def test_오타_주소_linday_는_없다(self):
        """5실에 linday@ 로 잘못 들어가 있었다 (2026-09-10 정정)."""
        모든주소 = (app.DEFAULT_EMAILS + app.CC_EMAILS
                 + [m for v in app.SALES_EMAILS.values() for m in v])
        self.assertNotIn("linday@unitrontech.com", 모든주소)

    def test_uses_today_index_for_date(self):
        app.send_office_emails(SAMPLE, "https://dash.example.com")
        self.assertIn("2026-07-14", self.sent[0]["subject"], "today_idx 가 가리키는 날짜")

    def test_one_office_failing_does_not_stop_the_rest(self):
        self.smtp.side_effect = [
            ValueError("SMTP auth failed"),
            {"ok": True, "sent": True, "via": "smtp", "n": 1},
        ]
        res = app.send_office_emails(SAMPLE, "https://dash.example.com")
        self.assertEqual(res["fail"], ["영업5실 (SMTP auth failed)"])
        self.assertEqual(res["ok"], ["영업1,2실"], "앞의 실이 실패해도 뒤의 실은 나간다")


class TestTodayIndex(unittest.TestCase):
    """메일 제목 날짜가 하루 앞서던 문제.

    현장에서 다음 날 시트를 빈 틀로 미리 만들어 두는 바람에 마지막 시트 = 내일이
    됐다 (실측: 2026-09-03 18:17 발송, 제목은 2026-09-04).
    """

    def test_skips_tomorrows_sheet(self):
        with mock.patch.object(app, "today_kst", return_value="2026-09-03"):
            self.assertEqual(app.today_index(["2026-09-03", "2026-09-04"]), 0)

    def test_takes_todays_sheet_when_it_arrives(self):
        with mock.patch.object(app, "today_kst", return_value="2026-09-04"):
            self.assertEqual(app.today_index(["2026-09-03", "2026-09-04"]), 1)

    def test_falls_back_to_last_when_all_future(self):
        """날짜를 잘못 적은 파일 — 화면이 비는 것보다 마지막 시트가 낫다."""
        with mock.patch.object(app, "today_kst", return_value="2026-09-03"):
            self.assertEqual(app.today_index(["2026-09-05", "2026-09-06"]), 1)

    def test_stale_file_keeps_its_last_day(self):
        """며칠 지난 파일을 다시 열어도 마지막 자료일이 오늘 자리다."""
        with mock.patch.object(app, "today_kst", return_value="2026-09-07"):
            self.assertEqual(app.today_index(["2026-09-02", "2026-09-03"]), 1)

    def test_kst_not_server_utc(self):
        """Render 는 UTC 로 돈다. 한국 아침 8시(=UTC 전날 23시) 업로드에서
        오늘 시트가 미래로 밀리면 안 된다."""
        import datetime as dt
        fixed = dt.datetime(2026, 9, 3, 23, 0, tzinfo=dt.timezone.utc)

        class FakeDT(dt.datetime):
            @classmethod
            def now(cls, tz=None):
                return fixed.astimezone(tz) if tz else fixed

        with mock.patch.object(app.datetime, "datetime", FakeDT):
            self.assertEqual(app.today_kst(), "2026-09-04")

    def test_office_block_uses_it(self):
        with mock.patch.object(app, "today_kst", return_value="2026-09-03"):
            blk = app.office_block("영업5실", {"2026-09-03": ([], []), "2026-09-04": ([], [])})
        self.assertEqual(blk["today_idx"], 0)
        self.assertEqual(blk["days"][blk["today_idx"]]["date"], "2026-09-03")


class TestNoSmtpConfigured(unittest.TestCase):
    def test_reports_not_configured_instead_of_silently_passing(self):
        with mock.patch.object(app, "resolve_smtp", return_value=None):
            res = app.send_office_emails(SAMPLE, "https://dash.example.com")
        self.assertEqual(res["ok"], [])
        self.assertTrue(res["error"], "SMTP 설정이 없으면 에러를 남겨야 한다")


class TestMailcfgHonesty(unittest.TestCase):
    """configured:true 인데 실제로는 계정이 없어서 발송이 조용히 실패하던 문제."""

    def test_host_alone_is_not_configured(self):
        with mock.patch.dict(app.os.environ,
                             {"SMTP_HOST": "smtp.gmail.com", "SMTP_USER": "", "SMTP_PASS": ""},
                             clear=False):
            self.assertFalse(app.smtp_ready(), "호스트만 있고 계정이 없으면 준비된 게 아니다")

    def test_host_user_pass_is_configured(self):
        with mock.patch.dict(app.os.environ,
                             {"SMTP_HOST": "smtp.gmail.com", "SMTP_USER": "bot@x.com",
                              "SMTP_PASS": "pw"},
                             clear=False):
            self.assertTrue(app.smtp_ready())


if __name__ == "__main__":
    unittest.main(verbosity=2)
