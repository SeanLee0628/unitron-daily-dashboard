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

    def test_4실_담당자가_받는다(self):
        """실 담당자는 자기 실 메일에만 들어간다. 다른 실에 새면 안 된다."""
        self.assertIn("jysong@unitrontech.com", app.OFFICE_EMAILS["4"])
        self.assertIn("sccho@unitrontech.com", app.OFFICE_EMAILS["4"])
        for slug in ("12", "3", "5"):
            self.assertNotIn("jysong@unitrontech.com", app.OFFICE_EMAILS[slug])
            self.assertNotIn("sccho@unitrontech.com", app.OFFICE_EMAILS[slug])

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
