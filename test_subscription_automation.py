from __future__ import annotations

import unittest
from datetime import datetime, timedelta, timezone

from subscription_automation import (
    ReminderMilestone,
    build_reminder_draft,
    parse_reminder_days,
    select_milestone,
)


class SubscriptionAutomationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.now = datetime(2026, 7, 16, 4, 0, tzinfo=timezone.utc)

    def test_parse_days_uses_safe_defaults(self) -> None:
        self.assertEqual(parse_reminder_days(""), (7, 3, 1))
        self.assertEqual(parse_reminder_days("1, 7,3,3,bad"), (7, 3, 1))

    def test_selects_nearest_relevant_window(self) -> None:
        self.assertEqual(
            select_milestone(self.now + timedelta(days=6), self.now),
            ReminderMilestone("d7", 6),
        )
        self.assertEqual(
            select_milestone(self.now + timedelta(days=2), self.now),
            ReminderMilestone("d3", 2),
        )
        self.assertEqual(
            select_milestone(self.now + timedelta(hours=2), self.now),
            ReminderMilestone("d1", 1),
        )

    def test_ignores_dates_outside_action_window(self) -> None:
        self.assertIsNone(select_milestone(self.now + timedelta(days=8), self.now))
        self.assertIsNone(select_milestone(self.now - timedelta(days=8), self.now))

    def test_expired_milestone_is_kept_for_review(self) -> None:
        self.assertEqual(
            select_milestone(self.now - timedelta(hours=1), self.now),
            ReminderMilestone("expired", 0, True),
        )

    def test_draft_supports_russian_and_uzbek(self) -> None:
        milestone = ReminderMilestone("d3", 3)
        ru = build_reminder_draft(
            lang="ru",
            active_until_text="20.07.2026 09:00",
            subscription_kind="paid",
            milestone=milestone,
            plans_text="1 месяц — 300 000 сум",
        )
        uz = build_reminder_draft(
            lang="uz",
            active_until_text="20.07.2026 09:00",
            subscription_kind="trial",
            milestone=milestone,
            plans_text="1 oy — 300 000 so‘m",
        )
        self.assertIn("Срок подписки", ru)
        self.assertIn("Obuna muddati", uz)
        self.assertIn("Sinov muddati", uz)


if __name__ == "__main__":
    unittest.main()
