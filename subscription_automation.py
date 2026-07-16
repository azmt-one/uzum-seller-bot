from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import datetime


DEFAULT_REMINDER_DAYS = (7, 3, 1)


@dataclass(frozen=True)
class ReminderMilestone:
    key: str
    days_remaining: int
    is_expired: bool = False


def parse_reminder_days(value: str | None) -> tuple[int, ...]:
    """Return unique positive reminder thresholds, from largest to smallest."""
    raw = str(value or "").replace(";", ",")
    result: set[int] = set()
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        try:
            days = int(part)
        except ValueError:
            continue
        if days > 0:
            result.add(days)
    return tuple(sorted(result or set(DEFAULT_REMINDER_DAYS), reverse=True))


def select_milestone(
    active_until: datetime,
    now: datetime,
    reminder_days: tuple[int, ...] = DEFAULT_REMINDER_DAYS,
    expired_queue_days: int = 7,
) -> ReminderMilestone | None:
    """Pick the most relevant reminder window for a subscription.

    A daily check can run at any hour. ``ceil`` keeps a user in the 1-day window
    until the exact expiration moment instead of silently dropping them early.
    """
    seconds_left = (active_until - now).total_seconds()
    if seconds_left <= 0:
        expired_for_days = abs(seconds_left) / 86400
        if expired_for_days <= max(1, int(expired_queue_days)):
            return ReminderMilestone("expired", 0, True)
        return None

    days_remaining = max(1, math.ceil(seconds_left / 86400))
    matching = sorted(day for day in reminder_days if days_remaining <= day)
    if not matching:
        return None
    threshold = matching[0]
    return ReminderMilestone(f"d{threshold}", days_remaining)


def build_reminder_draft(
    *,
    lang: str,
    active_until_text: str,
    subscription_kind: str,
    milestone: ReminderMilestone,
    plans_text: str,
) -> str:
    """Build a reviewable client message. This function never sends anything."""
    is_uz = str(lang or "ru").lower().startswith("uz")
    is_trial = subscription_kind == "trial"

    if is_uz:
        if milestone.is_expired:
            headline = "⛔ <b>Obuna muddati tugadi</b>"
            timing = f"Kirish muddati: <b>{active_until_text}</b> gacha edi."
        else:
            headline = "⏳ <b>Obuna muddati yaqinlashmoqda</b>"
            timing = (
                f"Kirish muddati tugashiga taxminan <b>{milestone.days_remaining} kun</b> qoldi.\n"
                f"Faol sana: <b>{active_until_text}</b> gacha."
            )
        kind = "Sinov muddati" if is_trial else "Obuna"
        return (
            f"{headline}\n\n"
            f"{kind}. {timing}\n\n"
            "Seller Pro uzluksiz ishlashi uchun obunani oldindan uzaytirishingiz mumkin.\n\n"
            f"<b>Tariflar:</b>\n{plans_text}\n\n"
            "To‘lov uchun: /subscribe"
        )

    if milestone.is_expired:
        headline = "⛔ <b>Срок подписки закончился</b>"
        timing = f"Доступ был активен до: <b>{active_until_text}</b>."
    else:
        headline = "⏳ <b>Срок подписки скоро закончится</b>"
        timing = (
            f"До окончания доступа осталось около <b>{milestone.days_remaining} дн.</b>\n"
            f"Доступ активен до: <b>{active_until_text}</b>."
        )
    kind = "Пробный период" if is_trial else "Подписка"
    return (
        f"{headline}\n\n"
        f"{kind}. {timing}\n\n"
        "Чтобы Seller Pro продолжил работать без перерыва, подписку можно продлить заранее.\n\n"
        f"<b>Тарифы:</b>\n{plans_text}\n\n"
        "Для оплаты: /subscribe"
    )
