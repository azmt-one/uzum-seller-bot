from __future__ import annotations

import asyncio
import hashlib
import json
import logging
import os
import tempfile
from datetime import datetime, timedelta, timezone
from html import escape
from urllib.parse import urlencode
from pathlib import Path
from typing import Any

from aiogram import Bot, Dispatcher, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import CallbackQuery, FSInputFile, InlineKeyboardButton, InlineKeyboardMarkup, KeyboardButton, Message, ReplyKeyboardMarkup
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from db import Database, TokenCipher
from formatters import (
    clean_num,
    compact_json_preview,
    excel_value,
    extract_items,
    flatten_sku_rows,
    format_order_line,
    format_product_line,
    format_shop_line,
    format_sku_stock_line,
    pick,
    safe,
    status_display,
)
from uzum_client import UzumClient

load_dotenv()
logging.basicConfig(level=logging.INFO)
logging.getLogger("httpx").setLevel(logging.WARNING)

TELEGRAM_BOT_TOKEN = (
    os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    or os.getenv("TELEGRAM_TOKEN", "").strip()
    or os.getenv("TOKEN", "").strip()
)
UZUM_API_BASE_URL = os.getenv(
    "UZUM_API_BASE_URL", "https://api-seller.uzum.uz/api/seller-openapi"
).strip()
DB_PATH = os.getenv("DB_PATH", "bot.db").strip()
ENCRYPTION_KEY = os.getenv("ENCRYPTION_KEY", "").strip()
ORDER_CHECK_INTERVAL_SECONDS = int(os.getenv("ORDER_CHECK_INTERVAL_SECONDS", "900") or "900")
NEW_ORDER_NOTIFICATIONS = (
    os.getenv("NEW_ORDER_NOTIFICATIONS", "0").strip().lower()
    not in {"0", "false", "no", "off"}
)
LOW_STOCK_NOTIFICATIONS = (
    os.getenv("LOW_STOCK_NOTIFICATIONS", "1").strip().lower()
    not in {"0", "false", "no", "off"}
)
LOW_STOCK_CHECK_INTERVAL_SECONDS = int(
    os.getenv("LOW_STOCK_CHECK_INTERVAL_SECONDS", "1800") or "1800"
)
LOW_STOCK_THRESHOLD = int(os.getenv("LOW_STOCK_THRESHOLD", "5") or "5")
OUT_OF_STOCK_NOTIFICATIONS = (
    os.getenv("OUT_OF_STOCK_NOTIFICATIONS", "1").strip().lower()
    not in {"0", "false", "no", "off"}
)
OUT_OF_STOCK_CHECK_INTERVAL_SECONDS = int(
    os.getenv("OUT_OF_STOCK_CHECK_INTERVAL_SECONDS", "1800") or "1800"
)
SALE_NOTIFICATIONS = (
    os.getenv("SALE_NOTIFICATIONS", "1").strip().lower()
    not in {"0", "false", "no", "off"}
)
SALE_CHECK_INTERVAL_SECONDS = int(os.getenv("SALE_CHECK_INTERVAL_SECONDS", "300") or "300")
STOCK_CHANGE_NOTIFICATIONS = (
    os.getenv("STOCK_CHANGE_NOTIFICATIONS", "0").strip().lower()
    not in {"0", "false", "no", "off"}
)
STOCK_CHANGE_CHECK_INTERVAL_SECONDS = int(
    os.getenv("STOCK_CHANGE_CHECK_INTERVAL_SECONDS", "900") or "900"
)
TRIAL_DAYS = int(os.getenv("TRIAL_DAYS", "3") or "3")
SUBSCRIPTION_PRICE_TEXT = os.getenv("SUBSCRIPTION_PRICE_TEXT", "250 000 сум / 1 месяц").strip()
PAYMENT_TEXT = os.getenv(
    "PAYMENT_TEXT",
    "Нажмите кнопку ниже, напишите администратору и отправьте чек. После проверки доступ будет продлён."
).strip()
SUBSCRIPTION_PLANS_TEXT = os.getenv(
    "SUBSCRIPTION_PLANS_TEXT",
    "1 месяц — 250 000 сум\n3 месяца — 650 000 сум\n6 месяцев — 1 200 000 сум\n\nБез ограничений по количеству магазинов продавца"
).strip()
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "azmt_one").strip().lstrip("@")
ADMIN_CONTACT_URL = os.getenv("ADMIN_CONTACT_URL", "").strip()
VIDEO_INSTRUCTION_URL = os.getenv("VIDEO_INSTRUCTION_URL", "https://t.me/uzum_assist_bot/2").strip()

# Подключение через сотрудника было экспериментом и отключено.
# Основной официальный способ: API-ключ продавца через /connect.
STAFF_UZUM_TOKEN = (
    os.getenv("STAFF_UZUM_TOKEN", "").strip()
    or os.getenv("TECHNICAL_UZUM_TOKEN", "").strip()
    or os.getenv("MASTER_UZUM_TOKEN", "").strip()
)
STAFF_PHONE = (
    os.getenv("STAFF_PHONE", "").strip()
    or os.getenv("BOT_STAFF_PHONE", "").strip()
    or os.getenv("TECHNICAL_STAFF_PHONE", "").strip()
)
STAFF_CONNECT_ENABLED = False
REPORT_INVOICE_PRODUCT_LIMIT = int(os.getenv("REPORT_INVOICE_PRODUCT_LIMIT", "10") or "10")
SMART_LOW_STOCK_DAYS = int(os.getenv("SMART_LOW_STOCK_DAYS", "3") or "3")
TOP_PRODUCTS_DAYS = int(os.getenv("TOP_PRODUCTS_DAYS", "30") or "30")
DEAD_STOCK_DAYS = int(os.getenv("DEAD_STOCK_DAYS", "30") or "30")
LOW_MARGIN_THRESHOLD_PERCENT = float(os.getenv("LOW_MARGIN_THRESHOLD_PERCENT", "10") or "10")
DAILY_REPORTS = (
    os.getenv("DAILY_REPORTS", "0").strip().lower()
    not in {"0", "false", "no", "off"}
)
DAILY_REPORT_HOUR_UZT = int(os.getenv("DAILY_REPORT_HOUR_UZT", "9") or "9")
SUBSCRIPTION_REMINDERS = (
    os.getenv("SUBSCRIPTION_REMINDERS", "1").strip().lower()
    not in {"0", "false", "no", "off"}
)
SUBSCRIPTION_REMINDER_DAYS = int(os.getenv("SUBSCRIPTION_REMINDER_DAYS", "1") or "1")

def _parse_admin_ids() -> set[int]:
    values: list[str] = []
    for key in ("ADMIN_IDS", "OWNER_TELEGRAM_ID", "OWNER_ID"):
        raw = os.getenv(key, "")
        if raw:
            values.extend(raw.replace(";", ",").split(","))
    ids: set[int] = set()
    for value in values:
        value = value.strip()
        if value.isdigit():
            ids.add(int(value))
    return ids

ADMIN_IDS = _parse_admin_ids()

if not TELEGRAM_BOT_TOKEN:
    raise RuntimeError(
        "TELEGRAM_BOT_TOKEN is empty. Set it in BotHost environment variables."
    )

# База и шифрование Uzum API-токена
db = Database(DB_PATH)
cipher = TokenCipher(ENCRYPTION_KEY)

bot = Bot(
    TELEGRAM_BOT_TOKEN,
    default=DefaultBotProperties(parse_mode=ParseMode.HTML),
)
dp = Dispatcher()

# --- Подписки / trial ---
def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


def _dt_to_db(dt: datetime | None) -> str | None:
    if dt is None:
        return None
    return dt.astimezone(timezone.utc).isoformat(timespec="seconds")


def _dt_from_db(value: Any) -> datetime | None:
    if not value:
        return None
    try:
        text = str(value).replace("Z", "+00:00")
        dt = datetime.fromisoformat(text)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        return None


def _fmt_dt(value: Any) -> str:
    dt = value if isinstance(value, datetime) else _dt_from_db(value)
    if not dt:
        return "—"
    return dt.astimezone(timezone(timedelta(hours=5))).strftime("%d.%m.%Y %H:%M")


def is_admin(telegram_id: int) -> bool:
    return int(telegram_id) in ADMIN_IDS


def init_subscription_tables() -> None:
    with db.connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS subscriptions (
                telegram_id INTEGER PRIMARY KEY,
                trial_started_at TEXT,
                trial_until TEXT,
                subscription_until TEXT,
                blocked INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.commit()


def get_subscription_row(telegram_id: int) -> dict[str, Any] | None:
    with db.connect() as conn:
        row = conn.execute("SELECT * FROM subscriptions WHERE telegram_id = ?", (int(telegram_id),)).fetchone()
    return dict(row) if row else None


def ensure_subscription(telegram_id: int) -> dict[str, Any]:
    row = get_subscription_row(telegram_id)
    if row:
        return row
    now = _utc_now()
    trial_until = now + timedelta(days=TRIAL_DAYS)
    with db.connect() as conn:
        conn.execute(
            """
            INSERT OR IGNORE INTO subscriptions
            (telegram_id, trial_started_at, trial_until, subscription_until, blocked, created_at, updated_at)
            VALUES (?, ?, ?, NULL, 0, ?, ?)
            """,
            (int(telegram_id), _dt_to_db(now), _dt_to_db(trial_until), _dt_to_db(now), _dt_to_db(now)),
        )
        conn.commit()
    return get_subscription_row(telegram_id) or {}


def subscription_active_until(row: dict[str, Any] | None) -> datetime | None:
    if not row:
        return None
    dates = [_dt_from_db(row.get("trial_until")), _dt_from_db(row.get("subscription_until"))]
    dates = [d for d in dates if d is not None]
    return max(dates) if dates else None


def has_active_subscription(telegram_id: int) -> bool:
    if is_admin(telegram_id):
        return True
    row = ensure_subscription(telegram_id)
    if int(row.get("blocked") or 0) == 1:
        return False
    until = subscription_active_until(row)
    return bool(until and until > _utc_now())


def subscription_status_text(telegram_id: int) -> str:
    row = ensure_subscription(telegram_id)
    if is_admin(telegram_id):
        return "👑 Админ-доступ: без ограничений"
    if int(row.get("blocked") or 0) == 1:
        return "⛔ Пользователь заблокирован"
    now = _utc_now()
    trial_until = _dt_from_db(row.get("trial_until"))
    paid_until = _dt_from_db(row.get("subscription_until"))
    until = subscription_active_until(row)
    if until and until > now:
        if paid_until and paid_until == until:
            return f"✅ Подписка активна до {_fmt_dt(paid_until)}"
        return f"🎁 Trial активен до {_fmt_dt(trial_until)}"
    return "⛔ Подписка закончилась"


def subscription_full_text(telegram_id: int) -> str:
    row = ensure_subscription(telegram_id)
    status = subscription_status_text(telegram_id)

    if is_admin(telegram_id):
        return (
            "💎 <b>Моя подписка</b>\n\n"
            f"Telegram ID: <code>{telegram_id}</code>\n"
            f"Статус: {status}\n\n"
            "Trial и дата оплаты для администратора не важны — доступ всегда открыт.\n\n"
            "Команды администратора:\n"
            "• <code>/users</code> — пользователи\n"
            "• <code>/extend ID 30</code> — продлить доступ\n"
            "• <code>/block ID</code> — заблокировать\n"
            "• <code>/unblock ID</code> — разблокировать\n"
            "• <code>/paid ID сумма дни</code> — записать оплату\n"
            "• <code>/payments</code> — история оплат\n"
            "• <code>/backup_db</code> — скачать базу"
        )

    return (
        "💎 <b>Моя подписка</b>\n\n"
        f"Telegram ID: <code>{telegram_id}</code>\n"
        f"Статус: {status}\n"
        f"Trial до: <b>{_fmt_dt(row.get('trial_until'))}</b>\n"
        f"Оплачено до: <b>{_fmt_dt(row.get('subscription_until'))}</b>\n\n"
        "Тарифы:\n"
        f"<b>{escape(SUBSCRIPTION_PLANS_TEXT)}</b>\n\n"
        f"{escape(PAYMENT_TEXT)}\n\n"
        "История оплат: <code>/my_payments</code>\n"
        "Поддержка: <code>/support</code>\n"
        "Заменить API-ключ: <code>/reconnect</code>\n"
        "Удалить API-ключ: <code>/disconnect</code>"
    )


async def require_active_subscription(message: Message, telegram_id: int | None = None) -> bool:
    if telegram_id is None:
        telegram_id = upsert_from_message(message)
    ensure_subscription(int(telegram_id))
    if has_active_subscription(int(telegram_id)):
        return True
    await message.answer(
        tr_user(int(telegram_id), "access_limited"),
        reply_markup=menu_for_message(message),
    )
    return False


def admin_only(telegram_id: int) -> bool:
    return is_admin(int(telegram_id))


def admin_contact_link() -> str | None:
    if ADMIN_CONTACT_URL:
        return ADMIN_CONTACT_URL
    if ADMIN_USERNAME:
        return f"https://t.me/{ADMIN_USERNAME}"
    return None


def admin_contact_markup() -> InlineKeyboardMarkup | None:
    url = admin_contact_link()
    if not url:
        return None
    return InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text="✍️ Написать администратору", url=url)]]
    )


def video_instruction_markup(lang: str = "ru") -> InlineKeyboardMarkup | None:
    if not VIDEO_INSTRUCTION_URL:
        return None
    button_text = "▶️ Videoni ko‘rish" if normalize_lang(lang) == "uz" else "▶️ Смотреть видео"
    return InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text=button_text, url=VIDEO_INSTRUCTION_URL)]]
    )


def help_links_markup(lang: str = "ru") -> InlineKeyboardMarkup | None:
    rows = []
    if VIDEO_INSTRUCTION_URL:
        rows.append([InlineKeyboardButton(
            text=("🎥 API ulash videosi" if normalize_lang(lang) == "uz" else "🎥 Видео подключения API"),
            url=VIDEO_INSTRUCTION_URL,
        )])
    admin_url = admin_contact_link()
    if admin_url:
        rows.append([InlineKeyboardButton(
            text=("✍️ Administratorga yozish" if normalize_lang(lang) == "uz" else "✍️ Написать администратору"),
            url=admin_url,
        )])
    if not rows:
        return None
    return InlineKeyboardMarkup(inline_keyboard=rows)


def admin_contact_text() -> str:
    if ADMIN_USERNAME:
        return f"@{escape(ADMIN_USERNAME)}"
    if ADMIN_CONTACT_URL:
        return escape(ADMIN_CONTACT_URL)
    return "администратору"


def extend_subscription_days(telegram_id: int, days: int) -> datetime:
    ensure_subscription(telegram_id)
    row = get_subscription_row(telegram_id) or {}
    now = _utc_now()
    candidates = [now, _dt_from_db(row.get("subscription_until")), _dt_from_db(row.get("trial_until"))]
    base = max([d for d in candidates if d is not None])
    new_until = base + timedelta(days=max(1, int(days)))
    with db.connect() as conn:
        conn.execute(
            "UPDATE subscriptions SET subscription_until = ?, blocked = 0, updated_at = ? WHERE telegram_id = ?",
            (_dt_to_db(new_until), _dt_to_db(now), int(telegram_id)),
        )
        conn.commit()
    return new_until


def set_trial_days(telegram_id: int, days: int) -> datetime:
    ensure_subscription(telegram_id)
    now = _utc_now()
    new_until = now + timedelta(days=max(1, int(days)))
    with db.connect() as conn:
        conn.execute(
            "UPDATE subscriptions SET trial_until = ?, blocked = 0, updated_at = ? WHERE telegram_id = ?",
            (_dt_to_db(new_until), _dt_to_db(now), int(telegram_id)),
        )
        conn.commit()
    return new_until


def set_blocked(telegram_id: int, blocked: bool) -> None:
    ensure_subscription(telegram_id)
    now = _utc_now()
    with db.connect() as conn:
        conn.execute(
            "UPDATE subscriptions SET blocked = ?, updated_at = ? WHERE telegram_id = ?",
            (1 if blocked else 0, _dt_to_db(now), int(telegram_id)),
        )
        conn.commit()




def init_business_tables() -> None:
    with db.connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS payment_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                telegram_id INTEGER NOT NULL,
                amount INTEGER NOT NULL DEFAULT 0,
                days INTEGER NOT NULL DEFAULT 0,
                admin_id INTEGER,
                comment TEXT,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.commit()


def record_payment(telegram_id: int, amount: int, days: int, admin_id: int | None = None, comment: str = "") -> int:
    init_business_tables()
    with db.connect() as conn:
        cur = conn.execute(
            """
            INSERT INTO payment_history (telegram_id, amount, days, admin_id, comment, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (int(telegram_id), int(amount), int(days), int(admin_id) if admin_id else None, comment.strip(), _dt_to_db(_utc_now())),
        )
        conn.commit()
        return int(cur.lastrowid)


def list_payments(telegram_id: int | None = None, limit: int = 20) -> list[dict[str, Any]]:
    init_business_tables()
    with db.connect() as conn:
        if telegram_id is None:
            rows = conn.execute(
                """
                SELECT id, telegram_id, amount, days, admin_id, comment, created_at
                FROM payment_history
                ORDER BY id DESC
                LIMIT ?
                """,
                (int(limit),),
            ).fetchall()
        else:
            rows = conn.execute(
                """
                SELECT id, telegram_id, amount, days, admin_id, comment, created_at
                FROM payment_history
                WHERE telegram_id = ?
                ORDER BY id DESC
                LIMIT ?
                """,
                (int(telegram_id), int(limit)),
            ).fetchall()
    return [dict(row) for row in rows]


def payment_line(row: dict[str, Any]) -> str:
    comment = (row.get("comment") or "").strip()
    comment_part = f" | {escape(comment)}" if comment else ""
    amount_text = f"{int(row.get('amount') or 0):,}".replace(",", " ")
    return (
        f"#{row.get('id')} | <code>{int(row.get('telegram_id') or 0)}</code> | "
        f"{amount_text} сум | {int(row.get('days') or 0)} дней | {_fmt_dt(row.get('created_at'))}{comment_part}"
    )



# --- Юнит-экономика / себестоимость SKU ---
def init_unit_economy_tables() -> None:
    with db.connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS unit_costs (
                telegram_id INTEGER NOT NULL,
                shop_id INTEGER NOT NULL,
                sku_key TEXT NOT NULL,
                title TEXT,
                cost REAL NOT NULL DEFAULT 0,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (telegram_id, shop_id, sku_key)
            )
            """
        )
        conn.commit()


def _unit_sku_key(value: Any) -> str:
    return str(value or "").strip().lower()


def save_unit_cost(telegram_id: int, shop_id: int, sku: str, cost: float, title: str = "") -> None:
    init_unit_economy_tables()
    sku_key = _unit_sku_key(sku)
    if not sku_key:
        return
    with db.connect() as conn:
        conn.execute(
            """
            INSERT INTO unit_costs (telegram_id, shop_id, sku_key, title, cost, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(telegram_id, shop_id, sku_key) DO UPDATE SET
                title = excluded.title,
                cost = excluded.cost,
                updated_at = excluded.updated_at
            """,
            (int(telegram_id), int(shop_id), sku_key, str(title or "").strip(), float(cost), _dt_to_db(_utc_now()) or ""),
        )
        conn.commit()


def delete_unit_cost(telegram_id: int, shop_id: int, sku: str) -> bool:
    init_unit_economy_tables()
    sku_key = _unit_sku_key(sku)
    with db.connect() as conn:
        cur = conn.execute(
            "DELETE FROM unit_costs WHERE telegram_id = ? AND shop_id = ? AND sku_key = ?",
            (int(telegram_id), int(shop_id), sku_key),
        )
        conn.commit()
        return bool(cur.rowcount)


def get_unit_cost_map(telegram_id: int, shop_id: int) -> dict[str, dict[str, Any]]:
    init_unit_economy_tables()
    with db.connect() as conn:
        rows = conn.execute(
            "SELECT sku_key, title, cost, updated_at FROM unit_costs WHERE telegram_id = ? AND shop_id = ?",
            (int(telegram_id), int(shop_id)),
        ).fetchall()
    return {str(r["sku_key"]): dict(r) for r in rows}


def list_unit_costs(telegram_id: int, shop_id: int, limit: int = 50) -> list[dict[str, Any]]:
    init_unit_economy_tables()
    with db.connect() as conn:
        rows = conn.execute(
            """
            SELECT sku_key, title, cost, updated_at
            FROM unit_costs
            WHERE telegram_id = ? AND shop_id = ?
            ORDER BY updated_at DESC
            LIMIT ?
            """,
            (int(telegram_id), int(shop_id), int(limit)),
        ).fetchall()
    return [dict(r) for r in rows]


def _parse_cost_command_args(text: str) -> tuple[str, float] | None:
    import re
    raw = parse_args(text or "").strip()
    # Формат: /cost SKU 60000 или /cost SKU 60 000
    m = re.match(r"^(.+?)\s+([0-9][0-9\s.,]*)$", raw)
    if not m:
        return None
    sku = m.group(1).strip()
    money_raw = m.group(2)
    digits = re.sub(r"[^0-9]", "", money_raw)
    if not sku or not digits:
        return None
    return sku, float(digits)

def list_subscription_users(limit: int = 30) -> list[dict[str, Any]]:
    with db.connect() as conn:
        rows = conn.execute(
            """
            SELECT u.telegram_id, u.username, u.first_name, u.default_shop_id,
                   s.trial_until, s.subscription_until, s.blocked
            FROM users u
            LEFT JOIN subscriptions s ON s.telegram_id = u.telegram_id
            ORDER BY u.updated_at DESC
            LIMIT ?
            """,
            (int(limit),),
        ).fetchall()
    return [dict(r) for r in rows]


def subscription_compact_line(row: dict[str, Any]) -> str:
    telegram_id = int(row.get("telegram_id"))
    username = row.get("username") or ""
    name = row.get("first_name") or ""
    label = f"@{username}" if username else (name or "без имени")
    if row.get("blocked"):
        status_label = "⛔ block"
    else:
        untils = [_dt_from_db(row.get("subscription_until")), _dt_from_db(row.get("trial_until"))]
        untils = [d for d in untils if d]
        until = max(untils) if untils else None
        status_label = "✅" if until and until > _utc_now() else "❌"
    until_value = row.get("subscription_until") or row.get("trial_until")
    return f"{status_label} <code>{telegram_id}</code> — {escape(str(label))} | до: {_fmt_dt(until_value)}"


def _subscription_until_for_row(row: dict[str, Any]) -> datetime | None:
    return subscription_active_until(row)


def get_admin_stats() -> dict[str, int]:
    now = _utc_now()
    with db.connect() as conn:
        total_users = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        connected = conn.execute("SELECT COUNT(*) FROM users WHERE uzum_token_encrypted IS NOT NULL").fetchone()[0]
        rows = conn.execute("SELECT telegram_id, trial_until, subscription_until, blocked FROM subscriptions").fetchall()
        payments_today = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) FROM payment_history WHERE created_at >= ?",
            (_dt_to_db(now.replace(hour=0, minute=0, second=0, microsecond=0)),),
        ).fetchone()[0]
        payments_30 = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) FROM payment_history WHERE created_at >= ?",
            (_dt_to_db(now - timedelta(days=30)),),
        ).fetchone()[0]
    active = paid = trial = expired = blocked = 0
    for r in rows:
        row = dict(r)
        if int(row.get("blocked") or 0) == 1:
            blocked += 1
            continue
        until = subscription_active_until(row)
        if until and until > now:
            active += 1
            paid_until = _dt_from_db(row.get("subscription_until"))
            if paid_until and paid_until == until:
                paid += 1
            else:
                trial += 1
        else:
            expired += 1
    return {
        "total_users": int(total_users or 0),
        "connected": int(connected or 0),
        "active": active,
        "paid": paid,
        "trial": trial,
        "expired": expired,
        "blocked": blocked,
        "payments_today": int(payments_today or 0),
        "payments_30": int(payments_30 or 0),
    }


def list_expiring_users(days: int = 3, limit: int = 50) -> list[dict[str, Any]]:
    now = _utc_now()
    until_limit = now + timedelta(days=max(1, int(days)))
    with db.connect() as conn:
        rows = conn.execute(
            """
            SELECT u.telegram_id, u.username, u.first_name, u.default_shop_id,
                   s.trial_until, s.subscription_until, s.blocked
            FROM subscriptions s
            LEFT JOIN users u ON u.telegram_id = s.telegram_id
            WHERE s.blocked = 0
            ORDER BY COALESCE(s.subscription_until, s.trial_until) ASC
            LIMIT ?
            """,
            (int(limit),),
        ).fetchall()
    result: list[dict[str, Any]] = []
    for r in rows:
        row = dict(r)
        if is_admin(int(row.get("telegram_id") or 0)):
            continue
        until = subscription_active_until(row)
        if until and now < until <= until_limit:
            result.append(row)
    return result


def list_blocked_users(limit: int = 50) -> list[dict[str, Any]]:
    with db.connect() as conn:
        rows = conn.execute(
            """
            SELECT u.telegram_id, u.username, u.first_name, u.default_shop_id,
                   s.trial_until, s.subscription_until, s.blocked
            FROM subscriptions s
            LEFT JOIN users u ON u.telegram_id = s.telegram_id
            WHERE s.blocked = 1
            ORDER BY s.updated_at DESC
            LIMIT ?
            """,
            (int(limit),),
        ).fetchall()
    return [dict(r) for r in rows]



def init_staff_connect_tables() -> None:
    with db.connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS staff_shop_connections (
                telegram_id INTEGER NOT NULL,
                shop_id INTEGER NOT NULL,
                status TEXT NOT NULL DEFAULT 'pending',
                error TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (telegram_id, shop_id)
            )
            """
        )
        conn.commit()


def save_staff_shop_status(telegram_id: int, shop_id: int, status: str, error: str = "") -> None:
    init_staff_connect_tables()
    now = _dt_to_db(_utc_now()) or ""
    with db.connect() as conn:
        conn.execute(
            """
            INSERT INTO staff_shop_connections (telegram_id, shop_id, status, error, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(telegram_id, shop_id) DO UPDATE SET
                status = excluded.status,
                error = excluded.error,
                updated_at = excluded.updated_at
            """,
            (int(telegram_id), int(shop_id), str(status), str(error)[:1000], now, now),
        )
        conn.commit()


def list_staff_shop_connections(limit: int = 30) -> list[dict[str, Any]]:
    init_staff_connect_tables()
    with db.connect() as conn:
        rows = conn.execute(
            """
            SELECT telegram_id, shop_id, status, error, created_at, updated_at
            FROM staff_shop_connections
            ORDER BY updated_at DESC
            LIMIT ?
            """,
            (int(limit),),
        ).fetchall()
    return [dict(r) for r in rows]


init_subscription_tables()
init_business_tables()
init_unit_economy_tables()
init_staff_connect_tables()

# --- Языки интерфейса ---
# Основной код отчётов остаётся совместимым с русскими командами, но клиент может выбрать язык меню и основных экранов.
SUPPORTED_LANGUAGES = {"ru", "uz"}


def init_language_tables() -> None:
    with db.connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS user_language (
                telegram_id INTEGER PRIMARY KEY,
                lang TEXT NOT NULL DEFAULT 'ru',
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.commit()


def normalize_lang(value: Any) -> str:
    lang = str(value or "ru").strip().lower()
    if lang.startswith("uz"):
        return "uz"
    return "ru"


def get_user_language(telegram_id: int | None) -> str:
    if telegram_id is None:
        return "ru"
    try:
        with db.connect() as conn:
            row = conn.execute(
                "SELECT lang FROM user_language WHERE telegram_id = ?",
                (int(telegram_id),),
            ).fetchone()
        if row:
            return normalize_lang(row[0] if not isinstance(row, dict) else row.get("lang"))
    except Exception:
        logging.exception("Failed to read user language")
    return "ru"


def set_user_language(telegram_id: int, lang: str) -> None:
    lang = normalize_lang(lang)
    now = _dt_to_db(_utc_now()) or ""
    with db.connect() as conn:
        conn.execute(
            """
            INSERT INTO user_language (telegram_id, lang, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(telegram_id) DO UPDATE SET lang = excluded.lang, updated_at = excluded.updated_at
            """,
            (int(telegram_id), lang, now),
        )
        conn.commit()



# Создаём таблицу языков после определения функции, чтобы не было NameError при запуске.

def language_title(lang: str) -> str:
    return "O‘zbekcha" if normalize_lang(lang) == "uz" else "Русский"


I18N: dict[str, dict[str, str]] = {
    "ru": {
        "choose_action": "Выберите действие",
        "choose_section": "Выберите раздел 👇",
        "main_menu": "Главное меню 👇",
        "cancelled": "Действие отменено.",
        "language_title": "🌐 <b>Язык интерфейса</b>",
        "language_body": "Выберите язык, на котором бот будет показывать меню и основные подсказки.",
        "language_set": "✅ Язык изменён: <b>Русский</b>",
        "language_button_ru": "🇷🇺 Русский",
        "language_button_uz": "🇺🇿 O‘zbekcha",
        "admin_only": "⛔ Админ-панель доступна только владельцу бота.",
        "access_limited": "⛔ <b>Доступ ограничен</b>\n\nTrial или подписка закончились.\nВаш Uzum-токен и настройки сохранены — после продления всё снова заработает.\n\nПроверить подписку: <code>/my_subscription</code>\nОплата: <code>/subscribe</code>",
        "connect_first": "Сначала подключите Uzum API-токен: <code>/connect</code>",
        "connection_deleted": "✅ Подключение к Uzum API удалено. Можно подключить заново через <code>/connect</code>.",
        "token_instruction_title": "🔑 <b>Где взять Uzum Seller API-ключ</b>",
        "support_title": "🆘 <b>Поддержка</b>",
        "security_title": "🔐 <b>Безопасность API-ключа</b>",
    },
    "uz": {
        "choose_action": "Amalni tanlang",
        "choose_section": "Bo‘limni tanlang 👇",
        "main_menu": "Asosiy menyu 👇",
        "cancelled": "Amal bekor qilindi.",
        "language_title": "🌐 <b>Interfeys tili</b>",
        "language_body": "Bot menyu va asosiy ko‘rsatmalarni qaysi tilda ko‘rsatishini tanlang.",
        "language_set": "✅ Til o‘zgartirildi: <b>O‘zbekcha</b>",
        "language_button_ru": "🇷🇺 Русский",
        "language_button_uz": "🇺🇿 O‘zbekcha",
        "admin_only": "⛔ Admin panel faqat bot egasi uchun.",
        "access_limited": "⛔ <b>Kirish cheklangan</b>\n\nTrial yoki obuna muddati tugagan.\nUzum tokeningiz va sozlamalaringiz saqlanadi — obuna uzaytirilgach hammasi yana ishlaydi.\n\nObunani tekshirish: <code>/my_subscription</code>\nTo‘lov: <code>/subscribe</code>",
        "connect_first": "Avval Uzum API-kalitini ulang: <code>/connect</code>",
        "connection_deleted": "✅ Uzum API ulanishi o‘chirildi. Qayta ulash uchun <code>/connect</code> buyrug‘idan foydalaning.",
        "token_instruction_title": "🔑 <b>Uzum Seller API-kalitini qayerdan olish mumkin</b>",
        "support_title": "🆘 <b>Yordam</b>",
        "security_title": "🔐 <b>API-kalit xavfsizligi</b>",
    },
}


def tr(lang: str, key: str) -> str:
    lang = normalize_lang(lang)
    return I18N.get(lang, I18N["ru"]).get(key, I18N["ru"].get(key, key))


def tr_user(telegram_id: int | None, key: str) -> str:
    return tr(get_user_language(telegram_id), key)



# --- Автоперевод сообщений с данными на узбекский ---
def translate_runtime_text_to_uz(text: str) -> str:
    """Лёгкий пост-процессор: переводит основные русские ответы и отчёты на узбекский.

    Меню уже переключается отдельными клавиатурами. Этот слой нужен для старых
    функций, где текст отчётов был собран на русском внутри бизнес-логики.
    Числа, ID, SKU, статусы и суммы не меняются.
    """
    if not isinstance(text, str) or not text:
        return text

    replacements = [
        # waiting / service
        ("⌛ Считаю баланс за 30 дней...", "⌛ 30 kunlik balans hisoblanmoqda..."),
        ("⌛ Считаю баланс по всем магазинам за 30 дней...", "⌛ Barcha do‘konlar bo‘yicha 30 kunlik balans hisoblanmoqda..."),
        ("⌛ Считаю продажи за сегодня...", "⌛ Bugungi sotuvlar hisoblanmoqda..."),
        ("⌛ Считаю продажи за вчера...", "⌛ Kechagi sotuvlar hisoblanmoqda..."),
        ("⌛ Считаю продажи за 7 дней...", "⌛ 7 kunlik sotuvlar hisoblanmoqda..."),
        ("⏳ Считаю продажи за сегодня, 7 и 30 дней...", "⏳ Bugun, 7 kun va 30 kunlik sotuvlar hisoblanmoqda..."),
        ("⏳ Считаю заказы по статусам...", "⏳ Buyurtmalar statuslar bo‘yicha hisoblanmoqda..."),
        ("⌛ Считаю топ товаров", "⌛ Top tovarlar hisoblanmoqda"),
        ("⌛ Считаю, на сколько дней хватит остатков...", "⌛ Qoldiq necha kunga yetishi hisoblanmoqda..."),
        ("⏳ Готовлю Excel-отчёт...", "⏳ Excel hisobot tayyorlanmoqda..."),
        ("⏳ Собираю утренний отчёт...", "⏳ Ertalabki hisobot tayyorlanmoqda..."),

        # titles
        ("💰 <b>Баланс Uzum FBO за 30 дней</b>", "💰 <b>Uzum FBO balansi 30 kun uchun</b>"),
        ("💰 <b>Продажи Uzum FBO/FBS за сегодня</b>", "💰 <b>Bugungi Uzum FBO/FBS sotuvlari</b>"),
        ("💰 <b>Продажи Uzum FBO/FBS за вчера</b>", "💰 <b>Kechagi Uzum FBO/FBS sotuvlari</b>"),
        ("💰 <b>Продажи Uzum FBO/FBS за 7 дней</b>", "💰 <b>7 kunlik Uzum FBO/FBS sotuvlari</b>"),
        ("💰 <b>Продажи Uzum FBO/FBS за 30 дней</b>", "💰 <b>30 kunlik Uzum FBO/FBS sotuvlari</b>"),
        ("🌐 <b>Баланс по всем магазинам за 30 дней</b>", "🌐 <b>Barcha do‘konlar bo‘yicha 30 kunlik balans</b>"),
        ("📊 <b>Сводка продаж</b>", "📊 <b>Sotuvlar xulosasi</b>"),
        ("📊 <b>Сводка заказов</b>", "📊 <b>Buyurtmalar xulosasi</b>"),
        ("📦 <b>Остатки</b>", "📦 <b>Qoldiq</b>"),
        ("⚠️ <b>Умное 'Заканчивается'</b>", "⚠️ <b>Qoldiq prognozi</b>"),
        ("🏆 <b>Топ товаров", "🏆 <b>Top tovarlar"),
        ("🐢 <b>Товары без продаж", "🐢 <b>Sotilmayotgan tovarlar"),
        ("🏪 <b>Ваши магазины</b>", "🏪 <b>Do‘konlaringiz</b>"),
        ("📄 <b>FBO-накладные поставки</b>", "📄 <b>FBO yuk xatlari</b>"),
        ("📦 <b>Состав FBO-накладной</b>", "📦 <b>FBO yuk xati tarkibi</b>"),
        ("🌙 <b>Утренний отчёт Uzum</b>", "🌙 <b>Uzum ertalabki hisoboti</b>"),
        ("🛒 <b>Новая продажа Uzum FBO</b>", "🛒 <b>Yangi Uzum FBO sotuvi</b>"),
        ("⚠️ <b>Заканчиваются товары</b>", "⚠️ <b>Tovarlar tugayapti</b>"),
        ("❌ <b>Товары закончились</b>", "❌ <b>Tovarlar tugagan</b>"),
        ("💎 <b>Моя подписка</b>", "💎 <b>Mening obunam</b>"),
        ("👑 <b>Админ-панель</b>", "👑 <b>Admin panel</b>"),

        # labels, finance
        ("Магазинов найдено:", "Topilgan do‘konlar:"),
        ("Магазинов:", "Do‘konlar soni:"),
        ("Магазин:", "Do‘kon:"),
        ("Текущий магазин:", "Joriy do‘kon:"),
        ("Активный магазин:", "Faol do‘kon:"),
        ("Позиции продаж:", "Sotuv pozitsiyalari:"),
        ("Кол-во товаров:", "Tovarlar soni:"),
        ("Возвраты:", "Qaytarishlar:"),
        ("Выручка:", "Tushum:"),
        ("Комиссия Uzum:", "Uzum komissiyasi:"),
        ("Комиссия:", "Komissiya:"),
        ("Логистика:", "Logistika:"),
        ("К выплате всего:", "Jami to‘lovga:"),
        ("К выплате:", "To‘lovga:"),
        ("Уже выведено:", "Allaqachon chiqarilgan:"),
        ("Остаток к выплате:", "To‘lovga qoldi:"),
        ("Статусы:", "Statuslar:"),
        ("Цена продажи:", "Sotuv narxi:"),
        ("ID заказа:", "Buyurtma ID:"),
        ("ID продажи:", "Sotuv ID:"),
        ("Статус:", "Status:"),
        ("Дата:", "Sana:"),
        ("Товар:", "Tovar:"),
        ("Кол-во:", "Soni:"),
        ("Кол-во товаров", "Tovarlar soni"),
        ("Позиции продаж", "Sotuv pozitsiyalari"),
        ("Возвраты", "Qaytarishlar"),
        ("Выручка", "Tushum"),
        ("Логистика", "Logistika"),
        ("Комиссия", "Komissiya"),
        ("К выплате", "To‘lovga"),
        ("Уже выведено", "Allaqachon chiqarilgan"),
        ("Остаток к выплате", "To‘lovga qoldi"),
        ("Позиций продаж", "Sotuv pozitsiyalari"),

        # products / stock
        ("Всего товаров:", "Jami tovarlar:"),
        ("Всего:", "Jami:"),
        ("Остаток:", "Qoldiq:"),
        ("Итого:", "Jami:"),
        ("Разница:", "Farq:"),
        ("Проверить остатки:", "Qoldiqni tekshirish:"),
        ("Уменьшился остаток по SKU", "SKU bo‘yicha qoldiq kamaydi"),
        ("Это может быть продажа, резерв, списание или изменение склада.", "Bu sotuv, rezerv, hisobdan chiqarish yoki ombor o‘zgarishi bo‘lishi mumkin."),
        ("Товары, которые заканчиваются", "Tugayotgan tovarlar"),
        ("Остаток меньше или равен", "Qoldiq kam yoki teng"),
        ("Товар закончился", "Tovar tugagan"),
        ("Потерянные товары", "Yo‘qolgan tovarlar"),
        ("Потеряно:", "Yo‘qolgan:"),
        ("Примерная сумма:", "Taxminiy summa:"),
        ("Продано:", "Sotilgan:"),
        ("Продаж не найдено.", "Sotuvlar topilmadi."),
        ("Не нашёл товаров с остатком и нулевыми продажами.", "Qoldig‘i bor, lekin sotuvi yo‘q tovarlar topilmadi."),
        ("Расчёт примерный", "Hisob-kitob taxminiy"),
        ("хватит примерно на", "taxminan yetadi"),
        ("дней", "kun"),
        ("дня", "kun"),
        ("день", "kun"),
        ("шт.", "dona"),
        ("шт", "dona"),
        ("сум", "so‘m"),

        # invoices / excel
        ("Накладная:", "Yuk xati:"),
        ("Накладная №", "Yuk xati №"),
        ("Создана:", "Yaratilgan:"),
        ("Окно поставки:", "Yetkazib berish oynasi:"),
        ("К поставке:", "Yetkazishga:"),
        ("Принято:", "Qabul qilingan:"),
        ("Сумма:", "Summa:"),
        ("Состав:", "Tarkibi:"),
        ("По накладной:", "Yuk xati bo‘yicha:"),
        ("Расхождение:", "Farq:"),
        ("Закупочная цена:", "Xarid narxi:"),
        ("Excel-отчёт готов", "Excel hisobot tayyor"),
        ("Отчёт готов", "Hisobot tayyor"),

        # explanations / errors
                ("Finance API пока не вернул строки продаж за сегодня. Если в кабинете продажа уже есть, она может появиться здесь позже.",
         "Finance API bugungi sotuvlarni hali qaytarmadi. Agar kabinetda sotuv ko‘rinsa, bu yerda biroz keyin paydo bo‘lishi mumkin."),
        ("за выбранный период", "tanlangan davr uchun"),
        ("за сегодня", "bugun uchun"),
        ("за вчера", "kecha uchun"),
        ("за 7 дней", "7 kun uchun"),
        ("за 30 дней", "30 kun uchun"),
        ("за последние 30 дней", "oxirgi 30 kun uchun"),
        ("Ничего не найдено", "Hech narsa topilmadi"),
        ("Данных нет", "Ma’lumot yo‘q"),
        ("ошибка", "xatolik"),
        ("Ошибка", "Xatolik"),
        ("Попробуйте позже", "Keyinroq urinib ko‘ring"),
    ]
    for old, new in replacements:
        text = text.replace(old, new)
    return text


_ORIGINAL_MESSAGE_ANSWER = Message.answer
_ORIGINAL_BOT_SEND_MESSAGE = Bot.send_message


async def _answer_with_runtime_translation(self: Message, text: Any = None, *args: Any, **kwargs: Any) -> Any:
    try:
        telegram_id = self.from_user.id if self.from_user else None
        if isinstance(text, str) and get_user_language(telegram_id) == "uz":
            text = translate_runtime_text_to_uz(text)
    except Exception:
        pass
    return await _ORIGINAL_MESSAGE_ANSWER(self, text, *args, **kwargs)


async def _send_message_with_runtime_translation(self: Bot, chat_id: Any, text: Any, *args: Any, **kwargs: Any) -> Any:
    try:
        if isinstance(text, str) and get_user_language(int(chat_id)) == "uz":
            text = translate_runtime_text_to_uz(text)
    except Exception:
        pass
    return await _ORIGINAL_BOT_SEND_MESSAGE(self, chat_id, text, *args, **kwargs)


Message.answer = _answer_with_runtime_translation

Bot.send_message = _send_message_with_runtime_translation

# --- Чистка узбекского текста ---
# Первый переводчик выше специально не трогает бизнес-логику, а делает замену текста на лету.
# Здесь финальный слой: убирает смешанные русско-узбекские фразы вроде "Продажи за 30 kun".
_LEGACY_TRANSLATE_RUNTIME_TEXT_TO_UZ = translate_runtime_text_to_uz


def translate_runtime_text_to_uz(text: str) -> str:
    if not isinstance(text, str) or not text:
        return text
    text = _LEGACY_TRANSLATE_RUNTIME_TEXT_TO_UZ(text)

    fixes = [
        # Заголовки продаж — без смеси русского и узбекского
        ("💰 <b>Продажи сегодня</b>", "💰 <b>Bugungi savdo</b>"),
        ("💰 <b>Продажи bugun uchun</b>", "💰 <b>Bugungi savdo</b>"),
        ("💰 <b>Продажи за 7 kun</b>", "💰 <b>7 kunlik savdo</b>"),
        ("💰 <b>Продажи 7 kun uchun</b>", "💰 <b>7 kunlik savdo</b>"),
        ("💰 <b>Продажи за 30 kun</b>", "💰 <b>30 kunlik savdo</b>"),
        ("💰 <b>Продажи 30 kun uchun</b>", "💰 <b>30 kunlik savdo</b>"),
        ("💰 <b>Продажи tanlangan davr uchun</b>", "💰 <b>Tanlangan davr savdosi</b>"),
        ("💰 <b>Продажи kecha uchun</b>", "💰 <b>Kechagi savdo</b>"),
        ("💰 <b>Продажи за выбранный период</b>", "💰 <b>Tanlangan davr savdosi</b>"),

        # Частые поля в финансовых отчётах
        ("Проданных строк/позиций:", "Sotuv pozitsiyalari:"),
        ("Sotilgan строк/позitsiyalar:", "Sotuv pozitsiyalari:"),
        ("Sotilgan qator/pozitsiyalar:", "Sotuv pozitsiyalari:"),
        ("Позиций/строк:", "Sotuv pozitsiyalari:"),
        ("строк:", "qatorlar:"),
        ("Строк:", "Qatorlar:"),
        ("Штук:", "Soni:"),
        ("штук:", "soni:"),
        ("Средняя строка:", "O‘rtacha savdo:"),
        ("O‘rtacha строка:", "O‘rtacha savdo:"),
        ("Отменённых строк:", "Bekor qilinganlar:"),
        ("Отмененных строк:", "Bekor qilinganlar:"),
        ("Bekor qilingan qatorlar:", "Bekor qilinganlar:"),
        ("Топ товаров по so‘mme:", "Summa bo‘yicha top tovarlar:"),
        ("Топ товаров по so‘mме:", "Summa bo‘yicha top tovarlar:"),
        ("Топ товаров по сумме:", "Summa bo‘yicha top tovarlar:"),
        ("Юнит-экономика", "Unit iqtisodiyot"),
        ("Себестоимость", "Tannarx"),
        ("Прибыль", "Foyda"),
        ("Маржа", "Marja"),
        ("Top товаров по so‘mme:", "Summa bo‘yicha top tovarlar:"),
        ("Top товаров по сумме:", "Summa bo‘yicha top tovarlar:"),
        ("Top tovarlar по so‘mme:", "Summa bo‘yicha top tovarlar:"),
        ("Top tovarlar по сумме:", "Summa bo‘yicha top tovarlar:"),
        ("по so‘mme", "summa bo‘yicha"),
        ("по so‘mме", "summa bo‘yicha"),
        ("по сумме", "summa bo‘yicha"),
        ("Без названия", "Nomsiz"),

        # Периоды и служебные фразы
        ("за 30 kun", "30 kun uchun"),
        ("за 7 kun", "7 kun uchun"),
        ("за 1 kun", "1 kun uchun"),
        ("Сегодня", "Bugun"),
        ("Вчера", "Kecha"),
        ("7 дней", "7 kun"),
        ("30 дней", "30 kun"),
        ("Ответ Finance API пришёл, но строки продаж не найдены.", "Finance API javob berdi, lekin savdo qatorlari topilmadi."),
        ("Фрагмент ответа:", "Javobdan parcha:"),
        ("Подробно:", "Batafsil:"),
        ("Показаны первые", "Birinchi"),
        ("позиций из", "pozitsiya ko‘rsatildi, jami"),

        # Остатки, накладные, общие поля
        ("Товары без продаж", "Sotilmayotgan tovarlar"),
        ("Не продаётся", "Sotilmayapti"),
        ("Прогноз остатков", "Qoldiq prognozi"),
        ("Все магазины", "Barcha do‘konlar"),
        ("Накладные FBO", "FBO yuk xatlari"),
        ("Состав накладной", "Yuk xati tarkibi"),
        ("Потерянные", "Yo‘qolganlar"),
        ("Заканчивается", "Tugayapti"),
        ("Заканчиваются", "Tugayapti"),
        ("Остатки", "Qoldiq"),
        ("Остаток", "Qoldiq"),
        ("Продано", "Sotilgan"),
        ("Возврат", "Qaytarilgan"),
        ("Возвраты", "Qaytarilganlar"),
        ("Комиссия Uzum", "Uzum komissiyasi"),
        ("Комиссия", "Komissiya"),
        ("Логистика", "Logistika"),
        ("Выручка", "Tushum"),
        ("К выплате всего", "Jami to‘lovga"),
        ("К выплате", "To‘lovga"),
        ("Уже выведено", "Chiqarilgan"),
        ("Остаток к выплате", "To‘lovga qoldi"),
        ("Статусы", "Statuslar"),
        ("Статус", "Status"),
        ("Магазин", "Do‘kon"),
        ("Товар", "Tovar"),
        ("Дата", "Sana"),
        ("Цена продажи", "Sotuv narxi"),
        ("ID заказа", "Buyurtma ID"),
        ("ID продажи", "Sotuv ID"),
        ("Кол-во", "Soni"),
        ("Итого", "Jami"),
        ("Сумма", "Summa"),
        ("Разница", "Farq"),
        ("Расхождение", "Farq"),
        ("Принято", "Qabul qilingan"),
        ("К поставке", "Yetkazishga"),
        ("Накладная", "Yuk xati"),
    ]
    for old, new in fixes:
        text = text.replace(old, new)

    # Безопасная замена валюты/единиц: только как отдельные слова в суммах.
    import re
    text = re.sub(r"(?<=\d)\s*сум\b", " so‘m", text)
    text = re.sub(r"(?<=\d)\s*шт\.?\b", " dona", text)
    return text



# --- Дополнительная чистка узбекского перевода: магазины и уведомления ---
_CLEAN3_TRANSLATE_RUNTIME_TEXT_TO_UZ = translate_runtime_text_to_uz


def translate_runtime_text_to_uz(text: str) -> str:
    if not isinstance(text, str) or not text:
        return text
    text = _CLEAN3_TRANSLATE_RUNTIME_TEXT_TO_UZ(text)

    fixes = [
        # Магазины
        ("🏪 <b>Ваши do‘konlari:</b>", "🏪 <b>Do‘konlaringiz:</b>"),
        ("🏪 <b>Ваши do‘konlar:</b>", "🏪 <b>Do‘konlaringiz:</b>"),
        ("🏪 <b>Ваши магазины:</b>", "🏪 <b>Do‘konlaringiz:</b>"),
        ("Текущий основной Do‘kon:", "Joriy asosiy do‘kon:"),
        ("Текущий основной магазин:", "Joriy asosiy do‘kon:"),
        ("Joriy основной do‘kon:", "Joriy asosiy do‘kon:"),
        ("Чтобы выбрать:", "Tanlash uchun:"),
        ("не выбран", "tanlanmagan"),

        # Уведомления: заголовки и статусы
        ("💸 <b>Уведомления о новых продажах</b>", "💸 <b>Yangi savdolar xabarnomalari</b>"),
        ("💸 <b>Yangi продажа xabarnomalari</b>", "💸 <b>Yangi savdolar xabarnomalari</b>"),
        ("🔔 <b>Уведомления</b>", "🔔 <b>Xabarnomalar</b>"),
        ("🔔 <b>Уведомления:</b>", "🔔 <b>Xabarnomalar:</b>"),
        ("Уведомления о новых продажах", "Yangi savdolar xabarnomalari"),
        ("Holat: ✅ включены", "Holat: ✅ yoqilgan"),
        ("Holat: ❌ выключены", "Holat: ❌ o‘chirilgan"),
        ("Status: ✅ включены", "Holat: ✅ yoqilgan"),
        ("Status: ❌ выключены", "Holat: ❌ o‘chirilgan"),
        ("Статус: ✅ включены", "Holat: ✅ yoqilgan"),
        ("Статус: ❌ выключены", "Holat: ❌ o‘chirilgan"),
        ("Проверка каждые:", "Tekshiruv har:"),
        ("Tekshiruv har: <b>", "Tekshiruv har <b>"),
        ("сек.", "soniya"),
        ("Состояние: продажи уже запомнены", "Holat: savdolar allaqachon eslab qolingan"),
        ("Состояние: инициализация при следующей проверке", "Holat: keyingi tekshiruvda ishga tushadi"),
        ("Holat: продажи уже запомнены", "Holat: savdolar allaqachon eslab qolingan"),
        ("Holat: инициализация при следующей проверке", "Holat: keyingi tekshiruvda ishga tushadi"),
        ("Бот смотрит Finance API bugun uchun.", "Bot bugungi savdolarni Finance API orqali tekshiradi."),
        ("Бот смотрит Finance API за сегодня.", "Bot bugungi savdolarni Finance API orqali tekshiradi."),
        ("Если Finance API отдаёт продажу с задержкой, уведомление тоже придёт с задержкой.", "Agar Finance API savdoni kechikib bersa, xabarnoma ham kechikib keladi."),
        ("Bot bugungi savdolarni Finance API orqali tekshiradi. Agar Finance API savdoni kechikib bersa, xabarnoma ham kechikib keladi.", "Bot bugungi savdolarni Finance API orqali tekshiradi. Agar savdo kechikib ko‘rinsa, xabarnoma ham biroz kechikib kelishi mumkin."),
        ("Do‘kon:", "Do‘kon:"),
    ]
    for old, new in fixes:
        text = text.replace(old, new)
    return text

def language_markup() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="🇷🇺 Русский", callback_data="set_lang:ru"),
                InlineKeyboardButton(text="🇺🇿 O‘zbekcha", callback_data="set_lang:uz"),
            ]
        ]
    )


MAIN_MENU_RU = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="🔌 Подключить"), KeyboardButton(text="🎥 Видеоинструкция")],
        [KeyboardButton(text="💎 Подписка"), KeyboardButton(text="🌐 Язык")],
        [KeyboardButton(text="ℹ️ Помощь")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Сначала подключите магазин",
)

MAIN_MENU_RU_ADMIN = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="🔌 Подключить"), KeyboardButton(text="🎥 Видеоинструкция")],
        [KeyboardButton(text="💎 Подписка"), KeyboardButton(text="🌐 Язык")],
        [KeyboardButton(text="ℹ️ Помощь"), KeyboardButton(text="👑 Админ")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Сначала подключите магазин",
)

MAIN_MENU_UZ = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="🔌 Ulash"), KeyboardButton(text="🎥 API ulash videosi")],
        [KeyboardButton(text="💎 Obuna"), KeyboardButton(text="🌐 Til")],
        [KeyboardButton(text="ℹ️ Yordam")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Avval do‘konni ulang",
)

MAIN_MENU_UZ_ADMIN = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="🔌 Ulash"), KeyboardButton(text="🎥 API ulash videosi")],
        [KeyboardButton(text="💎 Obuna"), KeyboardButton(text="🌐 Til")],
        [KeyboardButton(text="ℹ️ Yordam"), KeyboardButton(text="👑 Admin")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Avval do‘konni ulang",
)

# Главное меню после подключения API: простая структура по разделам.
MAIN_MENU_RU_CONNECTED = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="💰 Продажи"), KeyboardButton(text="📦 Склад")],
        [KeyboardButton(text="🧠 Что проверить"), KeyboardButton(text="🔔 Уведомления")],
        [KeyboardButton(text="📊 Отчёты"), KeyboardButton(text="🏪 Магазины")],
        [KeyboardButton(text="💎 Подписка"), KeyboardButton(text="🌐 Язык")],
        [KeyboardButton(text="ℹ️ Помощь")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Выберите раздел",
)

MAIN_MENU_RU_CONNECTED_ADMIN = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="💰 Продажи"), KeyboardButton(text="📦 Склад")],
        [KeyboardButton(text="🧠 Что проверить"), KeyboardButton(text="🔔 Уведомления")],
        [KeyboardButton(text="📊 Отчёты"), KeyboardButton(text="🏪 Магазины")],
        [KeyboardButton(text="💎 Подписка"), KeyboardButton(text="🌐 Язык")],
        [KeyboardButton(text="ℹ️ Помощь"), KeyboardButton(text="👑 Админ")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Выберите раздел",
)

MAIN_MENU_UZ_CONNECTED = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="💰 Savdo"), KeyboardButton(text="📦 Ombor")],
        [KeyboardButton(text="🧠 Tekshirish"), KeyboardButton(text="🔔 Xabarnomalar")],
        [KeyboardButton(text="📊 Hisobotlar"), KeyboardButton(text="🏪 Do‘konlar")],
        [KeyboardButton(text="💎 Obuna"), KeyboardButton(text="🌐 Til")],
        [KeyboardButton(text="ℹ️ Yordam")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Bo‘limni tanlang",
)

MAIN_MENU_UZ_CONNECTED_ADMIN = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="💰 Savdo"), KeyboardButton(text="📦 Ombor")],
        [KeyboardButton(text="🧠 Tekshirish"), KeyboardButton(text="🔔 Xabarnomalar")],
        [KeyboardButton(text="📊 Hisobotlar"), KeyboardButton(text="🏪 Do‘konlar")],
        [KeyboardButton(text="💎 Obuna"), KeyboardButton(text="🌐 Til")],
        [KeyboardButton(text="ℹ️ Yordam"), KeyboardButton(text="👑 Admin")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Bo‘limni tanlang",
)

SALES_MENU_RU = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📊 Сегодня"), KeyboardButton(text="📆 Вчера")],
        [KeyboardButton(text="🗓 7 дней"), KeyboardButton(text="📅 30 дней")],
        [KeyboardButton(text="💰 Баланс"), KeyboardButton(text="🌐 Все магазины")],
        [KeyboardButton(text="🏆 Топ товаров"), KeyboardButton(text="🐢 Не продаётся")],
        [KeyboardButton(text="🧾 Юнит-экономика"), KeyboardButton(text="💰 Прибыль")],
        [KeyboardButton(text="📥 Себестоимость Excel")],
        [KeyboardButton(text="⬅️ Главное меню")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Продажи",
)

SALES_MENU_UZ = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📊 Bugun"), KeyboardButton(text="📆 Kecha")],
        [KeyboardButton(text="🗓 7 kun"), KeyboardButton(text="📅 30 kun")],
        [KeyboardButton(text="💰 Balans"), KeyboardButton(text="🌐 Barcha do‘konlar")],
        [KeyboardButton(text="🏆 Top tovarlar"), KeyboardButton(text="🐢 Sotilmayapti")],
        [KeyboardButton(text="🧾 Unit iqtisodiyot"), KeyboardButton(text="💰 Foyda")],
        [KeyboardButton(text="📥 Tannarx Excel")],
        [KeyboardButton(text="⬅️ Asosiy menyu")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Savdo",
)

STOCK_MENU_RU = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📦 Остатки"), KeyboardButton(text="⚠️ Прогноз остатков")],
        [KeyboardButton(text="🧭 Потерянные"), KeyboardButton(text="📄 Накладные FBO")],
        [KeyboardButton(text="⬅️ Главное меню")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Склад",
)

STOCK_MENU_UZ = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📦 Qoldiq"), KeyboardButton(text="⚠️ Qoldiq prognozi")],
        [KeyboardButton(text="🧭 Yo‘qolganlar"), KeyboardButton(text="📄 FBO yuk xatlari")],
        [KeyboardButton(text="⬅️ Asosiy menyu")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Ombor",
)

NOTIFY_MENU_RU = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="💸 Новые продажи"), KeyboardButton(text="📉 Низкие остатки")],
        [KeyboardButton(text="❌ Нет в наличии"), KeyboardButton(text="⚙️ Статус")],
        [KeyboardButton(text="⬅️ Главное меню")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Уведомления",
)

NOTIFY_MENU_UZ = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="💸 Yangi savdolar"), KeyboardButton(text="📉 Kam qoldiq")],
        [KeyboardButton(text="❌ Qoldiq tugagan"), KeyboardButton(text="⚙️ Holat")],
        [KeyboardButton(text="⬅️ Asosiy menyu")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Xabarnomalar",
)

REPORT_MENU_RU = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📊 Excel отчёт"), KeyboardButton(text="🌙 Утренний отчёт")],
        [KeyboardButton(text="💰 Прибыль"), KeyboardButton(text="📥 Себестоимость Excel")],
        [KeyboardButton(text="✅ Проверить подключение"), KeyboardButton(text="🔐 Безопасность")],
        [KeyboardButton(text="⬅️ Главное меню")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Отчёты",
)

REPORT_MENU_UZ = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📊 Excel hisobot"), KeyboardButton(text="🌙 Ertalabki hisobot")],
        [KeyboardButton(text="💰 Foyda"), KeyboardButton(text="📥 Tannarx Excel")],
        [KeyboardButton(text="✅ Ulanishni tekshirish"), KeyboardButton(text="🔐 Xavfsizlik")],
        [KeyboardButton(text="⬅️ Asosiy menyu")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Hisobotlar",
)

ATTENTION_MENU_RU = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="🔍 Проверить сейчас")],
        [KeyboardButton(text="⚠️ Остатки"), KeyboardButton(text="🐢 Без продаж")],
        [KeyboardButton(text="🧾 Нет себестоимости"), KeyboardButton(text="📉 Низкая прибыль")],
        [KeyboardButton(text="❌ Отмены"), KeyboardButton(text="⬅️ Главное меню")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Что проверить",
)

ATTENTION_MENU_UZ = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="🔍 Hozir tekshirish")],
        [KeyboardButton(text="⚠️ Qoldiqlar"), KeyboardButton(text="🐢 Sotuv yo‘q")],
        [KeyboardButton(text="🧾 Tannarx yo‘q"), KeyboardButton(text="📉 Past foyda")],
        [KeyboardButton(text="❌ Bekor qilishlar"), KeyboardButton(text="⬅️ Asosiy menyu")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Nimani tekshiramiz",
)

ADMIN_PANEL_MENU_RU = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="👥 Пользователи"), KeyboardButton(text="💳 Оплаты")],
        [KeyboardButton(text="⏳ Скоро заканчиваются"), KeyboardButton(text="⛔ Заблокированные")],
        [KeyboardButton(text="✅ Проверить подключение"), KeyboardButton(text="🎥 Видеоинструкция")],
        [KeyboardButton(text="📦 Бэкап базы")],
        [KeyboardButton(text="📢 Рассылка"), KeyboardButton(text="⬅️ Главное меню")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Админ-панель",
)

ADMIN_PANEL_MENU_UZ = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="👥 Foydalanuvchilar"), KeyboardButton(text="💳 To‘lovlar")],
        [KeyboardButton(text="⏳ Tugayotganlar"), KeyboardButton(text="⛔ Bloklanganlar")],
        [KeyboardButton(text="✅ Ulanishni tekshirish"), KeyboardButton(text="🎥 API ulash videosi")],
        [KeyboardButton(text="📦 Baza zaxirasi")],
        [KeyboardButton(text="📢 Xabar yuborish"), KeyboardButton(text="⬅️ Asosiy menyu")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Admin panel",
)

# Для совместимости: если где-то осталась статичная разметка, будет русский вариант.
MAIN_MENU = MAIN_MENU_RU
ADMIN_PANEL_MENU = ADMIN_PANEL_MENU_RU
ANALYTICS_MENU = MAIN_MENU_RU
PRODUCTS_MENU = MAIN_MENU_RU
ORDERS_MENU = MAIN_MENU_RU
NOTIFICATIONS_MENU = MAIN_MENU_RU
SETTINGS_MENU = MAIN_MENU_RU


def _user_has_uzum_connection(telegram_id: int | None) -> bool:
    if not telegram_id:
        return False
    try:
        if hasattr(db, "has_uzum_connection"):
            return bool(db.has_uzum_connection(int(telegram_id)))
        user = db.get_user(int(telegram_id))
        return bool(user and user["uzum_token_encrypted"])
    except Exception:
        return False


def main_menu_for_user(telegram_id: int | None) -> ReplyKeyboardMarkup:
    lang = get_user_language(telegram_id)
    admin = is_admin(telegram_id)
    connected = _user_has_uzum_connection(telegram_id)

    if lang == "uz":
        if connected:
            return MAIN_MENU_UZ_CONNECTED_ADMIN if admin else MAIN_MENU_UZ_CONNECTED
        return MAIN_MENU_UZ_ADMIN if admin else MAIN_MENU_UZ

    if connected:
        return MAIN_MENU_RU_CONNECTED_ADMIN if admin else MAIN_MENU_RU_CONNECTED
    return MAIN_MENU_RU_ADMIN if admin else MAIN_MENU_RU


def sales_menu_for_user(telegram_id: int | None) -> ReplyKeyboardMarkup:
    return SALES_MENU_UZ if get_user_language(telegram_id) == "uz" else SALES_MENU_RU


def stock_menu_for_user(telegram_id: int | None) -> ReplyKeyboardMarkup:
    return STOCK_MENU_UZ if get_user_language(telegram_id) == "uz" else STOCK_MENU_RU


def notify_menu_for_user(telegram_id: int | None) -> ReplyKeyboardMarkup:
    return NOTIFY_MENU_UZ if get_user_language(telegram_id) == "uz" else NOTIFY_MENU_RU


def report_menu_for_user(telegram_id: int | None) -> ReplyKeyboardMarkup:
    return REPORT_MENU_UZ if get_user_language(telegram_id) == "uz" else REPORT_MENU_RU


def attention_menu_for_user(telegram_id: int | None) -> ReplyKeyboardMarkup:
    return ATTENTION_MENU_UZ if get_user_language(telegram_id) == "uz" else ATTENTION_MENU_RU


def _message_user_id(message: Message) -> int | None:
    try:
        return message.from_user.id if message.from_user else None
    except Exception:
        return None


def sales_menu_for_message(message: Message) -> ReplyKeyboardMarkup:
    return sales_menu_for_user(_message_user_id(message))


def stock_menu_for_message(message: Message) -> ReplyKeyboardMarkup:
    return stock_menu_for_user(_message_user_id(message))


def notify_menu_for_message(message: Message) -> ReplyKeyboardMarkup:
    return notify_menu_for_user(_message_user_id(message))


def report_menu_for_message(message: Message) -> ReplyKeyboardMarkup:
    return report_menu_for_user(_message_user_id(message))


def attention_menu_for_message(message: Message) -> ReplyKeyboardMarkup:
    return attention_menu_for_user(_message_user_id(message))


def menu_for_message(message: Message) -> ReplyKeyboardMarkup:
    try:
        telegram_id = message.from_user.id if message.from_user else None
    except Exception:
        telegram_id = None
    return main_menu_for_user(telegram_id)


def admin_menu_for_user(telegram_id: int | None) -> ReplyKeyboardMarkup:
    return ADMIN_PANEL_MENU_UZ if get_user_language(telegram_id) == "uz" else ADMIN_PANEL_MENU_RU


def admin_menu_for_message(message: Message) -> ReplyKeyboardMarkup:
    try:
        telegram_id = message.from_user.id if message.from_user else None
    except Exception:
        telegram_id = None
    return admin_menu_for_user(telegram_id)


# Переопределяем тексты подписки после инициализации языка, чтобы /my_subscription был на выбранном языке.
def subscription_status_text(telegram_id: int) -> str:
    row = ensure_subscription(telegram_id)
    lang = get_user_language(telegram_id)
    if is_admin(telegram_id):
        return "👑 Admin kirish: cheklovsiz" if lang == "uz" else "👑 Админ-доступ: без ограничений"
    if int(row.get("blocked") or 0) == 1:
        return "⛔ Foydalanuvchi bloklangan" if lang == "uz" else "⛔ Пользователь заблокирован"
    now = _utc_now()
    trial_until = _dt_from_db(row.get("trial_until"))
    paid_until = _dt_from_db(row.get("subscription_until"))
    until = subscription_active_until(row)
    if until and until > now:
        if paid_until and paid_until == until:
            return (f"✅ Obuna {_fmt_dt(paid_until)} gacha faol" if lang == "uz" else f"✅ Подписка активна до {_fmt_dt(paid_until)}")
        return (f"🎁 Trial {_fmt_dt(trial_until)} gacha faol" if lang == "uz" else f"🎁 Trial активен до {_fmt_dt(trial_until)}")
    return "⛔ Obuna muddati tugagan" if lang == "uz" else "⛔ Подписка закончилась"


def subscription_full_text(telegram_id: int) -> str:
    row = ensure_subscription(telegram_id)
    status = subscription_status_text(telegram_id)
    lang = get_user_language(telegram_id)

    if is_admin(telegram_id):
        if lang == "uz":
            return (
                "💎 <b>Mening obunam</b>\n\n"
                f"Telegram ID: <code>{telegram_id}</code>\n"
                f"Holat: {status}\n\n"
                "Admin uchun trial va to‘lov sanasi muhim emas — kirish doim ochiq.\n\n"
                "Admin buyruqlari:\n"
                "• <code>/admin</code> — admin panel\n"
                "• <code>/users</code> — foydalanuvchilar\n"
                "• <code>/extend ID 30</code> — kirishni uzaytirish\n"
                "• <code>/block ID</code> — bloklash\n"
                "• <code>/unblock ID</code> — blokdan chiqarish\n"
                "• <code>/paid ID summa kun</code> — to‘lovni yozish\n"
                "• <code>/payments</code> — to‘lovlar tarixi\n"
                "• <code>/backup_db</code> — baza zaxirasi"
            )
        return (
            "💎 <b>Моя подписка</b>\n\n"
            f"Telegram ID: <code>{telegram_id}</code>\n"
            f"Статус: {status}\n\n"
            "Trial и дата оплаты для администратора не важны — доступ всегда открыт.\n\n"
            "Команды администратора:\n"
            "• <code>/admin</code> — админ-панель\n"
            "• <code>/users</code> — пользователи\n"
            "• <code>/extend ID 30</code> — продлить доступ\n"
            "• <code>/block ID</code> — заблокировать\n"
            "• <code>/unblock ID</code> — разблокировать\n"
            "• <code>/paid ID сумма дни</code> — записать оплату\n"
            "• <code>/payments</code> — история оплат\n"
            "• <code>/backup_db</code> — скачать базу"
        )

    if lang == "uz":
        return (
            "💎 <b>Mening obunam</b>\n\n"
            f"Telegram ID: <code>{telegram_id}</code>\n"
            f"Holat: {status}\n"
            f"Trial tugash vaqti: <b>{_fmt_dt(row.get('trial_until'))}</b>\n"
            f"To‘langan muddat: <b>{_fmt_dt(row.get('subscription_until'))}</b>\n\n"
            "Tariflar:\n"
            f"<b>{escape(SUBSCRIPTION_PLANS_TEXT)}</b>\n\n"
            f"{escape(PAYMENT_TEXT)}\n\n"
            "To‘lovlar tarixi: <code>/my_payments</code>\n"
            "Yordam: <code>/support</code>\n"
            "API-kalitni almashtirish: <code>/reconnect</code>\n"
            "API-kalitni o‘chirish: <code>/disconnect</code>"
        )
    return (
        "💎 <b>Моя подписка</b>\n\n"
        f"Telegram ID: <code>{telegram_id}</code>\n"
        f"Статус: {status}\n"
        f"Trial до: <b>{_fmt_dt(row.get('trial_until'))}</b>\n"
        f"Оплачено до: <b>{_fmt_dt(row.get('subscription_until'))}</b>\n\n"
        "Тарифы:\n"
        f"<b>{escape(SUBSCRIPTION_PLANS_TEXT)}</b>\n\n"
        f"{escape(PAYMENT_TEXT)}\n\n"
        "История оплат: <code>/my_payments</code>\n"
        "Поддержка: <code>/support</code>\n"
        "Заменить API-ключ: <code>/reconnect</code>\n"
        "Удалить API-ключ: <code>/disconnect</code>"
    )

class ConnectStates(StatesGroup):
    waiting_for_token = State()
    waiting_for_shop_id = State()


class CostImportStates(StatesGroup):
    waiting_for_file = State()


def get_tg_id(message: Message) -> int:
    if not message.from_user:
        raise RuntimeError("Unknown Telegram user")
    return message.from_user.id


def upsert_from_message(message: Message) -> int:
    user = message.from_user
    if not user:
        raise RuntimeError("Unknown Telegram user")
    db.upsert_user(user.id, user.username, user.first_name)
    return user.id


def get_uzum_for_user(telegram_id: int) -> UzumClient | None:
    encrypted = db.get_encrypted_token(telegram_id)
    if not encrypted:
        return None
    token = cipher.decrypt(encrypted)
    return UzumClient(token, UZUM_API_BASE_URL)


def get_staff_uzum_client() -> UzumClient | None:
    if not STAFF_CONNECT_ENABLED or not STAFF_UZUM_TOKEN:
        return None
    return UzumClient(STAFF_UZUM_TOKEN, UZUM_API_BASE_URL)


def _shop_id_from_obj(shop: Any) -> int | None:
    if not isinstance(shop, dict):
        return None
    for key in ("id", "shopId", "shop_id", "storeId"):
        value = shop.get(key)
        if value is not None:
            try:
                return int(value)
            except Exception:
                pass
    return None


def _find_shop_in_list(shops: list[Any], shop_id: int) -> dict[str, Any] | None:
    for item in shops:
        if isinstance(item, dict) and _shop_id_from_obj(item) == int(shop_id):
            return item
    return None


def _fallback_shop_obj(shop_id: int) -> dict[str, Any]:
    return {
        "id": int(shop_id),
        "shopId": int(shop_id),
        "title": f"Магазин {int(shop_id)}",
        "name": f"Магазин {int(shop_id)}",
    }


async def notify_admins_staff_shop_connected(message: Message, telegram_id: int, shop_id: int) -> None:
    user = message.from_user
    username = f"@{user.username}" if user and user.username else "—"
    first_name = user.first_name if user and user.first_name else "—"
    text = (
        "🆕 <b>Магазин подключён через сотрудника</b>\n\n"
        f"Пользователь: {escape(first_name)} {escape(username)}\n"
        f"Telegram ID: <code>{telegram_id}</code>\n"
        f"Shop ID: <code>{shop_id}</code>\n"
        f"Доступ: {subscription_status_text(telegram_id)}"
    )
    for admin_id in ADMIN_IDS:
        if int(admin_id) == int(telegram_id):
            continue
        try:
            await bot.send_message(int(admin_id), text, reply_markup=admin_menu_for_user(int(admin_id)))
            await asyncio.sleep(0.1)
        except Exception:
            logging.exception("Failed to notify admin about staff shop connection")


async def connect_shop_by_staff(message: Message, shop_id_text: str, state: FSMContext | None = None) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    shop_id_text = (shop_id_text or "").strip()
    if not shop_id_text.isdigit():
        if lang == "uz":
            await message.answer(
                "Shop ID faqat raqamlardan iborat bo‘lishi kerak.\nMasalan: <code>113982</code>",
                reply_markup=menu_for_message(message),
            )
        else:
            await message.answer(
                "ID магазина должен состоять только из цифр.\nНапример: <code>113982</code>",
                reply_markup=menu_for_message(message),
            )
        return

    if not await require_active_subscription(message, telegram_id):
        return

    staff_client = get_staff_uzum_client()
    if staff_client is None:
        await message.answer(
            "⚠️ Простой способ подключения пока не настроен.\n\n"
            "Можно подключиться старым способом через API-ключ: <code>/connect</code>",
            reply_markup=menu_for_message(message),
        )
        if state:
            await state.clear()
        return

    shop_id = int(shop_id_text)
    save_staff_shop_status(telegram_id, shop_id, "pending")

    try:
        shops: list[Any] = []
        shop_obj: dict[str, Any] | None = None
        try:
            shops_data = await staff_client.get_shops()
            shops = extract_items(shops_data)
            shop_obj = _find_shop_in_list(shops, shop_id)
        except Exception:
            logging.exception("Staff connect: failed to load shops list")

        await staff_client.get_products(shop_id, page=0, size=1)

        # Проверяем доступ к Finance. Даже если продаж нет, метод должен ответить без 403.
        date_from, date_to = _days_range_ms(30)
        await _load_finance_orders(
            staff_client,
            shop_id,
            date_from_ms=date_from,
            date_to_ms=date_to,
            max_pages=1,
            page_size=1,
        )

        encrypted = cipher.encrypt(STAFF_UZUM_TOKEN)
        db.save_connection(telegram_id, encrypted, [shop_obj or _fallback_shop_obj(shop_id)])
        try:
            db.set_default_shop_id(telegram_id, shop_id)
        except Exception:
            pass
        save_staff_shop_status(telegram_id, shop_id, "connected")

        if state:
            await state.clear()

        if lang == "uz":
            text_ok = (
                "✅ <b>Do‘kon ulandi</b>\n\n"
                f"Shop ID: <code>{shop_id}</code>\n"
                "Xodim orqali kirish tasdiqlandi.\n\n"
                "Endi savdolar, qoldiqlar va hisobotlardan foydalanishingiz mumkin."
            )
        else:
            text_ok = (
                "✅ <b>Магазин подключён</b>\n\n"
                f"Shop ID: <code>{shop_id}</code>\n"
                "Доступ через сотрудника подтверждён.\n\n"
                "Теперь можно смотреть продажи, остатки и отчёты."
            )
        await message.answer(text_ok, reply_markup=menu_for_message(message))
        await notify_admins_staff_shop_connected(message, telegram_id, shop_id)

    except Exception as e:
        save_staff_shop_status(telegram_id, shop_id, "no_access", str(e))
        raw = str(e)
        low = raw.lower()
        if "403" in raw or "rbac" in low or "forbidden" in low or "access" in low:
            if lang == "uz":
                text = (
                    "⛔ <b>Do‘konga kirish topilmadi</b>\n\n"
                    "Ehtimol, xodim hali qo‘shilmagan yoki unga savdo/moliya/tovarlar bo‘yicha huquqlar berilmagan.\n\n"
                    "1. Uzum Seller kabinetida xodim qo‘shilganini tekshiring.\n"
                    "2. Savdo, moliya, tovarlar va qoldiq huquqlarini bering.\n"
                    "3. Keyin Shop ID ni qayta yuboring."
                )
            else:
                text = (
                    "⛔ <b>Доступ к магазину не найден</b>\n\n"
                    "Скорее всего, сотрудник ещё не добавлен или ему не дали права на продажи/финансы/товары.\n\n"
                    "1. Проверьте, что сотрудник добавлен в Uzum Seller.\n"
                    "2. Дайте права на продажи, финансы, товары и остатки.\n"
                    "3. После этого отправьте Shop ID ещё раз."
                )
            await message.answer(text, reply_markup=menu_for_message(message))
            return
        await send_api_error(message, e)


async def require_connection(message: Message) -> tuple[int, UzumClient, int] | None:
    telegram_id = upsert_from_message(message)
    if not await require_active_subscription(message, telegram_id):
        return None
    client = get_uzum_for_user(telegram_id)
    shop_id = db.get_default_shop_id(telegram_id)

    if client is None:
        lang = get_user_language(telegram_id)
        if lang == "uz":
            text = (
                "Avval do‘konni ulang.\n\n"
                "<code>/connect</code> buyrug‘ini bosing va Uzum Seller API-kalitingizni yuboring.\n"
                "Videoqo‘llanma: <code>/video</code>"
            )
        else:
            text = (
                "Сначала подключите магазин.\n\n"
                "Нажмите <code>/connect</code> и отправьте API-ключ из кабинета Uzum Seller.\n"
                "Видеоинструкция: <code>/video</code>"
            )
        await message.answer(text, reply_markup=menu_for_message(message))
        return None

    if shop_id is None:
        await message.answer(
            "Токен подключён, но основной магазин не выбран.\n"
            "Напишите <code>/shops</code>, потом <code>/setshop SHOP_ID</code>.",
            reply_markup=menu_for_message(message),
        )
        return None

    return telegram_id, client, int(shop_id)


async def send_api_error(message: Message, error: Exception) -> None:
    raw = str(error)
    low = raw.lower()
    if "401" in raw or "unauthorized" in low:
        user_text = (
            "🔐 <b>Uzum API-ключ не принят</b>\n\n"
            "Возможно, ключ неверный, удалён или истёк.\n"
            "Создайте новый ключ в кабинете Uzum Seller и подключите его через <code>/reconnect</code>."
        )
    elif "403" in raw or "rbac" in low or "forbidden" in low:
        user_text = (
            "⛔ <b>Нет доступа к этому методу Uzum API</b>\n\n"
            "Проверьте права API-ключа в кабинете Uzum Seller.\n"
            "Иногда отдельные методы недоступны со стороны Uzum для конкретного магазина."
        )
    elif "429" in raw or "too many" in low:
        user_text = (
            "⏳ <b>Uzum временно ограничил запросы</b>\n\n"
            "Слишком много запросов к Uzum API. Подождите несколько минут и попробуйте снова."
        )
    elif "500" in raw or "502" in raw or "503" in raw or "504" in raw:
        user_text = (
            "⚠️ <b>Uzum API временно недоступен</b>\n\n"
            "Это похоже на ошибку на стороне Uzum. Попробуйте позже."
        )
    else:
        text = escape(raw)
        if len(text) > 1200:
            text = text[:1200] + "\n..."
        user_text = f"⚠️ <b>Ошибка API</b>\n<code>{text}</code>"
    await message.answer(user_text, reply_markup=menu_for_message(message))


def parse_args(text: str) -> str:
    text = (text or "").strip()
    # Аргументы берём только у настоящих slash-команд.
    # Если пользователь нажал русскую кнопку вроде "📦 Товары",
    # это не должно уходить в Uzum API как поисковый запрос или статус заказа.
    if not text.startswith("/"):
        return ""
    parts = text.split(maxsplit=1)
    return parts[1].strip() if len(parts) > 1 else ""


async def load_products(
    client: UzumClient,
    shop_id: int,
    *,
    search_query: str = "",
    max_pages: int = 20,
    page_size: int = 100,
) -> list[Any]:
    all_products: list[Any] = []
    for page in range(max_pages):
        data = await client.get_products(
            shop_id, search_query=search_query, page=page, size=page_size
        )
        items = extract_items(data)
        if not items:
            break
        all_products.extend(items)
        if len(items) < page_size:
            break
    return all_products


async def load_sku_rows(
    client: UzumClient,
    shop_id: int,
    *,
    search_query: str = "",
    max_pages: int = 20,
) -> list[dict[str, Any]]:
    products = await load_products(
        client, shop_id, search_query=search_query, max_pages=max_pages, page_size=100
    )
    return flatten_sku_rows(products)


async def connect_token(
    message: Message, token: str, state: FSMContext | None = None
) -> None:
    telegram_id = upsert_from_message(message)
    token = token.strip()
    if not token or len(token) < 20:
        await message.answer(
            "Похоже, это не Uzum API-токен.\n"
            "Отправьте полный токен или нажмите /cancel.",
            reply_markup=menu_for_message(message),
        )
        return

    try:
        client = UzumClient(token, UZUM_API_BASE_URL)
        data = await client.get_shops()
        shops = extract_items(data)
        if not shops:
            await message.answer(
                "Токен сработал, но список магазинов не найден.\n"
                "Ответ API:\n<code>"
                + escape(compact_json_preview(data))
                + "</code>",
                reply_markup=menu_for_message(message),
            )
            return

        encrypted = cipher.encrypt(token)
        default_shop_id = db.save_connection(telegram_id, encrypted, shops)

        try:
            await message.delete()
        except Exception:
            pass

        lang = get_user_language(telegram_id)
        lines = [format_shop_line(shop) for shop in shops[:20]]
        if lang == "uz":
            text_ok = (
                "✅ <b>Do‘kon ulandi</b>\n\n"
                f"Topilgan do‘konlar: <b>{len(shops)}</b>\n"
                + "\n".join(lines)
                + "\n\nFaol do‘kon: "
                + (f"<code>{default_shop_id}</code>" if default_shop_id else "tanlanmagan")
                + "\n\nEndi asosiy bo‘limlardan foydalanishingiz mumkin:\n"
                "💰 <b>Savdo</b> — bugun, kecha, 7/30 kun\n"
                "📦 <b>Ombor</b> — qoldiq va prognoz\n"
                "📊 <b>Hisobotlar</b> — Excel va tekshiruv"
            )
        else:
            text_ok = (
                "✅ <b>Магазин подключён</b>\n\n"
                f"Найдено магазинов: <b>{len(shops)}</b>\n"
                + "\n".join(lines)
                + "\n\nАктивный магазин: "
                + (f"<code>{default_shop_id}</code>" if default_shop_id else "не выбран")
                + "\n\nТеперь пользуйтесь основными разделами:\n"
                "💰 <b>Продажи</b> — сегодня, вчера, 7/30 дней\n"
                "📦 <b>Склад</b> — остатки и прогноз\n"
                "📊 <b>Отчёты</b> — Excel и проверка подключения"
            )
        await message.answer(text_ok, reply_markup=menu_for_message(message))
        if state:
            await state.clear()
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("language", "lang", "til"))
async def language_command(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    await message.answer(
        f"{tr(lang, 'language_title')}\n\n{tr(lang, 'language_body')}",
        reply_markup=language_markup(),
    )


@dp.message(F.text == "🌐 Язык")
@dp.message(F.text == "🌐 Til")
async def language_button(message: Message) -> None:
    await language_command(message)


@dp.callback_query(F.data.startswith("set_lang:"))
async def set_language_callback(callback: CallbackQuery) -> None:
    if not callback.from_user:
        return
    telegram_id = callback.from_user.id
    db.upsert_user(telegram_id, callback.from_user.username, callback.from_user.first_name)
    lang = normalize_lang((callback.data or "").split(":", 1)[-1])
    set_user_language(telegram_id, lang)
    await callback.answer("OK")
    if callback.message:
        await callback.message.answer(tr(lang, "language_set"), reply_markup=main_menu_for_user(telegram_id))


@dp.message(Command("start", "help"))
async def start(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    ensure_subscription(telegram_id)
    lang = get_user_language(telegram_id)
    connected = "✅ подключён" if db.has_uzum_connection(telegram_id) else "❌ не подключён"
    if lang == "uz":
        connected = "✅ ulangan" if db.has_uzum_connection(telegram_id) else "❌ ulanmagan"
    sub_line = subscription_status_text(telegram_id)

    admin_part = ""
    if is_admin(telegram_id):
        if lang == "uz":
            admin_part = (
                "\n👑 <b>Admin buyruqlari</b>\n"
                "• <code>/admin</code> — admin panel\n"
                "• <code>/users</code> — foydalanuvchilar\n"
                "• <code>/extend ID 30</code> — kirishni uzaytirish\n"
                "• <code>/block ID</code> / <code>/unblock ID</code> — bloklash\n"
                "• <code>/paid ID summa kun</code> — to‘lovni yozish va uzaytirish\n"
                "• <code>/payments</code> — to‘lovlar tarixi\n"
                "• <code>/backup_db</code> — baza zaxirasi\n"
                "• <code>/broadcast matn</code> — hammaga xabar yuborish\n"
            )
        else:
            admin_part = (
                "\n👑 <b>Админ-команды</b>\n"
                "• <code>/admin</code> — админ-панель\n"
                "• <code>/users</code> — список пользователей\n"
                "• <code>/extend ID 30</code> — продлить доступ\n"
                "• <code>/block ID</code> / <code>/unblock ID</code> — блокировка\n"
                "• <code>/paid ID сумма дни</code> — записать оплату и продлить\n"
                "• <code>/payments</code> — история оплат\n"
                "• <code>/backup_db</code> — резервная копия базы\n"
                "• <code>/broadcast текст</code> — рассылка\n"
            )

    if lang == "uz":
        text = (
            "👋 <b>Uzum Seller Assistant</b>\n\n"
            "Sotuv, qoldiq va hisobotlarni Telegramda ko‘rish uchun yordamchi bot.\n\n"
            f"Uzum API: {connected}\n"
            f"Kirish: {sub_line}\n"
            f"Til: <b>{language_title(lang)}</b>\n\n"
            "🚀 <b>Boshlash uchun 3 qadam:</b>\n"
            "1. <code>/video</code> — API ulash videosini ko‘ring\n"
            "2. <code>/connect</code> — API-kalitni botga yuboring\n"
            "3. <b>💰 Savdo</b> yoki <b>📦 Ombor</b> bo‘limini tanlang\n\n"
            "Asosiy menyuda faqat eng kerakli bo‘limlar qoldirildi.\n"
            "Yordam: <code>/support</code>"
            + admin_part
        )
    else:
        text = (
            "👋 <b>Uzum Seller Assistant</b>\n\n"
            "Помощник для селлеров Uzum: продажи, остатки, уведомления и отчёты прямо в Telegram.\n\n"
            f"Uzum API: {connected}\n"
            f"Доступ: {sub_line}\n"
            f"Язык: <b>{language_title(lang)}</b>\n\n"
            "🚀 <b>Начать в 3 шага:</b>\n"
            "1. <code>/video</code> — посмотрите видеоинструкцию\n"
            "2. <code>/connect</code> — отправьте API-ключ боту\n"
            "3. Выберите раздел <b>💰 Продажи</b> или <b>📦 Склад</b>\n\n"
            "В главном меню оставлены только самые нужные разделы, чтобы не путаться.\n"
            "Поддержка: <code>/support</code>"
            + admin_part
        )
    await message.answer(text, reply_markup=menu_for_message(message))


@dp.message(Command("menu"))
async def menu(message: Message) -> None:
    upsert_from_message(message)
    await message.answer(tr_user(upsert_from_message(message), "choose_section"), reply_markup=menu_for_message(message))


@dp.message(Command("cancel"))
async def cancel(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer(tr_user(upsert_from_message(message), "cancelled"), reply_markup=menu_for_message(message))


@dp.message(Command("status"))
async def status(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    ensure_subscription(telegram_id)
    user = db.get_user(telegram_id)
    shops = db.list_shops(telegram_id)
    connected = bool(user and user["uzum_token_encrypted"])
    default_shop_id = user["default_shop_id"] if user else None
    await message.answer(
        "⚙️ <b>Статус</b>\n\n"
        f"Uzum API: {'✅ подключён' if connected else '❌ не подключён'}\n"
        f"Магазинов: {len(shops)}\n"
        f"Основной магазин: {f'<code>{default_shop_id}</code>' if default_shop_id else 'не выбран'}\n"
        f"Подписка: {subscription_status_text(telegram_id)}\n",
        reply_markup=menu_for_message(message),
    )


def api_already_connected_text(lang: str) -> str:
    if lang == "uz":
        return (
            "✅ <b>Do‘kon allaqachon ulangan</b>\n\n"
            "Tasodifan <b>🔌 Ulash</b> tugmasini bossangiz ham, eski API-kalit o‘chmaydi.\n\n"
            "API-kalitni almashtirish kerak bo‘lsa, faqat shunda <code>/reconnect</code> buyrug‘ini yuboring.\n"
            "Ulanishni butunlay o‘chirish uchun: <code>/disconnect</code>"
        )
    return (
        "✅ <b>Магазин уже подключён</b>\n\n"
        "Если вы случайно нажали <b>🔌 Подключить</b>, ничего страшного — старый API-ключ не удалён и не слетит.\n\n"
        "Чтобы заменить API-ключ, используйте только команду <code>/reconnect</code>.\n"
        "Чтобы полностью удалить подключение: <code>/disconnect</code>"
    )


@dp.message(Command("connect_shop", "staff_connect", "addshop"))
async def connect_shop_command(message: Message, state: FSMContext) -> None:
    # Подключение только через API-ключ продавца.
    # Безопасность: если API уже подключён, обычная кнопка/команда не переводит пользователя
    # в режим замены и не трогает старый ключ. Для замены есть отдельная команда /reconnect.
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    if db.has_uzum_connection(telegram_id):
        await state.clear()
        await message.answer(api_already_connected_text(lang), reply_markup=menu_for_message(message))
        return

    await state.set_state(ConnectStates.waiting_for_token)
    if lang == "uz":
        text = (
            "🔌 <b>Do‘konni ulash</b>\n\n"
            "Do‘konni ulash uchun Uzum Seller kabinetidan API-kalit yarating va shu yerga yuboring.\n\n"
            "🎥 Videoqo‘llanma: <code>/video</code>\n"
            "📌 Yozma yo‘riqnoma: <code>/api_token</code>\n\n"
            "API-kalitni keyingi xabarda yuboring.\n"
            "Bekor qilish: <code>/cancel</code>"
        )
    else:
        text = (
            "🔌 <b>Подключение магазина</b>\n\n"
            "Чтобы подключить магазин, создайте API-ключ в кабинете Uzum Seller и отправьте его сюда.\n\n"
            "🎥 Видеоинструкция: <code>/video</code>\n"
            "📌 Текстовая инструкция: <code>/api_token</code>\n\n"
            "Отправьте API-ключ следующим сообщением.\n"
            "Отмена: <code>/cancel</code>"
        )
    await message.answer(text, reply_markup=menu_for_message(message))


@dp.message(F.text == "🔌 Подключить")
@dp.message(F.text == "🔌 Ulash")
async def connect_shop_button(message: Message, state: FSMContext) -> None:
    await connect_shop_command(message, state)


@dp.message(ConnectStates.waiting_for_shop_id, F.text)
async def connect_waiting_shop_id(message: Message, state: FSMContext) -> None:
    await connect_token(message, message.text or "", state)


@dp.message(Command("connect", "reconnect"))
async def connect(message: Message, state: FSMContext) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    raw_text = (message.text or "").strip()
    command = raw_text.split()[0].lower() if raw_text.startswith("/") else ""
    is_reconnect = command.startswith("/reconnect")
    token = parse_args(raw_text)

    # Безопасность: /connect не заменяет уже подключённый ключ.
    # Замена разрешена только через явную команду /reconnect или если пользователь сразу передал токен в /reconnect <token>.
    if db.has_uzum_connection(telegram_id) and not is_reconnect:
        await state.clear()
        await message.answer(api_already_connected_text(lang), reply_markup=menu_for_message(message))
        return

    if token:
        await connect_token(message, token, state)
        return

    await state.set_state(ConnectStates.waiting_for_token)
    if lang == "uz":
        title = "🔁 <b>API-kalitni almashtirish</b>" if is_reconnect else "🔑 <b>API-kalitni ulash</b>"
        text_connect = (
            f"{title}\n\n"
            "Uzum Seller kabinetidan olingan API-kalitni keyingi xabarda yuboring.\n\n"
            "Kalitni qayerdan olish: <code>/api_token</code>\n\n"
            "Muhim:\n"
            "• API-kalit kabinet paroli emas;\n"
            "• eski kalit faqat yangi kalit muvaffaqiyatli tekshirilgandan keyin almashtiriladi;\n"
            "• tekshiruvdan so‘ng bot kalit yuborilgan xabarni o‘chirishga harakat qiladi;\n"
            "• bekor qilish: <code>/cancel</code>."
        )
    else:
        title = "🔁 <b>Замена API-ключа</b>" if is_reconnect else "🔑 <b>Подключение API-ключа</b>"
        text_connect = (
            f"{title}\n\n"
            "Отправьте следующим сообщением API-ключ из кабинета Uzum Seller.\n\n"
            "Где взять ключ: <code>/api_token</code>\n\n"
            "Важно:\n"
            "• API-ключ — это не пароль от кабинета;\n"
            "• старый ключ заменится только после успешной проверки нового;\n"
            "• после проверки бот постарается удалить сообщение с ключом;\n"
            "• отменить: <code>/cancel</code>."
        )
    await message.answer(text_connect, reply_markup=menu_for_message(message))


@dp.message(ConnectStates.waiting_for_token, F.text)
async def connect_waiting_token(message: Message, state: FSMContext) -> None:
    # Пока бот ждёт API-ключ, команды помощи не должны восприниматься как токен.
    # Иначе /video или /api_token попадают в проверку токена и пользователь видит ошибку.
    raw_text = (message.text or "").strip()
    command = raw_text.split()[0].lower() if raw_text.startswith("/") else ""

    if command in {"/video", "/api_video", "/instruction"}:
        await video_instruction(message)
        return

    if command in {"/api_token", "/token_help", "/how_token"}:
        await api_token_help(message)
        return

    if command == "/cancel":
        await cancel(message, state)
        return

    if command == "/menu":
        await state.clear()
        await menu(message)
        return

    if raw_text.startswith("/"):
        telegram_id = upsert_from_message(message)
        lang = get_user_language(telegram_id)
        if lang == "uz":
            await message.answer(
                "Hozir bot API-kalitni kutyapti.\n\n"
                "API-kalitni yuboring yoki bekor qilish uchun <code>/cancel</code> bosing.\n"
                "Yordam: <code>/video</code> yoki <code>/api_token</code>",
                reply_markup=menu_for_message(message),
            )
        else:
            await message.answer(
                "Сейчас бот ждёт API-ключ.\n\n"
                "Отправьте API-ключ или нажмите <code>/cancel</code>, чтобы отменить подключение.\n"
                "Помощь: <code>/video</code> или <code>/api_token</code>",
                reply_markup=menu_for_message(message),
            )
        return

    await connect_token(message, raw_text, state)


def disconnect_uzum_for_user(telegram_id: int) -> None:
    if hasattr(db, "disconnect_uzum"):
        db.disconnect_uzum(telegram_id)
        return
    with db.connect() as conn:
        conn.execute(
            "UPDATE users SET uzum_token_encrypted = NULL, default_shop_id = NULL, updated_at = ? WHERE telegram_id = ?",
            (_dt_to_db(_utc_now()), int(telegram_id)),
        )
        try:
            conn.execute("DELETE FROM shops WHERE telegram_id = ?", (int(telegram_id),))
        except Exception:
            pass
        conn.commit()


@dp.message(Command("disconnect"))
async def disconnect(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    disconnect_uzum_for_user(telegram_id)
    await message.answer(
        "✅ Подключение к Uzum API удалено. Можно подключить заново через <code>/connect</code>.",
        reply_markup=menu_for_message(message),
    )


@dp.message(Command("pinguzum"))
async def ping_uzum(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    client = get_uzum_for_user(telegram_id)
    if client is None:
        await message.answer(
            "Сначала подключите Uzum API-токен: <code>/connect</code>",
            reply_markup=menu_for_message(message),
        )
        return
    try:
        data = await client.get_shops()
        shops = extract_items(data)
        await message.answer(f"✅ Uzum API отвечает. Найдено магазинов: {len(shops)}", reply_markup=menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("shops"))
async def shops(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    client = get_uzum_for_user(telegram_id)
    if client is None:
        await message.answer("Сначала подключите Uzum API-токен: <code>/connect</code>", reply_markup=menu_for_message(message))
        return

    try:
        data = await client.get_shops()
        items = extract_items(data)
        if not items:
            await message.answer(
                "Ответ получен, но список магазинов не найден:\n<code>"
                + escape(compact_json_preview(data))
                + "</code>",
                reply_markup=menu_for_message(message),
            )
            return

        encrypted = db.get_encrypted_token(telegram_id)
        if encrypted:
            db.save_connection(telegram_id, encrypted, items)
        current = db.get_default_shop_id(telegram_id)
        lines = [format_shop_line(item) for item in items[:30]]
        await message.answer(
            "🏪 <b>Ваши магазины:</b>\n\n"
            + "\n".join(lines)
            + "\n\nТекущий основной магазин: "
            + (f"<code>{current}</code>" if current else "не выбран")
            + "\n\nЧтобы выбрать: <code>/setshop SHOP_ID</code>",
            reply_markup=menu_for_message(message),
        )
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("setshop"))
async def setshop(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    arg = parse_args(message.text or "")
    if not arg.isdigit():
        await message.answer(
            "Напишите так: <code>/setshop SHOP_ID</code>\nНапример: <code>/setshop 12345</code>",
            reply_markup=menu_for_message(message),
        )
        return

    shop_id = int(arg)
    ok = db.set_default_shop_id(telegram_id, shop_id)
    if not ok:
        await message.answer("Этот магазин не найден среди подключённых. Сначала обновите список: <code>/shops</code>", reply_markup=menu_for_message(message))
        return

    await message.answer(f"✅ Основной магазин выбран: <code>{shop_id}</code>", reply_markup=menu_for_message(message))


@dp.message(Command("products"))
async def products(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    search_query = parse_args(message.text or "")

    try:
        data = await client.get_products(shop_id, search_query=search_query, page=0, size=10)
        items = extract_items(data)
        if not items:
            await message.answer(
                "Товары не найдены. Ответ API:\n<code>"
                + escape(compact_json_preview(data))
                + "</code>",
                reply_markup=menu_for_message(message),
            )
            return

        title = f"📦 <b>Товары магазина</b> <code>{shop_id}</code>"
        if search_query:
            title += f" по запросу “{escape(search_query)}”"
        lines = [format_product_line(item) for item in items[:10]]
        await message.answer(title + ":\n\n" + "\n\n".join(lines), reply_markup=menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


STOCK_PAGE_SIZE = int(os.getenv("STOCK_PAGE_SIZE", "10"))
_stock_page_cache: dict[int, dict[str, Any]] = {}


def _stock_page_markup(page: int, total_pages: int, lang: str) -> InlineKeyboardMarkup | None:
    if total_pages <= 1:
        return None
    buttons: list[InlineKeyboardButton] = []
    if page > 0:
        buttons.append(InlineKeyboardButton(text="⬅️ Oldingi" if lang == "uz" else "⬅️ Назад", callback_data=f"stockpg:{page-1}"))
    buttons.append(InlineKeyboardButton(text=f"{page + 1}/{total_pages}", callback_data="stockpg:noop"))
    if page < total_pages - 1:
        buttons.append(InlineKeyboardButton(text="Keyingi ➡️" if lang == "uz" else "След. страница ➡️", callback_data=f"stockpg:{page+1}"))
    return InlineKeyboardMarkup(inline_keyboard=[buttons])


def _stock_page_text(session: dict[str, Any], page: int) -> str:
    rows = session["rows"]
    mode = session["mode"]
    title = session["title"]
    lang = session["lang"]
    page_size = session.get("page_size") or STOCK_PAGE_SIZE
    total = len(rows)
    total_pages = max(1, (total + page_size - 1) // page_size)
    page = max(0, min(page, total_pages - 1))
    start_i = page * page_size
    chunk = rows[start_i:start_i + page_size]
    lines = [format_sku_stock_line(row, mode=mode) for row in chunk]
    start_n = start_i + 1 if total else 0
    end_n = start_i + len(chunk)
    if lang == "uz":
        header = f"{title}\nKo‘rsatilmoqda: {start_n}–{end_n} / {total}\nSahifa: {page + 1}/{total_pages}"
    else:
        header = f"{title}\nПоказано: {start_n}–{end_n} из {total}\nСтраница: {page + 1}/{total_pages}"
    return header + "\n\n" + "\n\n".join(lines)


async def send_stock_list(message: Message, mode: str = "all") -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    search_query = parse_args(message.text or "")
    lang = get_user_language(message.from_user.id if message.from_user else None)

    try:
        rows = await load_sku_rows(client, shop_id, search_query=search_query, max_pages=20)
        if not rows:
            text = "SKU qoldiqlari topilmadi." if lang == "uz" else "SKU-остатки не найдены."
            await message.answer(text, reply_markup=menu_for_message(message))
            return

        if mode == "fbo":
            rows = [r for r in rows if (r.get("fbo") or 0) > 0]
            title = "📦 <b>FBO qoldiqlari / Uzum ombori</b>" if lang == "uz" else "📦 <b>Остатки FBO / склад Uzum</b>"
        elif mode == "fbs":
            rows = [r for r in rows if (r.get("fbs") or 0) > 0]
            title = "📦 <b>FBS/DBS qoldiqlari / sotuvchi ombori</b>" if lang == "uz" else "📦 <b>Остатки FBS/DBS / склад продавца</b>"
        else:
            title = "📦 <b>SKU bo‘yicha qoldiqlar: FBO + FBS/DBS + jami</b>" if lang == "uz" else "📦 <b>Остатки по SKU: FBO + FBS/DBS + итого</b>"

        if search_query:
            title += (f"\nQidiruv: {escape(search_query)}" if lang == "uz" else f"\nПоиск: {escape(search_query)}")
        if not rows:
            empty = "Hech narsa topilmadi." if lang == "uz" else "Ничего не найдено."
            await message.answer(title + "\n\n" + empty, reply_markup=menu_for_message(message))
            return

        user_id = message.from_user.id if message.from_user else 0
        page_size = max(5, min(20, STOCK_PAGE_SIZE))
        total_pages = max(1, (len(rows) + page_size - 1) // page_size)
        _stock_page_cache[user_id] = {
            "rows": rows,
            "mode": mode,
            "title": title,
            "lang": lang,
            "page_size": page_size,
        }
        await message.answer(
            _stock_page_text(_stock_page_cache[user_id], 0),
            reply_markup=_stock_page_markup(0, total_pages, lang),
        )
    except Exception as e:
        await send_api_error(message, e)


@dp.callback_query(F.data.startswith("stockpg:"))
async def stock_page_callback(callback: CallbackQuery) -> None:
    if not callback.from_user:
        await callback.answer()
        return
    raw = (callback.data or "").split(":", 1)[-1]
    if raw == "noop":
        await callback.answer()
        return
    try:
        page = int(raw)
    except ValueError:
        await callback.answer()
        return
    session = _stock_page_cache.get(callback.from_user.id)
    lang = get_user_language(callback.from_user.id)
    if not session:
        await callback.answer("Список устарел. Нажмите Остатки заново." if lang != "uz" else "Ro‘yxat eskirdi. Qoldiqni qayta bosing.", show_alert=True)
        return
    rows = session["rows"]
    page_size = session.get("page_size") or STOCK_PAGE_SIZE
    total_pages = max(1, (len(rows) + page_size - 1) // page_size)
    page = max(0, min(page, total_pages - 1))
    try:
        if callback.message:
            await callback.message.edit_text(
                _stock_page_text(session, page),
                reply_markup=_stock_page_markup(page, total_pages, session.get("lang", lang)),
            )
        await callback.answer()
    except Exception:
        await callback.answer("Не получилось открыть страницу" if lang != "uz" else "Sahifani ochib bo‘lmadi", show_alert=True)




# --- Универсальная постраничность для длинных списков ---
LIST_PAGE_SIZE = int(os.getenv("LIST_PAGE_SIZE", "10") or "10")
_paged_list_cache: dict[tuple[int, str], dict[str, Any]] = {}


def _page_size_safe(value: int | None = None) -> int:
    raw = value or LIST_PAGE_SIZE
    return max(5, min(20, int(raw or 10)))


def _section_text_and_markup(section: str, telegram_id: int, lang: str) -> tuple[str, ReplyKeyboardMarkup]:
    if section == "sales":
        text = "💰 <b>Savdo bo‘limi</b>\nKerakli davr yoki hisobotni tanlang 👇" if lang == "uz" else "💰 <b>Продажи</b>\nВыберите, что посмотреть 👇"
        return text, sales_menu_for_user(telegram_id)
    if section == "stock":
        text = "📦 <b>Ombor</b>\nQoldiq, prognoz yoki FBO yuk xatlarini tanlang 👇" if lang == "uz" else "📦 <b>Склад</b>\nОстатки, прогноз и FBO-накладные 👇"
        return text, stock_menu_for_user(telegram_id)
    if section == "attention":
        text = "🧠 <b>Nimani tekshirish kerak</b>\nKerakli bo‘limni tanlang 👇" if lang == "uz" else "🧠 <b>Что проверить</b>\nВыберите нужный раздел 👇"
        return text, attention_menu_for_user(telegram_id)
    if section == "reports":
        text = "📊 <b>Hisobotlar</b>\nExcel, foyda va tayyor hisobotlar 👇" if lang == "uz" else "📊 <b>Отчёты</b>\nExcel, прибыль и готовые отчёты 👇"
        return text, report_menu_for_user(telegram_id)
    text = "🏠 <b>Asosiy menyu</b>" if lang == "uz" else "🏠 <b>Главное меню</b>"
    return text, main_menu_for_user(telegram_id)


def _paged_markup(kind: str, page: int, total_pages: int, lang: str, section: str = "main") -> InlineKeyboardMarkup | None:
    rows: list[list[InlineKeyboardButton]] = []
    if total_pages > 1:
        nav: list[InlineKeyboardButton] = []
        if page > 0:
            nav.append(InlineKeyboardButton(text="⬅️ Oldingi" if lang == "uz" else "⬅️ Назад", callback_data=f"pglist:{kind}:{page-1}"))
        nav.append(InlineKeyboardButton(text=f"{page + 1}/{total_pages}", callback_data="pgnoop"))
        if page < total_pages - 1:
            nav.append(InlineKeyboardButton(text="Keyingi ➡️" if lang == "uz" else "След. страница ➡️", callback_data=f"pglist:{kind}:{page+1}"))
        rows.append(nav)
    rows.append([
        InlineKeyboardButton(text="⬅️ Bo‘limga qaytish" if lang == "uz" else "⬅️ Назад в раздел", callback_data=f"pgsection:{section}"),
        InlineKeyboardButton(text="🏠 Menyu" if lang == "uz" else "🏠 Меню", callback_data="pgmain"),
    ])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def _paged_text(session: dict[str, Any], page: int) -> str:
    items: list[str] = session.get("items") or []
    page_size = _page_size_safe(session.get("page_size"))
    total = len(items)
    total_pages = max(1, (total + page_size - 1) // page_size)
    page = max(0, min(page, total_pages - 1))
    start_i = page * page_size
    chunk = items[start_i:start_i + page_size]
    start_n = start_i + 1 if total else 0
    end_n = start_i + len(chunk)
    lang = session.get("lang", "ru")
    title = session.get("title", "")
    summary = [x for x in (session.get("summary") or []) if x]
    if lang == "uz":
        meta = f"Ko‘rsatilmoqda: <b>{start_n}–{end_n}</b> / <b>{total}</b>\nSahifa: <b>{page + 1}/{total_pages}</b>"
    else:
        meta = f"Показано: <b>{start_n}–{end_n}</b> из <b>{total}</b>\nСтраница: <b>{page + 1}/{total_pages}</b>"
    parts = [title, *summary, meta]
    if chunk:
        parts.append("\n\n".join(chunk))
    return "\n\n".join(parts)


async def send_paginated_list(
    message: Message,
    *,
    kind: str,
    title: str,
    items: list[str],
    summary: list[str] | None = None,
    empty_text: str | None = None,
    section: str = "main",
    page_size: int | None = None,
    reply_markup: ReplyKeyboardMarkup | None = None,
) -> None:
    user_id = message.from_user.id if message.from_user else 0
    lang = get_user_language(user_id)
    if not items:
        await message.answer(empty_text or ("Ma’lumot topilmadi." if lang == "uz" else "Данные не найдены."), reply_markup=reply_markup or menu_for_message(message))
        return
    session = {
        "title": title,
        "summary": summary or [],
        "items": items,
        "lang": lang,
        "section": section,
        "page_size": _page_size_safe(page_size),
    }
    _paged_list_cache[(user_id, kind)] = session
    total_pages = max(1, (len(items) + session["page_size"] - 1) // session["page_size"])
    await message.answer(
        _paged_text(session, 0),
        reply_markup=_paged_markup(kind, 0, total_pages, lang, section),
    )


@dp.callback_query(F.data == "pgnoop")
async def paged_noop_callback(callback: CallbackQuery) -> None:
    await callback.answer()


@dp.callback_query(F.data == "pgmain")
async def paged_main_callback(callback: CallbackQuery) -> None:
    uid = callback.from_user.id if callback.from_user else 0
    lang = get_user_language(uid)
    text, markup = _section_text_and_markup("main", uid, lang)
    if callback.message:
        await callback.message.answer(text, reply_markup=markup)
    await callback.answer()


@dp.callback_query(F.data.startswith("pgsection:"))
async def paged_section_callback(callback: CallbackQuery) -> None:
    uid = callback.from_user.id if callback.from_user else 0
    lang = get_user_language(uid)
    section = (callback.data or "").split(":", 1)[-1]
    text, markup = _section_text_and_markup(section, uid, lang)
    if callback.message:
        await callback.message.answer(text, reply_markup=markup)
    await callback.answer()


@dp.callback_query(F.data.startswith("pglist:"))
async def paged_list_callback(callback: CallbackQuery) -> None:
    uid = callback.from_user.id if callback.from_user else 0
    lang = get_user_language(uid)
    try:
        _, kind, raw_page = (callback.data or "").split(":", 2)
        page = int(raw_page)
    except Exception:
        await callback.answer()
        return
    session = _paged_list_cache.get((uid, kind))
    if not session:
        await callback.answer("Ro‘yxat eskirdi. Bo‘limni qayta oching." if lang == "uz" else "Список устарел. Откройте раздел заново.", show_alert=True)
        return
    page_size = _page_size_safe(session.get("page_size"))
    total_pages = max(1, (len(session.get("items") or []) + page_size - 1) // page_size)
    page = max(0, min(page, total_pages - 1))
    try:
        if callback.message:
            await callback.message.edit_text(
                _paged_text(session, page),
                reply_markup=_paged_markup(kind, page, total_pages, session.get("lang", lang), session.get("section", "main")),
            )
        await callback.answer()
    except Exception:
        await callback.answer("Sahifani ochib bo‘lmadi" if lang == "uz" else "Не получилось открыть страницу", show_alert=True)

@dp.message(Command("stock"))
async def stock(message: Message) -> None:
    await send_stock_list(message, mode="all")


@dp.message(Command("fbo"))
async def fbo(message: Message) -> None:
    await send_stock_list(message, mode="fbo")


@dp.message(Command("fbs"))
async def fbs(message: Message) -> None:
    await send_stock_list(message, mode="fbs")


@dp.message(Command("lowstock"))
async def lowstock(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    telegram_id = message.from_user.id if message.from_user else 0
    lang = get_user_language(telegram_id)
    arg = parse_args(message.text or "")
    threshold = int(arg) if arg.isdigit() else LOW_STOCK_THRESHOLD

    try:
        rows = await load_sku_rows(client, shop_id, max_pages=50)
        if not rows:
            await message.answer("SKU qoldiqlari topilmadi." if lang == "uz" else "SKU-остатки не найдены.", reply_markup=stock_menu_for_message(message))
            return

        low = [r for r in rows if r.get("total") is not None and r["total"] <= threshold]
        if not low:
            text = f"✅ Umumiy qoldiq ≤ {threshold} bo‘lgan SKU topilmadi." if lang == "uz" else f"✅ Товаров с общим остатком ≤ {threshold} не найдено."
            await message.answer(text, reply_markup=stock_menu_for_message(message))
            return

        items = [format_sku_stock_line(row, mode="all") for row in low]
        title = f"⚠️ <b>Kam qoldiqdagi tovarlar</b>" if lang == "uz" else f"⚠️ <b>Низкие остатки</b>"
        summary = [
            f"🏪 Do‘kon: <code>{shop_id}</code>" if lang == "uz" else f"🏪 Магазин: <code>{shop_id}</code>",
            f"📉 Chegara: ≤ <b>{threshold}</b> dona" if lang == "uz" else f"📉 Порог: ≤ <b>{threshold}</b> шт.",
        ]
        await send_paginated_list(message, kind="lowstock", title=title, summary=summary, items=items, section="stock", reply_markup=stock_menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("orders"))
async def orders(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    status = parse_args(message.text or "").upper() or "CREATED"

    try:
        data = await client.get_fbs_orders(shop_id, status=status, page=0, size=10)
        items = extract_items(data)
        if not items:
            await message.answer(f"Заказы со статусом <code>{escape(status)}</code> не найдены.", reply_markup=menu_for_message(message))
            return

        lines = [format_order_line(item) for item in items[:10]]
        await message.answer(
            f"🛒 <b>Заказы {escape(status)} для магазина</b> <code>{shop_id}</code>:\n\n"
            + "\n".join(lines),
            reply_markup=menu_for_message(message),
        )
    except Exception as e:
        await send_api_error(message, e)



# --- Продажи / Финансы ---
# Используем официальный Finance endpoint Uzum Seller OpenAPI:
# GET /v1/finance/orders?shopIds=...&dateFrom=...&dateTo=...

def _epoch_ms(dt: datetime) -> int:
    return int(dt.timestamp() * 1000)


UZT = timezone(timedelta(hours=5))


def _today_range_ms() -> tuple[int, int]:
    # Считаем день по времени Узбекистана, а не по UTC.
    now = datetime.now(UZT)
    start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    return _epoch_ms(start), _epoch_ms(now)


def _days_range_ms(days: int) -> tuple[int, int]:
    # 7 дней = с начала дня 6 дней назад до текущего момента по Ташкенту.
    now = datetime.now(UZT)
    start_today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    start = start_today - timedelta(days=max(1, days) - 1)
    return _epoch_ms(start), _epoch_ms(now)


def _yesterday_range_ms() -> tuple[int, int]:
    # Вчера = прошлый полный день по времени Узбекистана.
    now = datetime.now(UZT)
    start_today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    start_yesterday = start_today - timedelta(days=1)
    return _epoch_ms(start_yesterday), _epoch_ms(start_today)


def _last_7_days_range_ms() -> tuple[int, int]:
    return _days_range_ms(7)


def _deep_items(obj: Any) -> list[dict[str, Any]]:
    """Достаём список строк из разных возможных форматов ответа Uzum."""
    direct = extract_items(obj)
    if direct:
        return [x for x in direct if isinstance(x, dict)]

    keys = (
        "orderItems",
        "orders",
        "items",
        "content",
        "data",
        "payload",
        "result",
        "list",
        "financeOrders",
        "sellerOrders",
    )
    found: list[dict[str, Any]] = []

    def walk(x: Any) -> None:
        if isinstance(x, dict):
            for k in keys:
                v = x.get(k)
                if isinstance(v, list):
                    for item in v:
                        if isinstance(item, dict):
                            found.append(item)
                elif isinstance(v, dict):
                    walk(v)
        elif isinstance(x, list):
            for item in x:
                if isinstance(item, dict):
                    found.append(item)

    walk(obj)
    # Убираем явные дубли по JSON-представлению.
    unique: list[dict[str, Any]] = []
    seen: set[str] = set()
    for item in found:
        sig = json.dumps(item, ensure_ascii=False, sort_keys=True, default=str)[:1000]
        if sig not in seen:
            seen.add(sig)
            unique.append(item)
    return unique


def _num_from_value(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = (
            value.replace(" ", "")
            .replace("\u00a0", "")
            .replace("сум", "")
            .replace("UZS", "")
            .replace(",", ".")
        )
        try:
            return float(cleaned)
        except Exception:
            return None
    return None


def _pick_number(item: dict[str, Any], names: tuple[str, ...]) -> float | None:
    for name in names:
        value = item.get(name)
        number = _num_from_value(value)
        if number is not None:
            return number
    return None


def _finance_status(item: dict[str, Any]) -> str:
    value = pick(
        item,
        "status",
        "orderStatus",
        "financeStatus",
        "state",
        "statusName",
        "statusTitle",
    )
    if isinstance(value, dict):
        value = pick(value, "title", "name", "value", "code")
    return str(value or "UNKNOWN")


def _finance_title(item: dict[str, Any]) -> str:
    value = pick(
        item,
        "skuTitle",
        "productTitle",
        "productName",
        "title",
        "name",
        "skuName",
        "offerName",
    )
    if isinstance(value, dict):
        value = pick(value, "title", "name")
    return str(value or "Без названия")


def _finance_qty(item: dict[str, Any]) -> float:
    return _pick_number(
        item,
        (
            "quantity",
            "amount",
            "count",
            "qty",
            "skuAmount",
            "productAmount",
            "quantityPurchased",
        ),
    ) or 1.0


def _finance_revenue(item: dict[str, Any]) -> float:
    # Пробуем готовые суммы.
    direct = _pick_number(
        item,
        (
            "totalPrice",
            "totalAmount",
            "totalSum",
            "sellerAmount",
            "sellerPrice",
            "totalSellerPrice",
            "purchasePrice",
            "priceWithDiscount",
            "orderItemPrice",
            "orderPrice",
            "amountToWithdraw",
            "accrual",
            "sum",
        ),
    )
    if direct is not None:
        return max(0.0, direct)

    # Если есть только цена за штуку — умножаем на количество.
    price = _pick_number(item, ("price", "itemPrice", "skuPrice", "sellPrice"))
    if price is not None:
        return max(0.0, price * _finance_qty(item))
    return 0.0


def _is_cancelled_status(status: str) -> bool:
    s = status.upper()
    return "CANCEL" in s or "ОТМЕН" in s


def _format_money(value: float) -> str:
    return f"{value:,.0f}".replace(",", " ") + " сум"


async def _finance_orders_request(
    client: UzumClient,
    shop_id: int,
    *,
    date_from_ms: int,
    date_to_ms: int,
    page: int = 0,
    size: int = 100,
) -> Any:
    params = [
        ("page", page),
        ("size", size),
        ("group", "false"),
        # Важно: рабочий Noorza Bot использует dateFrom в секундах, dateTo в миллисекундах.
        # Если отправить dateFrom в миллисекундах, Uzum Finance может вернуть 0 строк.
        ("dateFrom", int(date_from_ms / 1000)),
        ("dateTo", date_to_ms),
        ("shopIds", shop_id),
    ]
    path = "/v1/finance/orders?" + urlencode(params)
    return await client._request("GET", path)


async def _load_finance_orders(
    client: UzumClient,
    shop_id: int,
    *,
    date_from_ms: int,
    date_to_ms: int,
    max_pages: int = 10,
    page_size: int = 100,
) -> tuple[list[dict[str, Any]], Any | None]:
    rows: list[dict[str, Any]] = []
    first_response: Any | None = None
    for page in range(max_pages):
        data = await _finance_orders_request(
            client,
            shop_id,
            date_from_ms=date_from_ms,
            date_to_ms=date_to_ms,
            page=page,
            size=page_size,
        )
        if first_response is None:
            first_response = data
        items = _deep_items(data)
        if not items:
            break
        rows.extend(items)
        if len(items) < page_size:
            break
    return rows, first_response


def _build_sales_stats(rows: list[dict[str, Any]]) -> dict[str, Any]:
    total_rows = len(rows)
    cancelled_rows = 0
    revenue = 0.0
    units = 0.0
    statuses: dict[str, int] = {}
    products: dict[str, dict[str, float | str]] = {}

    for item in rows:
        status = _finance_status(item)
        statuses[status] = statuses.get(status, 0) + 1
        qty = _finance_qty(item)
        amount = _finance_revenue(item)
        if _is_cancelled_status(status):
            cancelled_rows += 1
            continue
        revenue += amount
        units += qty
        title = _finance_title(item)
        if title not in products:
            products[title] = {"title": title, "qty": 0.0, "revenue": 0.0}
        products[title]["qty"] = float(products[title]["qty"]) + qty
        products[title]["revenue"] = float(products[title]["revenue"]) + amount

    top_products = sorted(
        products.values(), key=lambda x: float(x.get("revenue") or 0), reverse=True
    )[:5]
    avg = revenue / max(1, (total_rows - cancelled_rows))
    return {
        "rows": total_rows,
        "cancelled": cancelled_rows,
        "active_rows": max(0, total_rows - cancelled_rows),
        "revenue": revenue,
        "units": units,
        "avg": avg,
        "statuses": statuses,
        "top_products": top_products,
    }


def _short_period_title(days: int) -> str:
    if days == 1:
        return "сегодня"
    return f"за {days} дней"


async def _sales_period_stats(
    client: UzumClient, shop_id: int, days: int
) -> tuple[dict[str, Any], Any | None]:
    if days == 1:
        date_from, date_to = _today_range_ms()
    else:
        date_from, date_to = _days_range_ms(days)
    rows, first = await _load_finance_orders(
        client, shop_id, date_from_ms=date_from, date_to_ms=date_to
    )
    return _build_sales_stats(rows), first


def _format_sales_summary_line(title: str, stats: dict[str, Any]) -> str:
    return (
        f"<b>{escape(title)}</b>\n"
        f"• Выручка: <b>{_format_money(float(stats['revenue']))}</b>\n"
        f"• Позиций/строк: <b>{stats['active_rows']}</b>\n"
        f"• Штук: <b>{float(stats['units']):.0f}</b>\n"
        f"• Средняя строка: <b>{_format_money(float(stats['avg']))}</b>"
    )


def _format_sales_details(days: int, shop_id: int, stats: dict[str, Any], first_response: Any | None) -> str:
    title = _short_period_title(days)
    lines = [
        f"💰 <b>Продажи {title}</b>",
        f"Магазин: <code>{shop_id}</code>",
        "",
        f"Выручка: <b>{_format_money(float(stats['revenue']))}</b>",
        f"Проданных строк/позиций: <b>{stats['active_rows']}</b>",
        f"Штук: <b>{float(stats['units']):.0f}</b>",
        f"Средняя строка: <b>{_format_money(float(stats['avg']))}</b>",
        f"Отменённых строк: <b>{stats['cancelled']}</b>",
    ]

    top_products = stats.get("top_products") or []
    if top_products:
        lines.append("")
        lines.append("<b>Топ товаров по сумме:</b>")
        for idx, item in enumerate(top_products, start=1):
            title_item = str(item.get("title") or "Без названия")
            if len(title_item) > 70:
                title_item = title_item[:67] + "..."
            lines.append(
                f"{idx}. {escape(title_item)} — "
                f"{float(item.get('qty') or 0):.0f} шт, "
                f"{_format_money(float(item.get('revenue') or 0))}"
            )

    if stats.get("rows") == 0 and first_response is not None:
        lines.append("")
        lines.append("Ответ Finance API пришёл, но строки продаж не найдены.")
        lines.append("Фрагмент ответа:")
        lines.append("<code>" + escape(compact_json_preview(first_response)) + "</code>")

    return "\n".join(lines)


# --- Короткие разделы в стиле Noorza Bot ---
# Блок "Сегодня" работает в стиле второго бота: берёт Finance API за текущий день
# и показывает выручку, комиссию, логистику и к выплате.

def _deep_pick_number(obj: Any, names: tuple[str, ...]) -> float | None:
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k in names:
                n = _num_from_value(v)
                if n is not None:
                    return n
        for v in obj.values():
            n = _deep_pick_number(v, names)
            if n is not None:
                return n
    elif isinstance(obj, list):
        for v in obj:
            n = _deep_pick_number(v, names)
            if n is not None:
                return n
    return None


def _finance_gross_revenue(item: dict[str, Any]) -> float:
    # В Finance API Uzum поле sellPrice обычно является ценой за 1 штуку.
    # Поэтому для выручки умножаем sellPrice на amount, как в рабочем Noorza Bot.
    direct_total = _deep_pick_number(
        item,
        (
            "totalPrice", "totalAmount", "totalSum", "totalSellerPrice",
            "orderItemPrice", "orderPrice", "sellerPrice", "sellerAmount",
        ),
    )
    if direct_total is not None:
        return max(0.0, direct_total)

    unit_price = _deep_pick_number(
        item,
        (
            "sellPrice", "soldPrice", "productPrice", "skuPrice",
            "priceWithDiscount", "purchasePrice", "price",
        ),
    )
    if unit_price is not None:
        return max(0.0, unit_price * _finance_qty(item))
    return _finance_revenue(item)


def _finance_commission(item: dict[str, Any]) -> float:
    value = _deep_pick_number(
        item,
        (
            "commission", "commissionAmount", "commissionSum", "uzumCommission",
            "marketplaceCommission", "sellerCommission", "fee", "feeAmount",
        ),
    )
    return abs(value or 0.0)


def _finance_logistics(item: dict[str, Any]) -> float:
    value = _deep_pick_number(
        item,
        (
            "logisticDeliveryFee", "logistics", "logistic", "logisticAmount", "logisticsAmount",
            "logisticsSum", "delivery", "deliveryAmount", "deliveryPrice",
            "deliveryCost", "shipping", "shippingAmount",
        ),
    )
    return abs(value or 0.0)


def _finance_payout_direct(item: dict[str, Any]) -> float | None:
    return _deep_pick_number(
        item,
        (
            "sellerProfit", "amountToWithdraw", "toWithdraw", "withdrawAmount", "sellerPayout",
            "payout", "payoutAmount", "sellerAmount", "accrual", "accrualAmount",
        ),
    )


def _finance_withdrawn(item: dict[str, Any]) -> float:
    value = _deep_pick_number(
        item,
        (
            "withdrawnProfit", "withdrawn", "withdrawnAmount", "paid", "paidAmount", "transferred",
            "transferredAmount", "alreadyWithdrawn",
        ),
    )
    return max(0.0, value or 0.0)


async def _finance_orders_request_extra(
    client: UzumClient,
    shop_id: int,
    *,
    date_from_ms: int,
    date_to_ms: int,
    extra_params: list[tuple[str, Any]] | None = None,
    page: int = 0,
    size: int = 100,
) -> Any:
    params: list[tuple[str, Any]] = [
        ("page", page),
        ("size", size),
        ("group", "false"),
        # Важно: рабочий Noorza Bot использует dateFrom в секундах, dateTo в миллисекундах.
        ("dateFrom", int(date_from_ms / 1000)),
        ("dateTo", date_to_ms),
        ("shopIds", shop_id),
    ]
    if extra_params:
        params.extend(extra_params)
    path = "/v1/finance/orders?" + urlencode(params)
    return await client._request("GET", path)


async def _load_today_finance_flexible(
    client: UzumClient, shop_id: int
) -> tuple[list[dict[str, Any]], Any | None, str]:
    date_from, date_to = _today_range_ms()
    return await _load_finance_range_flexible(client, shop_id, date_from, date_to)


async def _load_finance_range_flexible(
    client: UzumClient, shop_id: int, date_from_ms: int, date_to_ms: int
) -> tuple[list[dict[str, Any]], Any | None, str]:
    attempts: list[tuple[str, list[tuple[str, Any]]]] = [
        ("без статуса", []),
        ("statuses=PROCESSING", [("statuses", "PROCESSING")]),
        ("statuses=TO_WITHDRAW", [("statuses", "TO_WITHDRAW")]),
        ("status=PROCESSING", [("status", "PROCESSING")]),
        ("status=TO_WITHDRAW", [("status", "TO_WITHDRAW")]),
    ]
    all_rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    first_response: Any | None = None
    used_attempts: list[str] = []
    for label, extra in attempts:
        try:
            data = await _finance_orders_request_extra(
                client,
                shop_id,
                date_from_ms=date_from_ms,
                date_to_ms=date_to_ms,
                extra_params=extra,
            )
            if first_response is None:
                first_response = data
            rows = _deep_items(data)
            used_attempts.append(label)
            for row in rows:
                sig = json.dumps(row, ensure_ascii=False, sort_keys=True, default=str)[:1500]
                if sig not in seen:
                    seen.add(sig)
                    all_rows.append(row)
            if all_rows and label == "без статуса":
                break
            await asyncio.sleep(0.15)
        except Exception as e:
            logging.info("Finance attempt failed: %s: %s", label, e)
            continue
    return all_rows, first_response, ", ".join(used_attempts)


def _build_noorza_today_stats(rows: list[dict[str, Any]]) -> dict[str, Any]:
    # Логика максимально приближена к рабочему Noorza Bot:
    # sellPrice * amount, commission, logisticDeliveryFee, sellerProfit, withdrawnProfit.
    active_rows = 0
    returns = 0.0
    units = 0.0
    revenue = 0.0
    commission = 0.0
    logistics = 0.0
    payout_total = 0.0
    withdrawn = 0.0
    statuses: dict[str, int] = {}
    for item in rows:
        status = _finance_status(item)
        qty = _finance_qty(item)
        if _is_cancelled_status(status):
            statuses[status] = statuses.get(status, 0) + 1
            continue

        statuses[status] = statuses.get(status, 0) + 1
        active_rows += 1
        units += qty
        returns += abs(_deep_pick_number(item, ("amountReturns", "returnAmount", "returnedAmount", "quantityReturns")) or 0.0)

        gross = _finance_gross_revenue(item)
        comm = _finance_commission(item)
        logi = _finance_logistics(item)
        payout_direct = _finance_payout_direct(item)
        payout = payout_direct if payout_direct is not None else max(0.0, gross - comm - logi)

        revenue += gross
        commission += comm
        logistics += logi
        payout_total += max(0.0, payout)
        withdrawn += _finance_withdrawn(item)
    return {
        "rows": active_rows,
        "units": units,
        "returns": returns,
        "revenue": revenue,
        "commission": commission,
        "logistics": logistics,
        "payout_total": payout_total,
        "withdrawn": withdrawn,
        "left_to_withdraw": max(0.0, payout_total - withdrawn),
        "statuses": statuses,
    }


def _format_noorza_today(shop_id: int, stats: dict[str, Any], rows: list[dict[str, Any]]) -> str:
    extra = ""
    if not rows:
        extra = (
            "\n\n<i>Пока продаж за выбранный период не найдено. "
            "Если продажа только появилась в кабинете, она может отобразиться чуть позже.</i>"
        )
    return (
        "💰 <b>Продажи за сегодня</b>\n"
        f"🏪 Магазин: <code>{shop_id}</code>\n\n"
        f"🛒 Продаж: <b>{int(stats['rows'])}</b>\n"
        f"📦 Товаров продано: <b>{float(stats['units']):.0f} шт.</b>\n"
        f"↩️ Возвратов: <b>{float(stats['returns']):.0f} шт.</b>\n\n"
        f"💵 Выручка: <b>{_format_money(float(stats['revenue']))}</b>\n"
        f"🏷 Комиссия Uzum: <b>{_format_money(float(stats['commission']))}</b>\n"
        f"🚚 Логистика: <b>{_format_money(float(stats['logistics']))}</b>\n\n"
        f"✅ К выплате: <b>{_format_money(float(stats['payout_total']))}</b>\n"
        f"💳 Уже выведено: <b>{_format_money(float(stats['withdrawn']))}</b>\n"
        f"🧾 Остаток к выплате: <b>{_format_money(float(stats['left_to_withdraw']))}</b>" + extra
    )


@dp.message(Command("today"))
async def today_sales(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⌛ Считаю продажи за сегодня...", reply_markup=menu_for_message(message))
    try:
        rows, _, _ = await _load_today_finance_flexible(client, shop_id)
        stats = _build_noorza_today_stats(rows)
        await message.answer(_format_noorza_today(shop_id, stats, rows), reply_markup=menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


def _format_noorza_period(title: str, shop_id: int, stats: dict[str, Any], rows: list[dict[str, Any]]) -> str:
    text = _format_noorza_today(shop_id, stats, rows)
    text = text.replace("💰 <b>Продажи за сегодня</b>", f"💰 <b>{escape(title)}</b>", 1)
    text = text.replace("💰 <b>Продажи Uzum FBO/FBS за сегодня</b>", f"💰 <b>{escape(title)}</b>", 1)
    if not rows:
        text = text.replace("за сегодня", "за выбранный период")
    return text


@dp.message(Command("yesterday"))
async def yesterday_sales(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⌛ Считаю продажи за вчера...", reply_markup=menu_for_message(message))
    try:
        date_from, date_to = _yesterday_range_ms()
        rows, _, _ = await _load_finance_range_flexible(client, shop_id, date_from, date_to)
        stats = _build_noorza_today_stats(rows)
        await message.answer(_format_noorza_period("Продажи Uzum FBO/FBS за вчера", shop_id, stats, rows), reply_markup=menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("week"))
@dp.message(Command("last7"))
async def week_sales(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⌛ Считаю продажи за 7 дней...", reply_markup=menu_for_message(message))
    try:
        date_from, date_to = _last_7_days_range_ms()
        rows, _, _ = await _load_finance_range_flexible(client, shop_id, date_from, date_to)
        stats = _build_noorza_today_stats(rows)
        await message.answer(_format_noorza_period("Продажи Uzum FBO/FBS за 7 дней", shop_id, stats, rows), reply_markup=menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("balance"))
async def balance(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⌛ Считаю баланс за 30 дней...", reply_markup=menu_for_message(message))
    try:
        date_from, date_to = _days_range_ms(30)
        rows, _, _ = await _load_finance_range_flexible(client, shop_id, date_from, date_to)
        stats = _build_noorza_today_stats(rows)
        await message.answer(
            _format_noorza_period("Баланс Uzum FBO за 30 дней", shop_id, stats, rows),
            reply_markup=menu_for_message(message),
        )
    except Exception as e:
        await send_api_error(message, e)


def _product_missing_qty(product: dict[str, Any]) -> int:
    """Количество потерянного товара из ответа Uzum Products API."""
    value = pick(
        product,
        "quantityMissing",
        "missingQuantity",
        "quantityLost",
        "lostQuantity",
        "missing",
    )
    number = _num_from_value(value)
    return int(number or 0)


def _product_available_qty(product: dict[str, Any]) -> int:
    """Доступный остаток товара из ответа Uzum Products API."""
    value = pick(
        product,
        "quantityAvailable",
        "quantityActive",
        "availableQuantity",
        "stock",
        "quantity",
    )
    number = _num_from_value(value)
    return int(number or 0)


def _product_status_text(product: dict[str, Any]) -> str:
    status = product.get("status") or product.get("productStatus") or {}
    if isinstance(status, dict):
        return str(pick(status, "title", "value", "name", "code") or "-")
    return str(status or "-")


def _product_title(product: dict[str, Any]) -> str:
    return str(
        pick(
            product,
            "productTitle",
            "skuFullTitle",
            "skuTitle",
            "title",
            "name",
        )
        or "Без названия"
    )


def _product_sku_text(product: dict[str, Any]) -> str:
    return str(
        pick(
            product,
            "skuFullTitle",
            "skuTitle",
            "skuName",
            "skuId",
            "sku",
            "productId",
        )
        or "-"
    )


def _short_text(value: Any, limit: int = 70) -> str:
    text = " ".join(str(value or "-").split())
    if len(text) <= limit:
        return text
    return text[: limit - 1] + "…"


def _split_long_message(text: str, limit: int = 3900) -> list[str]:
    parts: list[str] = []
    current = ""
    for block in text.split("\n\n"):
        candidate = block if not current else current + "\n\n" + block
        if len(candidate) <= limit:
            current = candidate
        else:
            if current:
                parts.append(current)
            current = block
    if current:
        parts.append(current)
    return parts


def _format_lost_product_line(product: dict[str, Any], idx: int) -> str:
    missing = _product_missing_qty(product)
    available = _product_available_qty(product)
    title = escape(_short_text(_product_title(product)))
    sku = escape(_short_text(_product_sku_text(product), limit=90))
    status = escape(_product_status_text(product))
    price = _pick_number(product, ("price", "sellPrice", "purchasePrice", "oldPrice")) or 0
    approx = missing * price

    line = (
        f"{idx}. <b>{title}</b>\n"
        f"SKU: {sku}\n"
        f"Потеряно: <b>{missing} шт.</b> | Остаток: {available} шт. | Статус: {status}"
    )
    if price:
        line += f"\nЦена: {_format_money(price)} | Примерная сумма: <b>{_format_money(approx)}</b>"
    return line


@dp.message(Command("lost"))
@dp.message(Command("missing"))
async def lost_goods(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    telegram_id = message.from_user.id if message.from_user else 0
    lang = get_user_language(telegram_id)

    await message.answer("⌛ Yo‘qolgan tovarlarni tekshiryapman..." if lang == "uz" else "⌛ Проверяю потерянные товары...", reply_markup=stock_menu_for_message(message))
    try:
        products = await load_products(client, shop_id, max_pages=50, page_size=100)
        products = [
            p
            for p in products
            if isinstance(p, dict)
            and not p.get("archived")
            and _product_missing_qty(p) > 0
        ]
        products.sort(key=lambda p: _product_missing_qty(p), reverse=True)

        if not products:
            text = (
                f"🧭 <b>Yo‘qolgan tovarlar</b>\n🏪 Do‘kon: <code>{shop_id}</code>\n\nYo‘qolgan tovarlar topilmadi."
                if lang == "uz"
                else f"🧭 <b>Потерянные товары</b>\n🏪 Магазин: <code>{shop_id}</code>\n\nПотерянных товаров не найдено."
            )
            await message.answer(text, reply_markup=stock_menu_for_message(message))
            return

        total_missing = sum(_product_missing_qty(p) for p in products)
        approx_value = sum(
            _product_missing_qty(p)
            * (_pick_number(p, ("price", "sellPrice", "purchasePrice", "oldPrice")) or 0)
            for p in products
        )
        title = "🧭 <b>Yo‘qolgan tovarlar</b>" if lang == "uz" else "🧭 <b>Потерянные товары</b>"
        summary = [
            f"🏪 Do‘kon: <code>{shop_id}</code>" if lang == "uz" else f"🏪 Магазин: <code>{shop_id}</code>",
            f"SKU: <b>{len(products)}</b> | Jami yo‘qolgan: <b>{total_missing}</b> dona" if lang == "uz" else f"SKU с потерями: <b>{len(products)}</b> | Всего потеряно: <b>{total_missing}</b> шт.",
            f"Taxminiy summa: <b>{_format_money(approx_value)}</b>" if lang == "uz" else f"Примерная сумма: <b>{_format_money(approx_value)}</b>",
        ]
        items = [_format_lost_product_line(product, idx) for idx, product in enumerate(products, start=1)]
        await send_paginated_list(message, kind="lost", title=title, summary=summary, items=items, section="stock", reply_markup=stock_menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


# --- FBO Invoice / накладные поставки ---
def _extract_list_any(data: Any) -> list[Any]:
    """Достаём список из разных форматов ответа Uzum API."""
    if isinstance(data, list):
        return data
    try:
        items = extract_items(data)
        if isinstance(items, list) and items:
            return items
    except Exception:
        pass
    if isinstance(data, dict):
        for key in (
            "content", "items", "data", "result", "results", "list", "records",
            "invoices", "invoiceList", "productList", "products",
        ):
            value = data.get(key)
            if isinstance(value, list):
                return value
            if isinstance(value, dict):
                nested = _extract_list_any(value)
                if nested:
                    return nested
    return []


def _value_by_path(item: Any, *paths: str) -> Any:
    if not isinstance(item, dict):
        return None
    for path in paths:
        cur: Any = item
        ok = True
        for part in path.split("."):
            if isinstance(cur, dict) and part in cur:
                cur = cur.get(part)
            else:
                ok = False
                break
        if ok and cur not in (None, ""):
            return cur
    return None


def _status_text_any(value: Any) -> str:
    if isinstance(value, dict):
        return str(
            value.get("text")
            or value.get("title")
            or value.get("name")
            or value.get("value")
            or "—"
        )
    if value in (None, ""):
        return "—"
    return str(value)


def _date_text_any(value: Any) -> str:
    if not value:
        return "—"
    text = str(value).strip()
    parsed = _dt_from_db(text)
    if parsed:
        return _fmt_dt(parsed)
    return text[:19].replace("T", " ")


def _num_any(value: Any) -> float:
    n = _num_from_value(value)
    return float(n or 0)


def _fmt_qty(value: Any) -> str:
    n = _num_any(value)
    if abs(n - int(n)) < 0.00001:
        return str(int(n))
    return str(round(n, 2)).rstrip("0").rstrip(".")


async def _request_fbo_invoices(
    client: UzumClient,
    shop_id: int,
    *,
    page: int = 0,
    size: int = 20,
) -> Any:
    params = [("size", int(size)), ("page", int(page))]
    path = f"/v1/shop/{int(shop_id)}/invoice?" + urlencode(params)
    return await client._request("GET", path)


async def _load_fbo_invoices(
    client: UzumClient,
    shop_id: int,
    *,
    max_pages: int = 3,
    page_size: int = 20,
) -> tuple[list[dict[str, Any]], Any | None]:
    rows: list[dict[str, Any]] = []
    first_response: Any | None = None
    for page in range(max_pages):
        data = await _request_fbo_invoices(client, shop_id, page=page, size=page_size)
        if first_response is None:
            first_response = data
        items = _extract_list_any(data)
        if not items:
            break
        rows.extend([x for x in items if isinstance(x, dict)])
        if len(items) < page_size:
            break
    return rows, first_response


async def _request_fbo_invoice_products(
    client: UzumClient,
    shop_id: int,
    invoice_id: int,
) -> Any:
    params = [("invoiceId", int(invoice_id))]
    path = f"/v1/shop/{int(shop_id)}/invoice/products?" + urlencode(params)
    return await client._request("GET", path)


def _invoice_id(item: dict[str, Any]) -> Any:
    return _value_by_path(item, "id", "invoiceId", "invoice.id")


def _invoice_number(item: dict[str, Any]) -> str:
    return str(
        _value_by_path(item, "invoiceNumber", "number", "invoice.number", "deliveryCertificate")
        or _invoice_id(item)
        or "—"
    )


def _invoice_status(item: dict[str, Any]) -> str:
    return _status_text_any(
        _value_by_path(item, "invoiceStatus", "status", "state", "invoiceStatus.value")
    )


def _format_invoice_line(item: dict[str, Any], idx: int) -> str:
    invoice_id = _invoice_id(item)
    number = escape(_short_text(_invoice_number(item), 80))
    status = escape(_short_text(_invoice_status(item), 60))
    created = _date_text_any(_value_by_path(item, "dateCreated", "createdAt", "creationDate"))
    accepted_date = _date_text_any(_value_by_path(item, "dateAccepted", "acceptedAt", "acceptanceDate"))
    time_from = _date_text_any(_value_by_path(item, "timeSlotReservation.timeFrom", "timeFrom"))
    time_to = _date_text_any(_value_by_path(item, "timeSlotReservation.timeTo", "timeTo"))
    total_to_stock = _value_by_path(item, "totalToStock", "quantityToStock", "totalQuantity")
    total_accepted = _value_by_path(item, "totalAccepted", "quantityAccepted", "acceptedQuantity")
    full_price = _value_by_path(item, "fullPrice", "totalPrice", "price")

    lines = [f"{idx}. <b>Накладная №{number}</b>"]
    if invoice_id not in (None, ""):
        lines.append(f"ID: <code>{escape(str(invoice_id))}</code>")
    lines.append(f"Статус: <b>{status}</b>")
    if created != "—":
        lines.append(f"Создана: {escape(created)}")
    if time_from != "—" or time_to != "—":
        lines.append(f"Окно поставки: {escape(time_from)} — {escape(time_to)}")
    if accepted_date != "—":
        lines.append(f"Принята: {escape(accepted_date)}")
    if total_to_stock not in (None, "") or total_accepted not in (None, ""):
        lines.append(f"К поставке: <b>{_fmt_qty(total_to_stock)}</b> шт. | Принято: <b>{_fmt_qty(total_accepted)}</b> шт.")
    if _num_any(full_price):
        lines.append(f"Сумма: <b>{_format_money(_num_any(full_price))}</b>")
    if invoice_id not in (None, ""):
        lines.append(f"Состав: <code>/invoice {escape(str(invoice_id))}</code>")
    return "\n".join(lines)


@dp.message(Command("invoices"))
@dp.message(Command("fbo_invoices"))
async def fbo_invoices(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req

    await message.answer("⌛ Загружаю FBO-накладные поставки...", reply_markup=menu_for_message(message))
    try:
        invoices, first_response = await _load_fbo_invoices(client, shop_id, max_pages=3, page_size=20)
        title = "📄 <b>FBO-накладные поставки</b>"
        if not invoices:
            text = (
                f"{title}\n"
                f"Магазин: <code>{shop_id}</code>\n\n"
                "Накладные не найдены или Uzum API вернул пустой список."
            )
            if first_response is not None:
                text += "\n\nПервые данные API:\n<code>" + escape(compact_json_preview(first_response)) + "</code>"
            await message.answer(text, reply_markup=menu_for_message(message))
            return

        lines = [
            title,
            f"Магазин: <code>{shop_id}</code>",
            f"Найдено: <b>{len(invoices)}</b>",
            "Чтобы посмотреть состав, отправьте <code>/invoice ID</code>. Например: <code>/invoice 123456</code>",
        ]
        for idx, item in enumerate(invoices[:50], start=1):
            lines.append(_format_invoice_line(item, idx))
        if len(invoices) > 50:
            lines.append(f"Показаны первые 50 накладных из {len(invoices)}.")

        for part in _split_long_message("\n\n".join(lines)):
            await message.answer(part, reply_markup=menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


def _format_invoice_product_line(item: dict[str, Any], idx: int) -> str:
    product_title = _value_by_path(item, "productTitle", "title", "product.name")
    sku_title = _value_by_path(item, "skuTitle", "sku.title", "skuName")
    title = escape(_short_text(product_title or sku_title or "Без названия", 90))
    sku_text = escape(_short_text(sku_title or "—", 90))
    item_id = _value_by_path(item, "id", "skuId", "productId")
    to_stock = _value_by_path(item, "quantityToStock", "toStock", "quantity")
    accepted = _value_by_path(item, "quantityAccepted", "accepted", "acceptedQuantity")
    purchase_price = _value_by_path(item, "purchasePrice", "price", "buyPrice")
    diff = _num_any(to_stock) - _num_any(accepted)

    lines = [f"{idx}. <b>{title}</b>"]
    if item_id not in (None, ""):
        lines.append(f"ID/SKU: <code>{escape(str(item_id))}</code>")
    if sku_text != "—":
        lines.append(f"SKU: {sku_text}")
    lines.append(f"По накладной: <b>{_fmt_qty(to_stock)}</b> шт. | Принято: <b>{_fmt_qty(accepted)}</b> шт.")
    if abs(diff) > 0.00001:
        sign = "−" if diff > 0 else "+"
        lines.append(f"Расхождение: <b>{sign}{_fmt_qty(abs(diff))}</b> шт.")
    if _num_any(purchase_price):
        lines.append(f"Закупочная цена: <b>{_format_money(_num_any(purchase_price))}</b>")

    sku_list = _value_by_path(item, "skuForInvoiceDtoList", "skuList", "skus")
    if isinstance(sku_list, list) and sku_list:
        sku_lines: list[str] = []
        for sku in sku_list[:3]:
            if not isinstance(sku, dict):
                continue
            sku_name = escape(_short_text(_value_by_path(sku, "skuTitle", "title", "name") or "SKU", 60))
            sku_to = _value_by_path(sku, "quantityToStock", "quantity", "toStock")
            sku_acc = _value_by_path(sku, "quantityAccepted", "accepted", "acceptedQuantity")
            sku_lines.append(f"• {sku_name}: {_fmt_qty(sku_to)} / принято {_fmt_qty(sku_acc)}")
        if sku_lines:
            if len(sku_list) > 3:
                sku_lines.append(f"• ещё {len(sku_list) - 3} SKU")
            lines.append("Внутри:\n" + "\n".join(sku_lines))
    return "\n".join(lines)


@dp.message(Command("invoice"))
@dp.message(Command("invoice_products"))
async def fbo_invoice_products(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    arg = parse_args(message.text or "")
    if not arg or not arg.split()[0].isdigit():
        await message.answer(
            "📄 <b>Состав FBO-накладной</b>\n\n"
            "Сначала откройте список накладных: <code>/invoices</code>\n"
            "Потом отправьте команду с ID накладной. Например:\n"
            "<code>/invoice 123456</code>",
            reply_markup=menu_for_message(message),
        )
        return
    invoice_id = int(arg.split()[0])

    await message.answer(f"⌛ Загружаю состав накладной <code>{invoice_id}</code>...", reply_markup=menu_for_message(message))
    try:
        data = await _request_fbo_invoice_products(client, shop_id, invoice_id)
        products = [x for x in _extract_list_any(data) if isinstance(x, dict)]
        title = "📦 <b>Состав FBO-накладной</b>"
        if not products:
            await message.answer(
                f"{title}\n"
                f"Магазин: <code>{shop_id}</code>\n"
                f"Накладная ID: <code>{invoice_id}</code>\n\n"
                "Товары не найдены или API вернул пустой состав.\n\n"
                "Ответ API:\n<code>" + escape(compact_json_preview(data)) + "</code>",
                reply_markup=menu_for_message(message),
            )
            return

        total_to_stock = sum(_num_any(_value_by_path(x, "quantityToStock", "toStock", "quantity")) for x in products)
        total_accepted = sum(_num_any(_value_by_path(x, "quantityAccepted", "accepted", "acceptedQuantity")) for x in products)
        diff = total_to_stock - total_accepted
        total_purchase = sum(
            _num_any(_value_by_path(x, "purchasePrice", "price", "buyPrice"))
            * _num_any(_value_by_path(x, "quantityToStock", "toStock", "quantity"))
            for x in products
        )

        lines = [
            title,
            f"Магазин: <code>{shop_id}</code>",
            f"Накладная ID: <code>{invoice_id}</code>",
            f"Позиций: <b>{len(products)}</b>",
            f"По накладной: <b>{_fmt_qty(total_to_stock)}</b> шт.",
            f"Принято: <b>{_fmt_qty(total_accepted)}</b> шт.",
        ]
        if abs(diff) > 0.00001:
            sign = "−" if diff > 0 else "+"
            lines.append(f"Расхождение: <b>{sign}{_fmt_qty(abs(diff))}</b> шт.")
        if total_purchase:
            lines.append(f"Сумма по закупочной цене: <b>{_format_money(total_purchase)}</b>")

        for idx, item in enumerate(products[:80], start=1):
            lines.append(_format_invoice_product_line(item, idx))
        if len(products) > 80:
            lines.append(f"Показаны первые 80 позиций из {len(products)}.")

        for part in _split_long_message("\n\n".join(lines)):
            await message.answer(part, reply_markup=menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("sales"))
async def sales(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⏳ Считаю продажи за сегодня, 7 и 30 дней...", reply_markup=menu_for_message(message))
    try:
        today, _ = await _sales_period_stats(client, shop_id, 1)
        week, _ = await _sales_period_stats(client, shop_id, 7)
        month, _ = await _sales_period_stats(client, shop_id, 30)
        await message.answer(
            "💰 <b>Сводка продаж</b>\n"
            f"Магазин: <code>{shop_id}</code>\n\n"
            + _format_sales_summary_line("Сегодня", today)
            + "\n\n"
            + _format_sales_summary_line("7 дней", week)
            + "\n\n"
            + _format_sales_summary_line("30 дней", month)
            + "\n\nПодробно: <code>/sales_today</code>, <code>/sales_7</code>, <code>/sales_30</code>",
            reply_markup=menu_for_message(message),
        )
    except Exception as e:
        await send_api_error(message, e)


async def _send_sales_details(message: Message, days: int) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    try:
        stats, first = await _sales_period_stats(client, shop_id, days)
        await message.answer(
            _format_sales_details(days, shop_id, stats, first), reply_markup=menu_for_message(message)
        )
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("sales_today"))
async def sales_today(message: Message) -> None:
    await _send_sales_details(message, 1)


@dp.message(Command("sales_7"))
async def sales_7(message: Message) -> None:
    await _send_sales_details(message, 7)


@dp.message(Command("sales_30"))
async def sales_30(message: Message) -> None:
    await _send_sales_details(message, 30)



# --- Общая сводка магазина / заказы по статусам ---
# Дополняет Finance API. Если Finance API вернул 0 продаж, сводка всё равно покажет
# текущие заказы FBS/DBS и состояние остатков.
FBS_STATUS_LABELS: dict[str, str] = {
    "CREATED": "Создан",
    "PACKING": "Сборка",
    "PENDING_DELIVERY": "Ожидает доставки",
    "DELIVERING": "В доставке",
    "DELIVERED": "Доставлен",
    "ACCEPTED_AT_DP": "Принят в ПВЗ/ДП",
    "DELIVERED_TO_CUSTOMER_DELIVERY_POINT": "В пункте выдачи",
    "COMPLETED": "Завершён",
    "CANCELED": "Отменён",
    "RETURNED": "Возврат",
}

# Минимальный набор статусов для сводки, чтобы не ловить 429 Too Many Requests.
# Полную детализацию статусов добавим позже, когда подберём лимиты Uzum API.
FBS_SUMMARY_STATUSES: tuple[str, ...] = (
    "CREATED",
    "PACKING",
    "COMPLETED",
    "CANCELED",
)
ORDER_SUMMARY_REQUEST_DELAY_SECONDS = float(os.getenv("ORDER_SUMMARY_REQUEST_DELAY_SECONDS", "0.45") or "0.45")


def _extract_count(data: Any) -> int:
    if isinstance(data, bool):
        return 0
    if isinstance(data, (int, float)):
        return int(data)
    if isinstance(data, str):
        try:
            return int(float(data))
        except Exception:
            return 0
    if isinstance(data, dict):
        for key in ("payload", "data", "result", "value", "count", "total", "totalElements", "totalAmount"):
            if key in data:
                value = data.get(key)
                if isinstance(value, dict):
                    nested = _extract_count(value)
                    if nested:
                        return nested
                else:
                    number = _num_from_value(value)
                    if number is not None:
                        return int(number)
        # Последняя попытка — ищем первое числовое поле.
        for value in data.values():
            if isinstance(value, (dict, list)):
                nested = _extract_count(value)
                if nested:
                    return nested
            else:
                number = _num_from_value(value)
                if number is not None:
                    return int(number)
    if isinstance(data, list):
        return len(data)
    return 0


async def _fbs_order_count(
    client: UzumClient,
    shop_id: int,
    status: str,
    *,
    date_from_ms: int,
    date_to_ms: int,
) -> int:
    params = [
        ("shopIds", shop_id),
        ("status", status),
        ("dateFrom", date_from_ms),
        ("dateTo", date_to_ms),
    ]
    path = "/v2/fbs/orders/count?" + urlencode(params)
    data = await client._request("GET", path)
    return _extract_count(data)


async def _orders_counts_for_days(client: UzumClient, shop_id: int, days: int) -> dict[str, int]:
    date_from, date_to = _today_range_ms() if days == 1 else _days_range_ms(days)
    counts: dict[str, int] = {}
    for status in FBS_SUMMARY_STATUSES:
        try:
            counts[status] = await _fbs_order_count(
                client, shop_id, status, date_from_ms=date_from, date_to_ms=date_to
            )
        except Exception as e:
            # Если Uzum вернул 429/403 по одному статусу, не валим всю сводку.
            logging.warning("FBS count failed status=%s days=%s: %s", status, days, e)
            counts[status] = 0

        # Маленькая пауза, чтобы не упираться в лимиты Uzum API.
        await asyncio.sleep(max(0.0, ORDER_SUMMARY_REQUEST_DELAY_SECONDS))
    return counts


def _format_orders_counts(title: str, counts: dict[str, int]) -> str:
    useful = {k: v for k, v in counts.items() if v}
    if not useful:
        return f"<b>{escape(title)}</b>\n• Заказов по основным статусам не найдено"
    lines = [f"<b>{escape(title)}</b>"]
    for status, count in useful.items():
        label = FBS_STATUS_LABELS.get(status, status)
        lines.append(f"• {escape(label)}: <b>{count}</b>")
    lines.append(f"• Итого: <b>{sum(useful.values())}</b>")
    return "\n".join(lines)


def _build_stock_stats(rows: list[dict[str, Any]]) -> dict[str, float | int]:
    total_skus = len(rows)
    total_units = 0.0
    fbo_units = 0.0
    fbs_units = 0.0
    low_count = 0
    zero_count = 0
    stock_value = 0.0
    active_count = 0

    for r in rows:
        total = _num_from_value(r.get("total")) or 0.0
        fbo = _num_from_value(r.get("fbo")) or 0.0
        fbs = _num_from_value(r.get("fbs")) or 0.0
        price = _num_from_value(r.get("price")) or 0.0
        status = str(status_display(r.get("status")) if r.get("status") else r.get("status") or "").upper()

        total_units += total
        fbo_units += fbo
        fbs_units += fbs
        stock_value += max(0.0, total) * max(0.0, price)
        if total <= 0:
            zero_count += 1
        elif total <= LOW_STOCK_THRESHOLD:
            low_count += 1
        if "RUN_OUT" not in status and total > 0:
            active_count += 1

    return {
        "total_skus": total_skus,
        "total_units": total_units,
        "fbo_units": fbo_units,
        "fbs_units": fbs_units,
        "low_count": low_count,
        "zero_count": zero_count,
        "stock_value": stock_value,
        "active_count": active_count,
    }


@dp.message(Command("orders_summary"))
async def orders_summary(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⏳ Считаю заказы по статусам...", reply_markup=menu_for_message(message))
    try:
        today = await _orders_counts_for_days(client, shop_id, 1)
        week = await _orders_counts_for_days(client, shop_id, 7)
        month = await _orders_counts_for_days(client, shop_id, 30)
        await message.answer(
            f"📊 <b>Сводка заказов</b>\nМагазин: <code>{shop_id}</code>\n\n"
            + _format_orders_counts("Сегодня", today)
            + "\n\n"
            + _format_orders_counts("7 дней", week)
            + "\n\n"
            + _format_orders_counts("30 дней", month),
            reply_markup=menu_for_message(message),
        )
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("dashboard", "summary", "report"))
async def dashboard(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⏳ Собираю общую сводку магазина...", reply_markup=menu_for_message(message))
    try:
        sales_7, _ = await _sales_period_stats(client, shop_id, 7)
        sales_30, _ = await _sales_period_stats(client, shop_id, 30)
        counts_7 = await _orders_counts_for_days(client, shop_id, 7)
        rows = await load_sku_rows(client, shop_id, max_pages=50)
        st = _build_stock_stats(rows)

        active_orders = sum(counts_7.get(s, 0) for s in ("CREATED", "PACKING"))
        completed_7 = counts_7.get("COMPLETED", 0)
        canceled_7 = counts_7.get("CANCELED", 0)
        returned_7 = counts_7.get("RETURNED", 0)

        await message.answer(
            f"📈 <b>Сводка магазина</b>\n"
            f"Магазин: <code>{shop_id}</code>\n\n"
            f"💰 <b>Продажи Finance API</b>\n"
            f"• 7 дней: <b>{_format_money(float(sales_7['revenue']))}</b> / строк: <b>{sales_7['active_rows']}</b>\n"
            f"• 30 дней: <b>{_format_money(float(sales_30['revenue']))}</b> / строк: <b>{sales_30['active_rows']}</b>\n\n"
            f"🛒 <b>Заказы FBS/DBS за 7 дней</b>\n"
            f"• Активные статусы: <b>{active_orders}</b>\n"
            f"• Завершено: <b>{completed_7}</b>\n"
            f"• Отменено: <b>{canceled_7}</b>\n"
            f"• Возвраты: <b>{returned_7}</b>\n\n"
            f"📦 <b>Остатки</b>\n"
            f"• SKU: <b>{int(st['total_skus'])}</b>\n"
            f"• Общий остаток: <b>{float(st['total_units']):.0f}</b> шт\n"
            f"• FBO: <b>{float(st['fbo_units']):.0f}</b> шт / FBS: <b>{float(st['fbs_units']):.0f}</b> шт\n"
            f"• Заканчиваются ≤ {LOW_STOCK_THRESHOLD}: <b>{int(st['low_count'])}</b>\n"
            f"• Нет в наличии: <b>{int(st['zero_count'])}</b>\n"
            f"• Примерная стоимость остатка: <b>{_format_money(float(st['stock_value']))}</b>\n\n"
            f"Подробно: <code>/sales</code>, <code>/orders_summary</code>, <code>/lowstock</code>",
            reply_markup=menu_for_message(message),
        )
    except Exception as e:
        await send_api_error(message, e)



# --- Отзывы покупателей ---
# В официальном Seller OpenAPI Uzum раздел отзывов может быть недоступен.
# Поэтому список/ответы сделаны безопасно: сначала используются переменные окружения
# REVIEWS_LIST_PATH и REVIEWS_REPLY_PATH, а если их нет — бот пробует несколько
# типовых путей и честно сообщает ошибку, если endpoint не найден.
REVIEWS_LIST_PATH = os.getenv("REVIEWS_LIST_PATH", "").strip()
REVIEWS_REPLY_PATH = os.getenv("REVIEWS_REPLY_PATH", "").strip()
REVIEWS_REPLY_METHOD = os.getenv("REVIEWS_REPLY_METHOD", "POST").strip().upper() or "POST"
REVIEWS_REPLY_BODY_FIELD = os.getenv("REVIEWS_REPLY_BODY_FIELD", "text").strip() or "text"


def _format_endpoint_template(template: str, *, shop_id: int, review_id: str = "", page: int = 0, size: int = 10) -> str:
    return (
        template.replace("{shop_id}", str(shop_id))
        .replace("{shopId}", str(shop_id))
        .replace("{review_id}", str(review_id))
        .replace("{reviewId}", str(review_id))
        .replace("{page}", str(page))
        .replace("{size}", str(size))
    )


def _review_candidates(shop_id: int, page: int = 0, size: int = 10) -> list[str]:
    if REVIEWS_LIST_PATH:
        return [_format_endpoint_template(REVIEWS_LIST_PATH, shop_id=shop_id, page=page, size=size)]
    return [
        f"/v1/reviews?shopId={shop_id}&page={page}&size={size}",
        f"/v1/reviews/shop/{shop_id}?page={page}&size={size}",
        f"/v1/shop/{shop_id}/reviews?page={page}&size={size}",
        f"/v1/feedbacks?shopId={shop_id}&page={page}&size={size}",
        f"/v1/shop/{shop_id}/feedbacks?page={page}&size={size}",
        f"/v1/comments?shopId={shop_id}&page={page}&size={size}",
        f"/v1/shop/{shop_id}/comments?page={page}&size={size}",
    ]


def _reply_candidates(shop_id: int, review_id: str) -> list[str]:
    if REVIEWS_REPLY_PATH:
        return [_format_endpoint_template(REVIEWS_REPLY_PATH, shop_id=shop_id, review_id=review_id)]
    return [
        f"/v1/reviews/{review_id}/reply",
        f"/v1/reviews/{review_id}/answer",
        f"/v1/review/{review_id}/reply",
        f"/v1/feedbacks/{review_id}/reply",
        f"/v1/feedback/{review_id}/answer",
        f"/v1/shop/{shop_id}/reviews/{review_id}/reply",
        f"/v1/shop/{shop_id}/feedbacks/{review_id}/reply",
        "/v1/reviews/reply",
        "/v1/feedbacks/reply",
    ]


def _find_review_lists(obj: Any) -> list[Any]:
    """Best-effort поиск списков отзывов в неизвестной структуре ответа."""
    direct = extract_items(obj)
    if direct:
        return direct
    found: list[Any] = []
    if isinstance(obj, dict):
        for key, value in obj.items():
            key_l = str(key).lower()
            if isinstance(value, list) and any(x in key_l for x in ("review", "feedback", "comment", "rating")):
                found.extend(value)
            elif isinstance(value, (dict, list)):
                found.extend(_find_review_lists(value))
    elif isinstance(obj, list):
        for value in obj:
            if isinstance(value, dict):
                found.append(value)
            elif isinstance(value, (dict, list)):
                found.extend(_find_review_lists(value))
    return found


def _review_id(review: Any) -> str:
    value = pick(review, "id", "reviewId", "feedbackId", "commentId", "uuid", "uid", default="—")
    return str(value)


def _review_answer(review: Any) -> str:
    return str(pick(review, "answer", "reply", "sellerAnswer", "sellerReply", "response", "commentAnswer", default="—"))


def format_review_line(review: Any) -> str:
    review_id = _review_id(review)
    product = pick(review, "productTitle", "productName", "title", "name", "skuTitle", default="—")
    rating = pick(review, "rating", "stars", "mark", "grade", "score", default="—")
    author = pick(review, "customerName", "userName", "buyerName", "clientName", "author", default="—")
    created = pick(review, "createdAt", "date", "createdDate", "publishedAt", default="—")
    text = pick(review, "text", "comment", "review", "content", "message", "description", default="—")
    answer = _review_answer(review)

    text_s = safe(text)
    if len(text_s) > 600:
        text_s = text_s[:600] + "..."

    answer_part = ""
    if answer not in (None, "", "—"):
        answer_s = safe(answer)
        if len(answer_s) > 300:
            answer_s = answer_s[:300] + "..."
        answer_part = f"\n💬 Ответ продавца: {answer_s}"

    return (
        f"• ID отзыва: <code>{safe(review_id)}</code>\n"
        f"Товар: {safe(product)}\n"
        f"Оценка: {safe(rating)} | Клиент: {safe(author)} | Дата: {safe(created)}\n"
        f"Отзыв: {text_s}"
        f"{answer_part}"
    )


async def get_reviews_from_uzum(client: UzumClient, shop_id: int, *, page: int = 0, size: int = 10) -> tuple[list[Any], str | None, str | None]:
    last_error: str | None = None
    last_path: str | None = None
    for path in _review_candidates(shop_id, page=page, size=size):
        try:
            data = await client._request("GET", path)
            items = _find_review_lists(data)
            return items, path, None
        except Exception as e:
            last_error = str(e)
            last_path = path
            continue
    return [], last_path, last_error


@dp.message(Command("reviews"))
async def reviews(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req

    items, path, error = await get_reviews_from_uzum(client, shop_id, page=0, size=10)
    if error and not items:
        await message.answer(
            "⭐ <b>Отзывы</b>\n\n"
            "Не смог получить отзывы через текущий Uzum API.\n"
            "Скорее всего, в вашем Seller OpenAPI нет открытого метода для отзывов.\n\n"
            "Что можно сделать:\n"
            "1. Если у вас есть endpoint отзывов из кабинета Uzum, добавьте в bothost переменные:\n"
            "<code>REVIEWS_LIST_PATH</code> и <code>REVIEWS_REPLY_PATH</code>.\n"
            "2. Потом перезапустите бота.\n\n"
            f"Последний путь: <code>{escape(str(path))}</code>\n"
            f"Ошибка: <code>{escape(error[:1000])}</code>",
            reply_markup=menu_for_message(message),
        )
        return

    if not items:
        await message.answer("⭐ Отзывы не найдены.", reply_markup=menu_for_message(message))
        return

    lines = [format_review_line(item) for item in items[:10]]
    await message.answer(
        "⭐ <b>Последние отзывы</b>\n\n"
        + "\n\n".join(lines)
        + "\n\nЧтобы ответить: <code>/reply ID_ОТЗЫВА ваш ответ</code>",
        reply_markup=menu_for_message(message),
    )


@dp.message(Command("reply"))
async def reply_review(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req

    arg = parse_args(message.text or "")
    if not arg or " " not in arg:
        await message.answer(
            "Напишите так:\n"
            "<code>/reply ID_ОТЗЫВА Спасибо за отзыв!</code>\n\n"
            "ID отзыва можно посмотреть через <code>/reviews</code>.",
            reply_markup=menu_for_message(message),
        )
        return

    review_id, answer_text = arg.split(maxsplit=1)
    review_id = review_id.strip()
    answer_text = answer_text.strip()
    if not review_id or not answer_text:
        await message.answer("Не вижу ID отзыва или текст ответа.", reply_markup=menu_for_message(message))
        return
    if len(answer_text) > 1000:
        await message.answer("Ответ слишком длинный. Сделайте до 1000 символов.", reply_markup=menu_for_message(message))
        return

    payloads = [
        {REVIEWS_REPLY_BODY_FIELD: answer_text},
        {"text": answer_text},
        {"reply": answer_text},
        {"answer": answer_text},
        {"message": answer_text},
        {"reviewId": review_id, "shopId": shop_id, "text": answer_text},
        {"feedbackId": review_id, "shopId": shop_id, "answer": answer_text},
    ]

    errors: list[str] = []
    tried = 0
    for path in _reply_candidates(shop_id, review_id):
        # Для кастомного REVIEWS_REPLY_PATH пробуем только первый payload, заданный полем REVIEWS_REPLY_BODY_FIELD.
        selected_payloads = payloads[:1] if REVIEWS_REPLY_PATH else payloads
        for payload in selected_payloads:
            tried += 1
            try:
                await client._request(REVIEWS_REPLY_METHOD, path, json=payload)
                await message.answer(
                    "✅ Ответ на отзыв отправлен.\n\n"
                    f"ID отзыва: <code>{escape(review_id)}</code>",
                    reply_markup=menu_for_message(message),
                )
                return
            except Exception as e:
                errors.append(f"{path}: {str(e)[:250]}")
                continue

    await message.answer(
        "⚠️ Не получилось отправить ответ на отзыв.\n\n"
        "Вероятно, ваш текущий Uzum Seller OpenAPI не поддерживает ответы на отзывы, "
        "или нужен точный endpoint из кабинета продавца.\n\n"
        "Можно добавить в bothost:\n"
        "<code>REVIEWS_REPLY_PATH=/точный/путь/{review_id}/reply</code>\n"
        "<code>REVIEWS_REPLY_BODY_FIELD=text</code>\n\n"
        f"Попыток: <b>{tried}</b>\n"
        f"Последняя ошибка:\n<code>{escape(errors[-1] if errors else '—')}</code>",
        reply_markup=menu_for_message(message),
    )


@dp.message(Command("subscribe"))
async def subscribe(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    ensure_subscription(telegram_id)
    lang = get_user_language(telegram_id)
    if lang == "uz":
        text = (
            "💎 <b>Uzum Seller Assistant obunasi</b>\n\n"
            "Nimalar kiradi:\n"
            "✅ bugun, kecha, 7 va 30 kunlik FBO/FBS savdolar\n"
            "✅ qoldiqlar va tugab borayotgan tovarlar\n"
            "✅ Uzum API bergan bo‘lsa, yo‘qolgan tovarlar\n"
            "✅ yangi savdolar haqida xabarlar\n"
            "✅ bir nechta do‘kon bilan ishlash\n"
            "✅ Excel hisobot va ertalabki hisobot\n\n"
            f"🎁 Trial: yangi foydalanuvchi uchun <b>{TRIAL_DAYS} kun</b>\n\n"
            "💰 <b>Tariflar</b>\n"
            f"{escape(SUBSCRIPTION_PLANS_TEXT)}\n\n"
            f"To‘lov uchun administratorga yozing: <b>{admin_contact_text()}</b>\n"
            f"{escape(PAYMENT_TEXT)}\n\n"
            "Chek tekshirilgach, administrator kirishni uzaytiradi.\n"
            "Holatni tekshirish: <code>/my_subscription</code>"
        )
    else:
        text = (
            "💎 <b>Подписка Uzum Seller Assistant</b>\n\n"
            "Что входит:\n"
            "✅ продажи FBO/FBS за сегодня, вчера, 7 и 30 дней\n"
            "✅ остатки и товары, которые заканчиваются\n"
            "✅ потерянные товары, если Uzum отдаёт их в API\n"
            "✅ уведомления о новых продажах\n"
            "✅ работа с несколькими магазинами\n"
            "✅ Excel-отчёт и утренний отчёт\n\n"
            f"🎁 Trial: <b>{TRIAL_DAYS} дня</b> для нового пользователя\n\n"
            "💰 <b>Тарифы</b>\n"
            f"{escape(SUBSCRIPTION_PLANS_TEXT)}\n\n"
            f"Для оплаты напишите администратору: <b>{admin_contact_text()}</b>\n"
            f"{escape(PAYMENT_TEXT)}\n\n"
            "После проверки чекa администратор продлит доступ.\n"
            "Проверить статус: <code>/my_subscription</code>"
        )
    await message.answer(text, reply_markup=admin_contact_markup() or menu_for_message(message))


@dp.message(Command("video", "api_video", "instruction"))
async def video_instruction(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    if lang == "uz":
        text = (
            "🎥 <b>API-kalitni ulash bo‘yicha video</b>\n\n"
            "Videoda qisqa ko‘rsatilgan:\n"
            "1. Uzum Seller kabinetida API kalit qayerda joylashgan.\n"
            "2. Yangi kalit qanday yaratiladi.\n"
            "3. Kalit botga <code>/connect</code> orqali qanday ulanadi.\n\n"
            "Videoni ko‘rish uchun pastdagi tugmani bosing 👇"
        )
    else:
        text = (
            "🎥 <b>Видеоинструкция по подключению API</b>\n\n"
            "В видео коротко показано:\n"
            "1. Где в кабинете Uzum Seller находятся ключи API.\n"
            "2. Как создать новый ключ.\n"
            "3. Как подключить ключ к боту через <code>/connect</code>.\n\n"
            "Нажмите кнопку ниже, чтобы открыть видео 👇"
        )
    await message.answer(text, reply_markup=video_instruction_markup(lang) or menu_for_message(message))


@dp.message(Command("api_token", "token_help", "how_token"))
async def api_token_help(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    if lang == "uz":
        text = (
            "🔑 <b>Uzum API-kalitini botga ulash</b>\n\n"
            "🎥 Videoqo‘llanma: <code>/video</code>\n\n"
            "API-kalitni faqat Uzum Seller kabinetidan olasiz.\n"
            "Bu kabinet paroli emas, uni istalgan vaqtda o‘chirishingiz mumkin.\n\n"
            "API-kalit qayerda:\n"
            "1. Yuqori o‘ng burchakdagi profil / avatarni bosing.\n"
            "2. <b>Mening profilim</b> bo‘limini oching.\n"
            "3. <b>API kalitlari</b> ni bosing.\n"
            "4. <b>Kalit yaratish</b> ni bosing.\n"
            "5. API-kalitni nusxa oling.\n"
            "6. Botga qayting va <code>/connect</code> ni bosing.\n"
            "7. Kalitni bitta xabar qilib yuboring.\n\n"
            "⚠️ API-kalitni begonalarga yubormang. Bot uni himoyalangan ko‘rinishda saqlaydi va xabarni o‘chirishga harakat qiladi."
        )
    else:
        text = (
            "🔑 <b>Как подключить Uzum API к боту</b>\n\n"
            "🎥 Видеоинструкция: <code>/video</code>\n\n"
            "API-ключ создаётся только в вашем кабинете Uzum Seller.\n"
            "Это не пароль от кабинета, ключ можно удалить в любой момент.\n\n"
            "Где взять API-ключ:\n"
            "1. Нажмите на профиль / аватарку в правом верхнем углу.\n"
            "2. Откройте <b>Мой профиль</b>.\n"
            "3. Нажмите <b>Ключи API</b>.\n"
            "4. Нажмите <b>Создать ключ</b>.\n"
            "5. Скопируйте API-ключ.\n"
            "6. Вернитесь в бот и нажмите <code>/connect</code>.\n"
            "7. Отправьте ключ одним сообщением.\n\n"
            "⚠️ Не отправляйте ключ посторонним. Бот хранит его защищённо и старается удалить сообщение с ключом после проверки."
        )
    await message.answer(text, reply_markup=video_instruction_markup(lang) or menu_for_message(message))


@dp.message(Command("security", "privacy"))
async def security(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    if lang == "uz":
        text = (
            "🔐 <b>API-kalit xavfsizligi</b>\n\n"
            "Uzum API-kalitingiz botda ko‘rsatilmaydi va sizga qayta xabar qilib yuborilmaydi.\n"
            "Ulangandan keyin bot kalit yuborilgan xabarni o‘chirishga harakat qiladi.\n"
            "Bazaga faqat himoyalangan versiya saqlanadi.\n\n"
            "Istalgan vaqtda ulanishni <code>/disconnect</code> orqali o‘chirishingiz mumkin.\n"
            "Kalitni almashtirish uchun <code>/reconnect</code> dan foydalaning."
        )
    else:
        text = (
            "🔐 <b>Безопасность API-ключа</b>\n\n"
            "Ваш Uzum API-ключ не показывается в боте и не отправляется обратно сообщением.\n"
            "После подключения бот старается удалить сообщение, где был отправлен ключ.\n"
            "В базе хранится только защищённая версия ключа.\n\n"
            "Вы можете в любой момент удалить подключение командой <code>/disconnect</code>.\n"
            "Чтобы заменить ключ, используйте <code>/reconnect</code>."
        )
    await message.answer(text, reply_markup=menu_for_message(message))


@dp.message(Command("support"))
async def support(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    user = db.get_user(telegram_id)
    shops = db.list_shops(telegram_id)
    connected = bool(user and user["uzum_token_encrypted"])
    if lang == "uz":
        text = (
            "🛟 <b>Uzum Seller Assistant yordami</b>\n\n"
            f"Telegram ID: <code>{telegram_id}</code>\n"
            f"Uzum API: {'✅ ulangan' if connected else '❌ ulanmagan'}\n"
            f"Topilgan do‘konlar: <b>{len(shops)}</b>\n"
            f"Obuna: {subscription_status_text(telegram_id)}\n\n"
            "Agar bot ma’lumotlarni ko‘rsatmasa, tekshiring:\n"
            "1. API-kalit Uzum Seller kabinetida faol.\n"
            "2. Kalit kerakli do‘konga ruxsatga ega.\n"
            "3. Tanlangan davr bo‘yicha Uzum kabinetida savdolar bor.\n"
            "4. API-kalitni o‘zgartirgan bo‘lsangiz — <code>/reconnect</code> ni bosing.\n\n"
            f"Administrator bilan bog‘lanish: <b>{admin_contact_text()}</b>"
        )
    else:
        text = (
            "🛟 <b>Поддержка Uzum Seller Assistant</b>\n\n"
            f"Ваш Telegram ID: <code>{telegram_id}</code>\n"
            f"Uzum API: {'✅ подключён' if connected else '❌ не подключён'}\n"
            f"Магазинов найдено: <b>{len(shops)}</b>\n"
            f"Подписка: {subscription_status_text(telegram_id)}\n\n"
            "Если бот не показывает данные, проверьте:\n"
            "1. API-ключ активен в кабинете Uzum Seller.\n"
            "2. У ключа есть доступ к нужному магазину.\n"
            "3. В кабинете Uzum есть продажи за выбранный период.\n"
            "4. Если меняли API-ключ — нажмите <code>/reconnect</code>.\n\n"
            f"Связаться с администратором: <b>{admin_contact_text()}</b>"
        )
    await message.answer(text, reply_markup=admin_contact_markup() or menu_for_message(message))


@dp.message(Command("my_payments"))
async def my_payments(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    rows = list_payments(telegram_id, 10)
    if not rows:
        await message.answer("💳 История оплат пока пустая.", reply_markup=menu_for_message(message))
        return
    await message.answer(
        "💳 <b>Мои оплаты</b>\n\n" + "\n".join(payment_line(row) for row in rows),
        reply_markup=menu_for_message(message),
    )

@dp.message(Command("my_subscription", "subscription"))
async def my_subscription(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    await message.answer(subscription_full_text(telegram_id), reply_markup=menu_for_message(message))


@dp.message(Command("admin"))
async def admin_panel(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        await message.answer("⛔ Админ-панель доступна только владельцу бота.", reply_markup=menu_for_message(message))
        return
    init_business_tables()
    stats = get_admin_stats()
    money_today = f"{stats['payments_today']:,}".replace(",", " ")
    money_30 = f"{stats['payments_30']:,}".replace(",", " ")
    await message.answer(
        "👑 <b>Админ-панель</b>\n\n"
        f"👥 Пользователей всего: <b>{stats['total_users']}</b>\n"
        f"🔑 Подключили Uzum API: <b>{stats['connected']}</b>\n"
        f"✅ Активных доступов: <b>{stats['active']}</b>\n"
        f"💳 Платных: <b>{stats['paid']}</b>\n"
        f"🎁 Trial: <b>{stats['trial']}</b>\n"
        f"⛔ Истекли: <b>{stats['expired']}</b>\n"
        f"🚫 Заблокированы: <b>{stats['blocked']}</b>\n\n"
        f"💰 Оплаты сегодня: <b>{money_today}</b> сум\n"
        f"💰 Оплаты за 30 дней: <b>{money_30}</b> сум\n\n"
        "Быстрые команды:\n"
        "• <code>/paid1 ID</code> — 1 месяц / 250 000 сум\n"
        "• <code>/paid3 ID</code> — 3 месяца / 650 000 сум\n"
        "• <code>/paid6 ID</code> — 6 месяцев / 1 200 000 сум\n"
        "• <code>/expiring</code> — кто скоро заканчивается\n"
        "• <code>/backup_db</code> — скачать базу",
        reply_markup=admin_menu_for_message(message),
    )


@dp.message(Command("check"))
async def check_connection(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    ensure_subscription(telegram_id)
    user = db.get_user(telegram_id)
    client = get_uzum_for_user(telegram_id)
    shop_id = db.get_default_shop_id(telegram_id)
    lines = [
        "✅ <b>Проверка подключения</b>",
        f"Telegram ID: <code>{telegram_id}</code>",
        f"Подписка: {subscription_status_text(telegram_id)}",
        f"Uzum API: {'✅ подключён' if client else '❌ не подключён'}",
        f"Активный магазин: {f'<code>{shop_id}</code>' if shop_id else '—'}",
    ]
    if client is None:
        lines.append("\nЧто делать: нажмите <code>/connect</code> и отправьте Uzum API-ключ.")
        await message.answer("\n".join(lines), reply_markup=menu_for_message(message))
        return
    try:
        data = await client.get_shops()
        shops = extract_items(data)
        encrypted = db.get_encrypted_token(telegram_id)
        if encrypted and shops:
            db.save_connection(telegram_id, encrypted, shops)
        lines.append(f"Магазинов найдено: <b>{len(shops)}</b>")
    except Exception as e:
        lines.append("Магазины: ❌ ошибка")
        lines.append(f"Причина: <code>{escape(str(e))[:500]}</code>")
        await message.answer("\n".join(lines), reply_markup=menu_for_message(message))
        return
    if shop_id:
        try:
            rows = await load_sku_rows(client, int(shop_id), max_pages=1)
            lines.append(f"Остатки/товары: ✅ доступно, SKU строк: <b>{len(rows)}</b>")
        except Exception as e:
            lines.append("Остатки/товары: ❌ ошибка")
            lines.append(f"Причина: <code>{escape(str(e))[:500]}</code>")
        try:
            date_from, date_to = _today_range_ms()
            sales_rows, _ = await _load_finance_orders(client, int(shop_id), date_from_ms=date_from, date_to_ms=date_to, max_pages=1, page_size=20)
            lines.append(f"Finance API: ✅ доступно, продаж сегодня: <b>{len(sales_rows)}</b>")
        except Exception as e:
            lines.append("Finance API: ❌ ошибка")
            lines.append(f"Причина: <code>{escape(str(e))[:500]}</code>")
    lines.append("\nЕсли здесь всё ✅ — бот готов к работе.")
    await message.answer("\n".join(lines), reply_markup=menu_for_message(message))


@dp.message(Command("expiring"))
async def admin_expiring(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    rows = list_expiring_users(3, 50)
    if not rows:
        await message.answer("⏳ В ближайшие 3 дня подписки не заканчиваются.", reply_markup=admin_menu_for_message(message))
        return
    await message.answer(
        "⏳ <b>Заканчиваются в ближайшие 3 дня</b>\n\n" + "\n".join(subscription_compact_line(r) for r in rows),
        reply_markup=admin_menu_for_message(message),
    )


@dp.message(Command("blocked"))
async def admin_blocked_users(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    rows = list_blocked_users(50)
    if not rows:
        await message.answer("⛔ Заблокированных пользователей нет.", reply_markup=admin_menu_for_message(message))
        return
    await message.answer(
        "⛔ <b>Заблокированные пользователи</b>\n\n" + "\n".join(subscription_compact_line(r) for r in rows),
        reply_markup=admin_menu_for_message(message),
    )


@dp.message(Command("staff_connections"))
async def staff_connections_admin(message: Message) -> None:
    await message.answer(
        "🔑 Подключение через сотрудника отключено. Сейчас используется официальный способ через API-ключ: /connect",
        reply_markup=admin_menu_for_message(message),
    )


@dp.message(F.text == "🔌 Через сотрудника")
@dp.message(F.text == "🔌 Xodim orqali")
async def staff_connections_button(message: Message) -> None:
    await staff_connections_admin(message)


@dp.message(Command("users"))
async def admin_users(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    if not admin_only(telegram_id):
        return
    rows = list_subscription_users(30)
    if not rows:
        await message.answer("Пользователей пока нет.", reply_markup=menu_for_message(message))
        return
    lines = [subscription_compact_line(row) for row in rows]
    await message.answer(
        "👥 <b>Пользователи</b>\n\n"
        + "\n".join(lines)
        + "\n\nКоманды: <code>/extend ID 30</code>, <code>/paid ID сумма дни</code>, <code>/payments</code>",
        reply_markup=menu_for_message(message),
    )


@dp.message(Command("user"))
async def admin_user_info(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "")
    if not arg.split() or not arg.split()[0].isdigit():
        await message.answer("Напишите так: <code>/user TELEGRAM_ID</code>", reply_markup=menu_for_message(message))
        return
    target = int(arg.split()[0])
    row = ensure_subscription(target)
    user = db.get_user(target)
    username_text = f"@{escape(str(user['username']))}" if user and user['username'] else "—"
    shop_text = f"<code>{user['default_shop_id']}</code>" if user and user['default_shop_id'] else "—"
    payments = list_payments(target, 5)
    payments_text = "\n".join(payment_line(p) for p in payments) if payments else "—"
    await message.answer(
        "👤 <b>Пользователь</b>\n\n"
        f"ID: <code>{target}</code>\n"
        f"Username: {username_text}\n"
        f"Uzum API: {'✅ подключён' if user and user['uzum_token_encrypted'] else '❌ не подключён'}\n"
        f"Магазин: {shop_text}\n"
        f"Статус: {subscription_status_text(target)}\n"
        f"Trial до: <b>{_fmt_dt(row.get('trial_until'))}</b>\n"
        f"Оплачено до: <b>{_fmt_dt(row.get('subscription_until'))}</b>\n\n"
        f"💳 Последние оплаты:\n{payments_text}",
        reply_markup=menu_for_message(message),
    )


@dp.message(Command("extend"))
async def admin_extend(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    parts = parse_args(message.text or "").split()
    if len(parts) < 2 or not parts[0].isdigit() or not parts[1].isdigit():
        await message.answer("Напишите так: <code>/extend TELEGRAM_ID 30</code>", reply_markup=menu_for_message(message))
        return
    target = int(parts[0])
    days = int(parts[1])
    new_until = extend_subscription_days(target, days)
    await message.answer(
        f"✅ Доступ продлён для <code>{target}</code> на {days} дней.\n"
        f"Активен до: <b>{_fmt_dt(new_until)}</b>",
        reply_markup=menu_for_message(message),
    )
    try:
        await bot.send_message(
            target,
            f"✅ Ваша подписка продлена на {days} дней.\nАктивна до: <b>{_fmt_dt(new_until)}</b>",
            reply_markup=main_menu_for_user(target),
        )
    except Exception:
        pass




@dp.message(Command("paid"))
async def admin_paid(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    parts = parse_args(message.text or "").split(maxsplit=3)
    if len(parts) < 3 or not parts[0].isdigit() or not parts[1].replace("_", "").isdigit() or not parts[2].isdigit():
        await message.answer(
            "Напишите так: <code>/paid TELEGRAM_ID СУММА ДНИ комментарий</code>\n"
            "Пример: <code>/paid 123456789 250000 30 чек Click</code>",
            reply_markup=menu_for_message(message),
        )
        return
    target = int(parts[0])
    amount = int(parts[1].replace("_", ""))
    days = int(parts[2])
    comment = parts[3] if len(parts) > 3 else "ручная оплата"
    new_until = extend_subscription_days(target, days)
    payment_id = record_payment(target, amount, days, admin_id, comment)
    amount_text = f"{amount:,}".replace(",", " ")
    await message.answer(
        f"✅ Оплата записана #{payment_id}\n"
        f"Пользователь: <code>{target}</code>\n"
        f"Сумма: <b>{amount_text}</b> сум\n"
        f"Продление: <b>{days}</b> дней\n"
        f"Доступ до: <b>{_fmt_dt(new_until)}</b>",
        reply_markup=menu_for_message(message),
    )
    try:
        await bot.send_message(
            target,
            "✅ <b>Оплата подтверждена</b>\n\n"
            f"Подписка продлена на <b>{days}</b> дней.\n"
            f"Доступ активен до: <b>{_fmt_dt(new_until)}</b>\n\n"
            "Спасибо за оплату!",
            reply_markup=main_menu_for_user(target),
        )
    except Exception:
        pass


@dp.message(Command("paid1"))
async def admin_paid_1_month(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split(maxsplit=1)
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/paid1 TELEGRAM_ID</code>", reply_markup=menu_for_message(message))
        return
    target = int(arg[0])
    new_until = extend_subscription_days(target, 30)
    payment_id = record_payment(target, 250000, 30, admin_id, "1 месяц")
    await message.answer(f"✅ Оплата #{payment_id}: 250 000 сум. Доступ для <code>{target}</code> до <b>{_fmt_dt(new_until)}</b>", reply_markup=menu_for_message(message))
    try:
        await bot.send_message(target, f"✅ Оплата подтверждена. Подписка продлена на 1 месяц. Доступ до: <b>{_fmt_dt(new_until)}</b>", reply_markup=main_menu_for_user(target))
    except Exception:
        pass


@dp.message(Command("paid3"))
async def admin_paid_3_months(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split(maxsplit=1)
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/paid3 TELEGRAM_ID</code>", reply_markup=menu_for_message(message))
        return
    target = int(arg[0])
    new_until = extend_subscription_days(target, 90)
    payment_id = record_payment(target, 650000, 90, admin_id, "3 месяца")
    await message.answer(f"✅ Оплата #{payment_id}: 650 000 сум. Доступ для <code>{target}</code> до <b>{_fmt_dt(new_until)}</b>", reply_markup=menu_for_message(message))
    try:
        await bot.send_message(target, f"✅ Оплата подтверждена. Подписка продлена на 3 месяца. Доступ до: <b>{_fmt_dt(new_until)}</b>", reply_markup=main_menu_for_user(target))
    except Exception:
        pass


@dp.message(Command("paid6"))
async def admin_paid_6_months(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split(maxsplit=1)
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/paid6 TELEGRAM_ID</code>", reply_markup=menu_for_message(message))
        return
    target = int(arg[0])
    new_until = extend_subscription_days(target, 180)
    payment_id = record_payment(target, 1200000, 180, admin_id, "6 месяцев")
    await message.answer(f"✅ Оплата #{payment_id}: 1 200 000 сум. Доступ для <code>{target}</code> до <b>{_fmt_dt(new_until)}</b>", reply_markup=menu_for_message(message))
    try:
        await bot.send_message(target, f"✅ Оплата подтверждена. Подписка продлена на 6 месяцев. Доступ до: <b>{_fmt_dt(new_until)}</b>", reply_markup=main_menu_for_user(target))
    except Exception:
        pass


@dp.message(Command("payments"))
async def admin_payments(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    args = parse_args(message.text or "").split()
    target = int(args[0]) if args and args[0].isdigit() else None
    rows = list_payments(target, 20)
    if not rows:
        await message.answer("💳 История оплат пока пустая.", reply_markup=menu_for_message(message))
        return
    title = f"💳 <b>Оплаты пользователя <code>{target}</code></b>" if target else "💳 <b>Последние оплаты</b>"
    await message.answer(title + "\n\n" + "\n".join(payment_line(row) for row in rows), reply_markup=menu_for_message(message))


@dp.message(Command("backup_db"))
async def admin_backup_db(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    path = Path(DB_PATH)
    if not path.exists():
        await message.answer(f"❌ База не найдена: <code>{escape(str(path))}</code>", reply_markup=menu_for_message(message))
        return
    await message.answer("📦 Отправляю резервную копию базы. Храните файл аккуратно — там данные пользователей.", reply_markup=menu_for_message(message))
    try:
        await message.answer_document(FSInputFile(str(path), filename=f"bot_backup_{datetime.now(UZT).strftime('%Y%m%d_%H%M')}.db"))
    except Exception as e:
        await send_api_error(message, e)

@dp.message(Command("trial"))
async def admin_trial(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    parts = parse_args(message.text or "").split()
    if len(parts) < 2 or not parts[0].isdigit() or not parts[1].isdigit():
        await message.answer("Напишите так: <code>/trial TELEGRAM_ID 3</code>", reply_markup=menu_for_message(message))
        return
    target = int(parts[0])
    days = int(parts[1])
    new_until = set_trial_days(target, days)
    await message.answer(f"🎁 Trial для <code>{target}</code> до <b>{_fmt_dt(new_until)}</b>", reply_markup=menu_for_message(message))


@dp.message(Command("block"))
async def admin_block(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split()
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/block TELEGRAM_ID</code>", reply_markup=menu_for_message(message))
        return
    target = int(arg[0])
    set_blocked(target, True)
    await message.answer(f"⛔ Пользователь <code>{target}</code> заблокирован.", reply_markup=menu_for_message(message))


@dp.message(Command("unblock"))
async def admin_unblock(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split()
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/unblock TELEGRAM_ID</code>", reply_markup=menu_for_message(message))
        return
    target = int(arg[0])
    set_blocked(target, False)
    await message.answer(f"✅ Пользователь <code>{target}</code> разблокирован.", reply_markup=menu_for_message(message))


@dp.message(Command("broadcast"))
async def admin_broadcast(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    text = parse_args(message.text or "")
    if not text:
        await message.answer("Напишите так: <code>/broadcast текст рассылки</code>", reply_markup=menu_for_message(message))
        return
    rows = list_subscription_users(500)
    sent = 0
    for row in rows:
        target = int(row["telegram_id"])
        try:
            await bot.send_message(target, "📢 <b>Сообщение от администратора</b>\n\n" + text, reply_markup=main_menu_for_user(target))
            sent += 1
            await asyncio.sleep(0.05)
        except Exception:
            pass
    await message.answer(f"✅ Рассылка завершена. Отправлено: {sent}", reply_markup=menu_for_message(message))


@dp.message(Command("debug_product"))
async def debug_product(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req

    try:
        data = await client.get_products(shop_id, page=0, size=1)
        items = extract_items(data)
        if not items:
            await message.answer(
                "Товар для debug не найден. Ответ API:\n<code>"
                + escape(compact_json_preview(data, limit=3000))
                + "</code>",
                reply_markup=menu_for_message(message),
            )
            return

        await message.answer(
            "🧪 <b>Первый товар — сырой JSON</b>\n\n<code>"
            + escape(compact_json_preview(items[0], limit=3200))
            + "</code>",
            reply_markup=menu_for_message(message),
        )
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("export_products"))
async def export_products(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req

    try:
        rows = await load_sku_rows(client, shop_id, max_pages=50)
        if not rows:
            await message.answer("SKU-остатки для экспорта не найдены.", reply_markup=menu_for_message(message))
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Stocks"
        ws.append([
            "Product ID",
            "SKU ID",
            "Barcode",
            "Seller code",
            "Category",
            "Product title",
            "SKU title",
            "Price",
            "FBO / склад Uzum",
            "FBS/DBS / склад продавца",
            "Итого доступно",
            "Активно",
            "Продано",
            "Возвраты",
            "Недостача",
            "Брак",
            "Ожидает",
            "Статус",
        ])

        for r in rows:
            ws.append([
                excel_value(r.get("product_id")),
                excel_value(r.get("sku_id")),
                excel_value(r.get("barcode")),
                excel_value(r.get("seller_item_code")),
                excel_value(r.get("category")),
                excel_value(r.get("product_title")),
                excel_value(r.get("sku_full_title") or r.get("sku_title")),
                excel_value(r.get("price")),
                excel_value(r.get("fbo")),
                excel_value(r.get("fbs")),
                excel_value(r.get("total")),
                excel_value(r.get("active")),
                excel_value(r.get("sold")),
                excel_value(r.get("returned")),
                excel_value(r.get("missing")),
                excel_value(r.get("defected")),
                excel_value(r.get("pending")),
                excel_value(status_display(r.get("status")) if r.get("status") else ""),
            ])

        for column in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = min(max(max_len + 2, 12), 70)

        tmp_dir = Path(tempfile.gettempdir())
        filename = tmp_dir / f"uzum_stocks_{shop_id}.xlsx"
        wb.save(filename)
        await message.answer(f"✅ Экспортировано SKU-остатков: {len(rows)}", reply_markup=menu_for_message(message))
        await message.answer_document(FSInputFile(filename))
    except Exception as e:
        await send_api_error(message, e)


# --- Подробный Excel-отчёт ---
def _excel_style_sheet(ws, *, freeze: str = "A2") -> None:
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if freeze:
        ws.freeze_panes = freeze
    try:
        ws.auto_filter.ref = ws.dimensions
    except Exception:
        pass
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def _excel_autowidth(ws, *, min_width: int = 10, max_width: int = 45) -> None:
    for column in ws.columns:
        letter = column[0].column_letter
        max_len = 0
        for cell in column:
            value = cell.value
            if value is None:
                continue
            text = str(value)
            max_len = max(max_len, max(len(line) for line in text.splitlines() or [text]))
        ws.column_dimensions[letter].width = min(max(max_len + 2, min_width), max_width)


def _excel_money_format(ws, columns: tuple[str, ...], start_row: int = 2) -> None:
    for col in columns:
        for cell in ws[col][start_row - 1:]:
            cell.number_format = '#,##0'


def _finance_date_for_excel(item: dict[str, Any]) -> str:
    value = pick(
        item,
        "createdAt",
        "dateCreated",
        "orderDate",
        "date",
        "created",
        "paymentDate",
        "updatedAt",
    )
    if isinstance(value, dict):
        value = pick(value, "date", "value", "createdAt")
    return _date_text_any(value)


def _finance_order_id_for_excel(item: dict[str, Any]) -> str:
    value = pick(
        item,
        "id",
        "orderId",
        "order_id",
        "postingNumber",
        "operationId",
        "financeOrderId",
        "number",
    )
    if value in (None, ""):
        return "—"
    return str(value)


def _finance_sku_for_excel(item: dict[str, Any]) -> str:
    value = pick(
        item,
        "skuTitle",
        "skuName",
        "skuId",
        "barcode",
        "productId",
        "sellerSku",
        "offerId",
    )
    if isinstance(value, dict):
        value = pick(value, "title", "name", "id", "value")
    return str(value or "—")


def _report_periods_ms() -> list[tuple[str, tuple[int, int]]]:
    return [
        ("Сегодня", _today_range_ms()),
        ("Вчера", _yesterday_range_ms()),
        ("7 дней", _last_7_days_range_ms()),
        ("30 дней", _days_range_ms(30)),
    ]


async def _build_full_excel_report(client: UzumClient, shop_id: int) -> Path:
    generated_at = datetime.now(UZT).strftime("%Y-%m-%d_%H-%M")

    # 1) Продажи по периодам
    finance_by_period: dict[str, list[dict[str, Any]]] = {}
    stats_by_period: dict[str, dict[str, Any]] = {}
    for label, (date_from, date_to) in _report_periods_ms():
        rows, _, _ = await _load_finance_range_flexible(client, shop_id, date_from, date_to)
        finance_by_period[label] = rows
        stats_by_period[label] = _build_noorza_today_stats(rows)
        await asyncio.sleep(0.1)

    # 2) Остатки / потерянные
    stock_rows = await load_sku_rows(client, shop_id, max_pages=50)
    low_stock_rows = []
    missing_rows = []
    for r in stock_rows:
        total = _num_any(r.get("total"))
        fbo = _num_any(r.get("fbo"))
        fbs = _num_any(r.get("fbs"))
        missing = _num_any(r.get("missing"))
        active = r.get("active")
        if missing > 0:
            missing_rows.append(r)
        if active is not False and total <= LOW_STOCK_THRESHOLD and (fbo + fbs + total) >= 0:
            low_stock_rows.append(r)

    # 3) FBO накладные и состав первых накладных
    invoices, _ = await _load_fbo_invoices(client, shop_id, max_pages=3, page_size=20)
    invoice_product_rows: list[tuple[Any, dict[str, Any]]] = []
    for item in invoices[:max(0, REPORT_INVOICE_PRODUCT_LIMIT)]:
        invoice_id = _invoice_id(item)
        if invoice_id in (None, ""):
            continue
        try:
            data = await _request_fbo_invoice_products(client, shop_id, int(invoice_id))
            products = [x for x in _extract_list_any(data) if isinstance(x, dict)]
            for p in products:
                invoice_product_rows.append((invoice_id, p))
            await asyncio.sleep(0.1)
        except Exception:
            logging.exception("Excel report: failed to load invoice products for %s", invoice_id)
            continue

    wb = Workbook()

    # Лист 1: Сводка
    ws = wb.active
    ws.title = "Сводка"
    ws.append(["Показатель", "Значение"])
    ws.append(["Магазин", shop_id])
    ws.append(["Дата создания отчёта", datetime.now(UZT).strftime("%d.%m.%Y %H:%M")])
    ws.append(["Период продаж в деталях", "30 дней"])
    ws.append(["SKU в остатках", len(stock_rows)])
    ws.append(["SKU заканчиваются", len(low_stock_rows)])
    ws.append(["SKU с потерями", len(missing_rows)])
    ws.append(["FBO накладных найдено", len(invoices)])
    ws.append(["Состав накладных загружен", f"для первых {min(len(invoices), max(0, REPORT_INVOICE_PRODUCT_LIMIT))} накладных"])
    ws.append([])
    ws.append(["Период", "Позиций", "Товаров, шт", "Возвраты, шт", "Выручка", "Комиссия Uzum", "Логистика", "К выплате", "Уже выведено", "Остаток к выплате", "Статусы"])
    for label in ("Сегодня", "Вчера", "7 дней", "30 дней"):
        st = stats_by_period.get(label, {})
        statuses = st.get("statuses") or {}
        status_text = "; ".join(f"{k}: {v}" for k, v in sorted(statuses.items()))
        ws.append([
            label,
            int(st.get("rows") or 0),
            float(st.get("units") or 0),
            float(st.get("returns") or 0),
            float(st.get("revenue") or 0),
            float(st.get("commission") or 0),
            float(st.get("logistics") or 0),
            float(st.get("payout_total") or 0),
            float(st.get("withdrawn") or 0),
            float(st.get("left_to_withdraw") or 0),
            status_text,
        ])
    _excel_style_sheet(ws)
    _excel_money_format(ws, ("E", "F", "G", "H", "I", "J"), start_row=12)
    _excel_autowidth(ws, max_width=55)

    # Лист 2: Продажи 30 дней
    ws = wb.create_sheet("Продажи 30 дней")
    ws.append(["Дата", "Статус", "ID заказа/операции", "Товар", "SKU/код", "Кол-во", "Выручка", "Комиссия", "Логистика", "К выплате", "Выведено", "Сырой фрагмент"])
    for item in finance_by_period.get("30 дней", []):
        gross = _finance_gross_revenue(item)
        comm = _finance_commission(item)
        logi = _finance_logistics(item)
        direct = _finance_payout_direct(item)
        payout = direct if direct is not None else max(0.0, gross - comm - logi)
        ws.append([
            _finance_date_for_excel(item),
            _finance_status(item),
            _finance_order_id_for_excel(item),
            _finance_title(item),
            _finance_sku_for_excel(item),
            _finance_qty(item),
            gross,
            comm,
            logi,
            max(0.0, payout),
            _finance_withdrawn(item),
            compact_json_preview(item, limit=700),
        ])
    _excel_style_sheet(ws)
    _excel_money_format(ws, ("G", "H", "I", "J", "K"))
    _excel_autowidth(ws, max_width=60)

    # Лист 3: Остатки
    ws = wb.create_sheet("Остатки")
    ws.append(["Product ID", "SKU ID", "Barcode", "Код продавца", "Категория", "Товар", "SKU", "Цена", "FBO", "FBS/DBS", "Итого", "Активно", "Продано", "Возвраты", "Потеряно", "Брак", "Ожидает", "Статус"])
    for r in stock_rows:
        ws.append([
            excel_value(r.get("product_id")), excel_value(r.get("sku_id")), excel_value(r.get("barcode")),
            excel_value(r.get("seller_item_code")), excel_value(r.get("category")), excel_value(r.get("product_title")),
            excel_value(r.get("sku_full_title") or r.get("sku_title")), excel_value(r.get("price")), excel_value(r.get("fbo")),
            excel_value(r.get("fbs")), excel_value(r.get("total")), excel_value(r.get("active")), excel_value(r.get("sold")),
            excel_value(r.get("returned")), excel_value(r.get("missing")), excel_value(r.get("defected")), excel_value(r.get("pending")),
            excel_value(status_display(r.get("status")) if r.get("status") else ""),
        ])
    _excel_style_sheet(ws)
    _excel_money_format(ws, ("H",))
    _excel_autowidth(ws, max_width=55)

    # Лист 4: Заканчивается
    ws = wb.create_sheet("Заканчивается")
    ws.append(["Product ID", "SKU ID", "Товар", "SKU", "Цена", "FBO", "FBS/DBS", "Итого", "Порог", "Статус"])
    for r in sorted(low_stock_rows, key=lambda x: _num_any(x.get("total"))):
        ws.append([
            excel_value(r.get("product_id")), excel_value(r.get("sku_id")), excel_value(r.get("product_title")),
            excel_value(r.get("sku_full_title") or r.get("sku_title")), excel_value(r.get("price")), excel_value(r.get("fbo")),
            excel_value(r.get("fbs")), excel_value(r.get("total")), LOW_STOCK_THRESHOLD,
            excel_value(status_display(r.get("status")) if r.get("status") else ""),
        ])
    _excel_style_sheet(ws)
    _excel_money_format(ws, ("E",))
    _excel_autowidth(ws, max_width=55)

    # Лист 5: Потерянные
    ws = wb.create_sheet("Потерянные")
    ws.append(["Product ID", "SKU ID", "Товар", "SKU", "Цена", "Потеряно", "Примерная сумма", "FBO", "FBS/DBS", "Итого", "Статус"])
    for r in sorted(missing_rows, key=lambda x: _num_any(x.get("missing")), reverse=True):
        price = _num_any(r.get("price"))
        missing = _num_any(r.get("missing"))
        ws.append([
            excel_value(r.get("product_id")), excel_value(r.get("sku_id")), excel_value(r.get("product_title")),
            excel_value(r.get("sku_full_title") or r.get("sku_title")), price, missing, price * missing,
            excel_value(r.get("fbo")), excel_value(r.get("fbs")), excel_value(r.get("total")),
            excel_value(status_display(r.get("status")) if r.get("status") else ""),
        ])
    _excel_style_sheet(ws)
    _excel_money_format(ws, ("E", "G"))
    _excel_autowidth(ws, max_width=55)

    # Лист 6: FBO накладные
    ws = wb.create_sheet("FBO накладные")
    ws.append(["ID", "Номер", "Статус", "Создана", "Окно от", "Окно до", "Принята", "К поставке", "Принято", "Расхождение", "Сумма"])
    for item in invoices:
        to_stock = _num_any(_value_by_path(item, "totalToStock", "quantityToStock", "totalQuantity"))
        accepted = _num_any(_value_by_path(item, "totalAccepted", "quantityAccepted", "acceptedQuantity"))
        ws.append([
            excel_value(_invoice_id(item)),
            excel_value(_invoice_number(item)),
            excel_value(_invoice_status(item)),
            excel_value(_date_text_any(_value_by_path(item, "dateCreated", "createdAt", "creationDate"))),
            excel_value(_date_text_any(_value_by_path(item, "timeSlotReservation.timeFrom", "timeFrom"))),
            excel_value(_date_text_any(_value_by_path(item, "timeSlotReservation.timeTo", "timeTo"))),
            excel_value(_date_text_any(_value_by_path(item, "dateAccepted", "acceptedAt", "acceptanceDate"))),
            to_stock,
            accepted,
            to_stock - accepted,
            _num_any(_value_by_path(item, "fullPrice", "totalPrice", "price")),
        ])
    _excel_style_sheet(ws)
    _excel_money_format(ws, ("K",))
    _excel_autowidth(ws, max_width=55)

    # Лист 7: Состав накладных
    ws = wb.create_sheet("Состав накладных")
    ws.append(["Invoice ID", "Product/SKU ID", "Товар", "SKU", "По накладной", "Принято", "Расхождение", "Закупочная цена", "Сумма по накладной"])
    for invoice_id, item in invoice_product_rows:
        to_stock = _num_any(_value_by_path(item, "quantityToStock", "toStock", "quantity"))
        accepted = _num_any(_value_by_path(item, "quantityAccepted", "accepted", "acceptedQuantity"))
        price = _num_any(_value_by_path(item, "purchasePrice", "price", "buyPrice"))
        ws.append([
            excel_value(invoice_id),
            excel_value(_value_by_path(item, "id", "skuId", "productId")),
            excel_value(_value_by_path(item, "productTitle", "title", "product.name")),
            excel_value(_value_by_path(item, "skuTitle", "sku.title", "skuName")),
            to_stock,
            accepted,
            to_stock - accepted,
            price,
            price * to_stock,
        ])
    _excel_style_sheet(ws)
    _excel_money_format(ws, ("H", "I"))
    _excel_autowidth(ws, max_width=55)

    tmp_dir = Path(tempfile.gettempdir())
    filename = tmp_dir / f"uzum_full_report_{shop_id}_{generated_at}.xlsx"
    wb.save(filename)
    return filename


@dp.message(Command("report_excel"))
@dp.message(Command("report"))
@dp.message(Command("full_report"))
async def report_excel(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req

    await message.answer(
        "⌛ Готовлю подробный Excel-отчёт...\n"
        "Это может занять 20–60 секунд: собираю продажи, остатки и FBO-накладные.",
        reply_markup=menu_for_message(message),
    )
    try:
        filename = await _build_full_excel_report(client, shop_id)
        await message.answer_document(
            FSInputFile(filename),
            caption=(
                "✅ <b>Подробный Excel-отчёт готов</b>\n\n"
                "Внутри листы: Сводка, Продажи 30 дней, Остатки, Заканчивается, "
                "Потерянные, FBO накладные, Состав накладных."
            ),
            reply_markup=menu_for_message(message),
        )
    except Exception as e:
        await send_api_error(message, e)


# --- Уведомления о новых заказах ---
# Логика простая и безопасная:
# 1) при первом запуске бот запоминает текущие CREATED-заказы и не спамит ими;
# 2) дальше каждые ORDER_CHECK_INTERVAL_SECONDS секунд проверяет новые CREATED-заказы;
# 3) если появился новый заказ, пишет продавцу в Telegram.
_seen_order_keys_by_user: dict[int, set[str]] = {}
_orders_watch_initialized: set[int] = set()


def order_key(order: Any) -> str:
    """Делаем стабильный ключ заказа из ID. Если ID в ответе API спрятан, берём hash JSON."""
    if isinstance(order, dict):
        for key in (
            "id",
            "orderId",
            "order_id",
            "shipmentId",
            "shipment_id",
            "postingNumber",
            "posting_number",
            "number",
            "barcode",
        ):
            value = order.get(key)
            if value not in (None, ""):
                return f"{key}:{value}"

        # Иногда ID лежит внутри вложенного объекта.
        for value in order.values():
            if isinstance(value, dict):
                nested = order_key(value)
                if nested:
                    return nested

    raw = json.dumps(order, ensure_ascii=False, sort_keys=True, default=str)
    return "hash:" + hashlib.sha1(raw.encode("utf-8")).hexdigest()


def connected_users_for_order_watch() -> list[dict[str, Any]]:
    """Берём всех пользователей, у кого подключён Uzum-токен и выбран магазин."""
    with db.connect() as conn:
        rows = conn.execute(
            """
            SELECT telegram_id, default_shop_id, uzum_token_encrypted
            FROM users
            WHERE uzum_token_encrypted IS NOT NULL
              AND default_shop_id IS NOT NULL
            """
        ).fetchall()
    return [dict(row) for row in rows if has_active_subscription(int(row["telegram_id"]))]


async def check_new_orders_once() -> None:
    users = connected_users_for_order_watch()
    for row in users:
        telegram_id = int(row["telegram_id"])
        shop_id = int(row["default_shop_id"])
        encrypted_token = row["uzum_token_encrypted"]

        try:
            token = cipher.decrypt(encrypted_token)
            client = UzumClient(token, UZUM_API_BASE_URL)
            data = await client.get_fbs_orders(shop_id, status="CREATED", page=0, size=20)
            items = extract_items(data)
        except Exception:
            logging.exception("Order watcher: failed to check orders for %s", telegram_id)
            continue

        keys_now = [order_key(item) for item in items]
        known = _seen_order_keys_by_user.setdefault(telegram_id, set())

        # Первый проход: просто запоминаем текущие заказы, чтобы не прислать старые как новые.
        if telegram_id not in _orders_watch_initialized:
            known.update(keys_now)
            _orders_watch_initialized.add(telegram_id)
            logging.info(
                "Order watcher initialized for user=%s shop=%s orders=%s",
                telegram_id,
                shop_id,
                len(keys_now),
            )
            continue

        new_items = [item for item, key in zip(items, keys_now) if key not in known]
        known.update(keys_now)

        # Чтобы память не росла бесконечно.
        if len(known) > 1000:
            _seen_order_keys_by_user[telegram_id] = set(keys_now)

        if not new_items:
            continue

        lines = [format_order_line(item) for item in new_items[:5]]
        more = "" if len(new_items) <= 5 else f"\n\nЕщё новых заказов: {len(new_items) - 5}"
        text = (
            f"🔔 <b>Новый заказ CREATED</b>\n"
            f"Магазин: <code>{shop_id}</code>\n"
            f"Новых заказов: <b>{len(new_items)}</b>\n\n"
            + "\n\n".join(lines)
            + more
            + "\n\nОткрыть список: <code>/orders</code>"
        )

        try:
            await bot.send_message(telegram_id, text, reply_markup=main_menu_for_user(telegram_id))
        except Exception:
            logging.exception("Order watcher: failed to send notification to %s", telegram_id)


async def order_watch_loop() -> None:
    await asyncio.sleep(10)
    logging.info(
        "Order watcher started. Interval: %s seconds. Enabled: %s",
        ORDER_CHECK_INTERVAL_SECONDS,
        NEW_ORDER_NOTIFICATIONS,
    )
    while True:
        try:
            await check_new_orders_once()
        except Exception:
            logging.exception("Order watcher loop error")
        await asyncio.sleep(max(60, ORDER_CHECK_INTERVAL_SECONDS))


@dp.message(Command("notify_status"))
async def notify_status(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    req = await require_connection(message)
    if req is None:
        return
    _, _, shop_id = req
    initialized = telegram_id in _orders_watch_initialized
    await message.answer(
        "🔔 <b>Уведомления о новых заказах</b>\n\n"
        f"Статус: {'✅ включены' if NEW_ORDER_NOTIFICATIONS else '❌ выключены'}\n"
        f"Магазин: <code>{shop_id}</code>\n"
        f"Проверка каждые: <b>{max(60, ORDER_CHECK_INTERVAL_SECONDS)}</b> сек.\n"
        f"Состояние: {'заказы уже запомнены' if initialized else 'инициализация при следующей проверке'}\n\n"
        "Бот уведомит, когда появится новый заказ со статусом <code>CREATED</code>.",
        reply_markup=menu_for_message(message),
    )


# --- Уведомления о низких остатках ---
# Логика:
# 1) при первом запуске бот запоминает текущие SKU с остатком ниже порога и не спамит ими;
# 2) дальше проверяет остатки каждые LOW_STOCK_CHECK_INTERVAL_SECONDS секунд;
# 3) если SKU впервые стал ниже/равен порогу, бот присылает уведомление.
_seen_low_stock_keys_by_user: dict[int, set[str]] = {}
_low_stock_watch_initialized: set[int] = set()


def stock_row_key(row: dict[str, Any]) -> str:
    for key in ("sku_id", "barcode", "seller_item_code", "product_id"):
        value = row.get(key)
        if value not in (None, ""):
            return f"{key}:{value}"
    raw = json.dumps(row, ensure_ascii=False, sort_keys=True, default=str)
    return "hash:" + hashlib.sha1(raw.encode("utf-8")).hexdigest()


async def check_low_stock_once() -> None:
    users = connected_users_for_order_watch()
    threshold = max(0, LOW_STOCK_THRESHOLD)

    for row in users:
        telegram_id = int(row["telegram_id"])
        shop_id = int(row["default_shop_id"])
        encrypted_token = row["uzum_token_encrypted"]

        try:
            token = cipher.decrypt(encrypted_token)
            client = UzumClient(token, UZUM_API_BASE_URL)
            rows = await load_sku_rows(client, shop_id, max_pages=50)
        except Exception:
            logging.exception("Low stock watcher: failed to check stocks for %s", telegram_id)
            continue

        low_rows = [
            r
            for r in rows
            if r.get("total") is not None and int(r.get("total") or 0) <= threshold
        ]
        low_keys_now = [stock_row_key(r) for r in low_rows]
        known = _seen_low_stock_keys_by_user.setdefault(telegram_id, set())

        # Первый проход: запоминаем текущие низкие остатки, чтобы не прислать старые как новые.
        if telegram_id not in _low_stock_watch_initialized:
            known.update(low_keys_now)
            _low_stock_watch_initialized.add(telegram_id)
            logging.info(
                "Low stock watcher initialized for user=%s shop=%s low_skus=%s threshold=%s",
                telegram_id,
                shop_id,
                len(low_keys_now),
                threshold,
            )
            continue

        new_low_rows = [r for r, key in zip(low_rows, low_keys_now) if key not in known]
        known.update(low_keys_now)

        # Если товар восстановился выше порога, удаляем его из known.
        # Тогда при повторном падении ниже порога бот снова уведомит.
        _seen_low_stock_keys_by_user[telegram_id] = set(low_keys_now)

        if not new_low_rows:
            continue

        lines = [format_sku_stock_line(item, mode="all") for item in new_low_rows[:10]]
        more = "" if len(new_low_rows) <= 10 else f"\n\nЕщё SKU с низким остатком: {len(new_low_rows) - 10}"
        text = (
            f"📉 <b>Товар заканчивается</b>\n"
            f"Магазин: <code>{shop_id}</code>\n"
            f"Порог: ≤ <b>{threshold}</b> шт.\n"
            f"Новых позиций с низким остатком: <b>{len(new_low_rows)}</b>\n\n"
            + "\n\n".join(lines)
            + more
            + f"\n\nПоказать все низкие остатки: <code>/lowstock {threshold}</code>"
        )

        try:
            await bot.send_message(telegram_id, text, reply_markup=main_menu_for_user(telegram_id))
        except Exception:
            logging.exception("Low stock watcher: failed to send notification to %s", telegram_id)


async def low_stock_watch_loop() -> None:
    await asyncio.sleep(20)
    logging.info(
        "Low stock watcher started. Interval: %s seconds. Threshold: %s. Enabled: %s",
        LOW_STOCK_CHECK_INTERVAL_SECONDS,
        LOW_STOCK_THRESHOLD,
        LOW_STOCK_NOTIFICATIONS,
    )
    while True:
        try:
            await check_low_stock_once()
        except Exception:
            logging.exception("Low stock watcher loop error")
        await asyncio.sleep(max(300, LOW_STOCK_CHECK_INTERVAL_SECONDS))


@dp.message(Command("lowstock_notify_status"))
async def lowstock_notify_status(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    req = await require_connection(message)
    if req is None:
        return
    _, _, shop_id = req
    initialized = telegram_id in _low_stock_watch_initialized
    await message.answer(
        "📉 <b>Уведомления о низких остатках</b>\n\n"
        f"Статус: {'✅ включены' if LOW_STOCK_NOTIFICATIONS else '❌ выключены'}\n"
        f"Магазин: <code>{shop_id}</code>\n"
        f"Порог: ≤ <b>{LOW_STOCK_THRESHOLD}</b> шт.\n"
        f"Проверка каждые: <b>{max(300, LOW_STOCK_CHECK_INTERVAL_SECONDS)}</b> сек.\n"
        f"Состояние: {'остатки уже запомнены' if initialized else 'инициализация при следующей проверке'}\n\n"
        "Бот уведомит, когда товар впервые опустится до порога или ниже.",
        reply_markup=menu_for_message(message),
    )


# --- Уведомления о товарах с нулевым остатком ---
# Логика:
# 1) при первом запуске бот запоминает текущие SKU с остатком 0 и не спамит ими;
# 2) дальше проверяет остатки каждые OUT_OF_STOCK_CHECK_INTERVAL_SECONDS секунд;
# 3) если SKU впервые стал равен 0, бот присылает отдельное срочное уведомление.
_seen_out_of_stock_keys_by_user: dict[int, set[str]] = {}
_out_of_stock_watch_initialized: set[int] = set()


async def check_out_of_stock_once() -> None:
    users = connected_users_for_order_watch()

    for row in users:
        telegram_id = int(row["telegram_id"])
        shop_id = int(row["default_shop_id"])
        encrypted_token = row["uzum_token_encrypted"]

        try:
            token = cipher.decrypt(encrypted_token)
            client = UzumClient(token, UZUM_API_BASE_URL)
            rows = await load_sku_rows(client, shop_id, max_pages=50)
        except Exception:
            logging.exception("Out of stock watcher: failed to check stocks for %s", telegram_id)
            continue

        zero_rows = [
            r
            for r in rows
            if r.get("total") is not None and int(r.get("total") or 0) == 0
        ]
        zero_keys_now = [stock_row_key(r) for r in zero_rows]
        known = _seen_out_of_stock_keys_by_user.setdefault(telegram_id, set())

        # Первый проход: запоминаем текущие нулевые остатки, чтобы не прислать старые как новые.
        if telegram_id not in _out_of_stock_watch_initialized:
            known.update(zero_keys_now)
            _out_of_stock_watch_initialized.add(telegram_id)
            logging.info(
                "Out of stock watcher initialized for user=%s shop=%s zero_skus=%s",
                telegram_id,
                shop_id,
                len(zero_keys_now),
            )
            continue

        new_zero_rows = [r for r, key in zip(zero_rows, zero_keys_now) if key not in known]

        # Если товар снова появился в наличии, удаляем его из known.
        # Тогда при повторном падении до 0 бот снова уведомит.
        _seen_out_of_stock_keys_by_user[telegram_id] = set(zero_keys_now)

        if not new_zero_rows:
            continue

        lines = [format_sku_stock_line(item, mode="all") for item in new_zero_rows[:10]]
        more = "" if len(new_zero_rows) <= 10 else f"\n\nЕщё SKU с нулевым остатком: {len(new_zero_rows) - 10}"
        text = (
            f"❌ <b>Товар закончился</b>\n"
            f"Магазин: <code>{shop_id}</code>\n"
            f"Новых позиций с остатком 0: <b>{len(new_zero_rows)}</b>\n\n"
            + "\n\n".join(lines)
            + more
            + "\n\nПоказать товары с низким остатком: <code>/lowstock 0</code>"
        )

        try:
            await bot.send_message(telegram_id, text, reply_markup=main_menu_for_user(telegram_id))
        except Exception:
            logging.exception("Out of stock watcher: failed to send notification to %s", telegram_id)


async def out_of_stock_watch_loop() -> None:
    await asyncio.sleep(30)
    logging.info(
        "Out of stock watcher started. Interval: %s seconds. Enabled: %s",
        OUT_OF_STOCK_CHECK_INTERVAL_SECONDS,
        OUT_OF_STOCK_NOTIFICATIONS,
    )
    while True:
        try:
            await check_out_of_stock_once()
        except Exception:
            logging.exception("Out of stock watcher loop error")
        await asyncio.sleep(max(300, OUT_OF_STOCK_CHECK_INTERVAL_SECONDS))


@dp.message(Command("outofstock_notify_status"))
async def outofstock_notify_status(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    req = await require_connection(message)
    if req is None:
        return
    _, _, shop_id = req
    initialized = telegram_id in _out_of_stock_watch_initialized
    await message.answer(
        "❌ <b>Уведомления о нулевых остатках</b>\n\n"
        f"Статус: {'✅ включены' if OUT_OF_STOCK_NOTIFICATIONS else '❌ выключены'}\n"
        f"Магазин: <code>{shop_id}</code>\n"
        f"Проверка каждые: <b>{max(300, OUT_OF_STOCK_CHECK_INTERVAL_SECONDS)}</b> сек.\n"
        f"Состояние: {'нулевые остатки уже запомнены' if initialized else 'инициализация при следующей проверке'}\n\n"
        "Бот уведомит, когда товар впервые опустится до остатка <b>0</b>.",
        reply_markup=menu_for_message(message),
    )


# --- Уведомления о новых продажах через Finance API ---
# Важно: старое уведомление "Новые заказы" смотрит только FBS/DBS заказы CREATED.
# Если продажа была FBO или уже не имеет статус CREATED, она может не попасть в этот список.
# Этот watcher смотрит финансовые строки продаж за сегодня и присылает уведомление о новых строках.
_seen_sale_keys_by_user: dict[int, set[str]] = {}
_sales_watch_initialized: set[int] = set()


def sale_key(item: dict[str, Any]) -> str:
    for key in (
        "id",
        "orderItemId",
        "orderItem_id",
        "orderId",
        "order_id",
        "skuId",
        "sku_id",
        "postingNumber",
        "number",
        "barcode",
    ):
        value = item.get(key)
        if value not in (None, ""):
            # Добавляем дату/сумму, чтобы разные продажи одного SKU не слиплись.
            date_part = ""
            for dk in ("date", "orderDate", "createdAt", "operationDate", "saleDate"):
                dv = item.get(dk)
                if dv not in (None, ""):
                    date_part = str(dv)
                    break
            amount = _finance_revenue(item)
            qty = _finance_qty(item)
            return f"{key}:{value}|{date_part}|{qty}|{amount}"
    raw = json.dumps(item, ensure_ascii=False, sort_keys=True, default=str)
    return "hash:" + hashlib.sha1(raw.encode("utf-8")).hexdigest()


def _deep_pick_value(obj: Any, names: tuple[str, ...]) -> Any:
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k in names and v not in (None, ""):
                return v
        for v in obj.values():
            found = _deep_pick_value(v, names)
            if found not in (None, ""):
                return found
    elif isinstance(obj, list):
        for v in obj:
            found = _deep_pick_value(v, names)
            if found not in (None, ""):
                return found
    return None


def _finance_sku_title(item: dict[str, Any]) -> str:
    value = _deep_pick_value(item, ("skuTitle", "skuName", "skuFullTitle", "offerName", "barcode", "skuId"))
    if isinstance(value, dict):
        value = pick(value, "title", "name", "value", "id")
    return str(value or "-")


def _finance_order_id(item: dict[str, Any]) -> str:
    return str(_deep_pick_value(item, ("orderId", "order_id", "orderNumber", "postingNumber")) or "-")


def _finance_sale_id(item: dict[str, Any]) -> str:
    return str(_deep_pick_value(item, ("id", "saleId", "operationId", "orderItemId", "orderItem_id")) or "-")


def _finance_date_value(item: dict[str, Any]) -> Any:
    return _deep_pick_value(item, ("date", "saleDate", "operationDate", "createdAt", "orderDate", "createdDate"))


def _format_finance_date(value: Any) -> str:
    if value in (None, ""):
        return "-"
    if isinstance(value, (int, float)):
        ts = float(value)
        if ts > 10_000_000_000:
            ts = ts / 1000
        try:
            return datetime.fromtimestamp(ts, UZT).strftime("%d.%m.%Y %H:%M")
        except Exception:
            return str(value)
    text_value = str(value).strip()
    try:
        iso = text_value.replace("Z", "+00:00")
        dt = datetime.fromisoformat(iso)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=UZT)
        return dt.astimezone(UZT).strftime("%d.%m.%Y %H:%M")
    except Exception:
        return text_value[:16]


def format_sale_line(item: dict[str, Any]) -> str:
    title = escape(_finance_title(item))
    qty = _finance_qty(item)
    amount = _finance_revenue(item)
    status = escape(_finance_status(item))
    return (
        f"• <b>{title}</b>\n"
        f"  Штук: <b>{qty:g}</b> | Сумма: <b>{_format_money(amount)}</b> | Статус: <code>{status}</code>"
    )


def build_new_sale_message(item: dict[str, Any], shop_id: int | None = None, lang: str = "ru") -> str:
    lang = normalize_lang(lang)
    title = escape(_finance_title(item))
    sku = escape(_finance_sku_title(item))
    qty = _finance_qty(item)

    # Для уведомления показываем цену за штуку, как в Noorza Bot.
    unit_price = _deep_pick_number(item, ("sellPrice", "soldPrice", "price", "skuPrice", "productPrice"))
    if unit_price is None:
        unit_price = _finance_gross_revenue(item) / max(1.0, qty)

    commission = _finance_commission(item)
    logistics = _finance_logistics(item)
    payout_direct = _finance_payout_direct(item)
    payout = payout_direct if payout_direct is not None else max(0.0, _finance_gross_revenue(item) - commission - logistics)

    if lang == "uz":
        shop_line = f"🏪 Do‘kon: <code>{shop_id}</code>\n" if shop_id is not None else ""
        return (
            "🛒 <b>Yangi savdo</b>\n\n"
            + shop_line +
            f"📦 Tovar: <b>{title}</b>\n"
            f"🔖 SKU: <code>{sku}</code>\n"
            f"🔢 Soni: <b>{qty:g} dona</b>\n\n"
            f"💵 Narx: <b>{_format_money(float(unit_price or 0))}</b>\n"
            f"🏷 Komissiya: <b>{_format_money(float(commission))}</b>\n"
            f"🚚 Logistika: <b>{_format_money(float(logistics))}</b>\n"
            f"✅ To‘lovga: <b>{_format_money(float(payout))}</b>\n\n"
            f"🆔 Buyurtma: <code>{escape(_finance_order_id(item))}</code>\n"
            f"📌 Status: <code>{escape(_finance_status(item))}</code>\n"
            f"🕒 Sana: {escape(_format_finance_date(_finance_date_value(item)))}"
        )

    shop_line = f"🏪 Магазин: <code>{shop_id}</code>\n" if shop_id is not None else ""
    return (
        "🛒 <b>Новая продажа</b>\n\n"
        + shop_line +
        f"📦 Товар: <b>{title}</b>\n"
        f"🔖 SKU: <code>{sku}</code>\n"
        f"🔢 Кол-во: <b>{qty:g} шт.</b>\n\n"
        f"💵 Цена: <b>{_format_money(float(unit_price or 0))}</b>\n"
        f"🏷 Комиссия: <b>{_format_money(float(commission))}</b>\n"
        f"🚚 Логистика: <b>{_format_money(float(logistics))}</b>\n"
        f"✅ К выплате: <b>{_format_money(float(payout))}</b>\n\n"
        f"🆔 Заказ: <code>{escape(_finance_order_id(item))}</code>\n"
        f"📌 Статус: <code>{escape(_finance_status(item))}</code>\n"
        f"🕒 Дата: {escape(_format_finance_date(_finance_date_value(item)))}"
    )


async def check_new_sales_once() -> None:
    users = connected_users_for_order_watch()
    date_from, date_to = _today_range_ms()

    for row in users:
        telegram_id = int(row["telegram_id"])
        shop_id = int(row["default_shop_id"])
        encrypted_token = row["uzum_token_encrypted"]

        try:
            token = cipher.decrypt(encrypted_token)
            client = UzumClient(token, UZUM_API_BASE_URL)
            rows, first_response = await _load_finance_orders(
                client,
                shop_id,
                date_from_ms=date_from,
                date_to_ms=date_to,
                max_pages=5,
                page_size=100,
            )
        except Exception:
            logging.exception("Sales watcher: failed to check sales for %s", telegram_id)
            continue

        keys_now = [sale_key(item) for item in rows]
        known = _seen_sale_keys_by_user.setdefault(telegram_id, set())

        # Первый проход: запоминаем текущие продажи за сегодня, чтобы не прислать старые.
        if telegram_id not in _sales_watch_initialized:
            known.update(keys_now)
            _sales_watch_initialized.add(telegram_id)
            logging.info(
                "Sales watcher initialized for user=%s shop=%s sales_rows=%s",
                telegram_id,
                shop_id,
                len(keys_now),
            )
            continue

        new_rows = [item for item, key in zip(rows, keys_now) if key not in known]
        known.update(keys_now)

        if len(known) > 3000:
            _seen_sale_keys_by_user[telegram_id] = set(keys_now)

        if not new_rows:
            continue

        # Отправляем каждую новую продажу отдельным сообщением в стиле Noorza Bot.
        for item in new_rows[:10]:
            try:
                await bot.send_message(
                    telegram_id,
                    build_new_sale_message(item, shop_id=shop_id, lang=get_user_language(telegram_id)),
                    reply_markup=main_menu_for_user(telegram_id),
                )
                await asyncio.sleep(0.15)
            except Exception:
                logging.exception("Sales watcher: failed to send sale notification to %s", telegram_id)

        if len(new_rows) > 10:
            try:
                await bot.send_message(
                    telegram_id,
                    f"➕ Ещё новых строк продаж: <b>{len(new_rows) - 10}</b>\nПодробно: <code>/balance</code>",
                    reply_markup=main_menu_for_user(telegram_id),
                )
            except Exception:
                logging.exception("Sales watcher: failed to send summary notification to %s", telegram_id)


async def sales_watch_loop() -> None:
    await asyncio.sleep(40)
    logging.info(
        "Sales watcher started. Interval: %s seconds. Enabled: %s",
        SALE_CHECK_INTERVAL_SECONDS,
        SALE_NOTIFICATIONS,
    )
    while True:
        try:
            await check_new_sales_once()
        except Exception:
            logging.exception("Sales watcher loop error")
        await asyncio.sleep(max(60, SALE_CHECK_INTERVAL_SECONDS))


@dp.message(Command("sales_notify_status"))
async def sales_notify_status(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    req = await require_connection(message)
    if req is None:
        return
    _, _, shop_id = req
    initialized = telegram_id in _sales_watch_initialized
    await message.answer(
        "💸 <b>Уведомления о новых продажах</b>\n\n"
        f"Статус: {'✅ включены' if SALE_NOTIFICATIONS else '❌ выключены'}\n"
        f"Магазин: <code>{shop_id}</code>\n"
        f"Проверка каждые: <b>{max(60, SALE_CHECK_INTERVAL_SECONDS)}</b> сек.\n"
        f"Состояние: {'продажи уже запомнены' if initialized else 'инициализация при следующей проверке'}\n\n"
        "Бот смотрит Finance API за сегодня. Если Finance API отдаёт продажу с задержкой, уведомление тоже придёт с задержкой.",
        reply_markup=menu_for_message(message),
    )



# --- Уведомления об изменении остатков: FBO + FBS/DBS ---
# Это нужно для FBO-продаж: FBO-заказ может не появиться в FBS/DBS CREATED,
# но остаток на складе Uzum уменьшается. Бот сравнивает общий остаток, FBO и FBS.
_stock_snapshot_by_user: dict[int, dict[str, dict[str, Any]]] = {}
_stock_change_watch_initialized: set[int] = set()


def _stock_qty(value: Any) -> int:
    num = _num_from_value(value)
    if num is None:
        return 0
    return int(num)


def _stock_snapshot_item(row: dict[str, Any]) -> dict[str, Any]:
    return {
        "total": _stock_qty(row.get("total")),
        "fbo": _stock_qty(row.get("fbo")),
        "fbs": _stock_qty(row.get("fbs")),
        "title": str(
            row.get("title")
            or row.get("productTitle")
            or row.get("skuTitle")
            or row.get("name")
            or row.get("product_name")
            or "SKU"
        ),
        "price": row.get("price"),
        "row": row,
    }


def _stock_change_snapshot(rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    snapshot: dict[str, dict[str, Any]] = {}
    for row in rows:
        snapshot[stock_row_key(row)] = _stock_snapshot_item(row)
    return snapshot


def _format_stock_change_line(key: str, before: dict[str, Any], after: dict[str, Any]) -> str:
    title = escape(str(after.get("title") or before.get("title") or key))
    old_total = int(before.get("total") or 0)
    new_total = int(after.get("total") or 0)
    old_fbo = int(before.get("fbo") or 0)
    new_fbo = int(after.get("fbo") or 0)
    old_fbs = int(before.get("fbs") or 0)
    new_fbs = int(after.get("fbs") or 0)

    parts = []
    if old_total != new_total:
        parts.append(f"Итого: <b>{old_total}</b> → <b>{new_total}</b>")
    if old_fbo != new_fbo:
        parts.append(f"FBO: <b>{old_fbo}</b> → <b>{new_fbo}</b>")
    if old_fbs != new_fbs:
        parts.append(f"FBS/DBS: <b>{old_fbs}</b> → <b>{new_fbs}</b>")
    if not parts:
        parts.append("остаток изменился")

    delta = new_total - old_total
    delta_text = f" | Разница: <b>{delta}</b> шт" if delta else ""
    return f"• <b>{title}</b>\n  " + " | ".join(parts) + delta_text


async def check_stock_change_once() -> None:
    users = connected_users_for_order_watch()

    for row in users:
        telegram_id = int(row["telegram_id"])
        shop_id = int(row["default_shop_id"])
        encrypted_token = row["uzum_token_encrypted"]

        try:
            token = cipher.decrypt(encrypted_token)
            client = UzumClient(token, UZUM_API_BASE_URL)
            rows = await load_sku_rows(client, shop_id, max_pages=20)
            snapshot_now = _stock_change_snapshot(rows)
        except Exception:
            logging.exception("Stock change watcher: failed to check stocks for %s", telegram_id)
            continue

        previous = _stock_snapshot_by_user.setdefault(telegram_id, {})

        # Первый проход: только запоминаем, чтобы не прислать старые изменения.
        if telegram_id not in _stock_change_watch_initialized:
            _stock_snapshot_by_user[telegram_id] = snapshot_now
            _stock_change_watch_initialized.add(telegram_id)
            logging.info(
                "Stock change watcher initialized for user=%s shop=%s skus=%s",
                telegram_id,
                shop_id,
                len(snapshot_now),
            )
            continue

        decreased: list[tuple[str, dict[str, Any], dict[str, Any]]] = []
        for key, after in snapshot_now.items():
            before = previous.get(key)
            if not before:
                continue
            before_total = int(before.get("total") or 0)
            after_total = int(after.get("total") or 0)
            before_fbo = int(before.get("fbo") or 0)
            after_fbo = int(after.get("fbo") or 0)
            before_fbs = int(before.get("fbs") or 0)
            after_fbs = int(after.get("fbs") or 0)

            if after_total < before_total or after_fbo < before_fbo or after_fbs < before_fbs:
                decreased.append((key, before, after))

        _stock_snapshot_by_user[telegram_id] = snapshot_now

        if not decreased:
            continue

        lines = [_format_stock_change_line(key, before, after) for key, before, after in decreased[:10]]
        more = "" if len(decreased) <= 10 else f"\n\nЕщё изменений: {len(decreased) - 10}"
        text = (
            "📦 <b>Изменение остатков</b>\n"
            f"Магазин: <code>{shop_id}</code>\n"
            "Уменьшился остаток по SKU. Это может быть продажа, резерв, списание или изменение склада.\n\n"
            + "\n\n".join(lines)
            + more
            + "\n\nПроверить остатки: <code>/stock</code>"
        )

        try:
            await bot.send_message(telegram_id, text, reply_markup=main_menu_for_user(telegram_id))
        except Exception:
            logging.exception("Stock change watcher: failed to send notification to %s", telegram_id)


async def stock_change_watch_loop() -> None:
    await asyncio.sleep(70)
    logging.info(
        "Stock change watcher started. Interval: %s seconds. Enabled: %s",
        STOCK_CHANGE_CHECK_INTERVAL_SECONDS,
        STOCK_CHANGE_NOTIFICATIONS,
    )
    while True:
        try:
            await check_stock_change_once()
        except Exception:
            logging.exception("Stock change watcher loop error")
        await asyncio.sleep(max(60, STOCK_CHANGE_CHECK_INTERVAL_SECONDS))


@dp.message(Command("stock_change_notify_status"))
async def stock_change_notify_status(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    req = await require_connection(message)
    if req is None:
        return
    _, _, shop_id = req
    initialized = telegram_id in _stock_change_watch_initialized
    await message.answer(
        "📦 <b>Уведомления об изменении остатков</b>\n\n"
        f"Статус: {'✅ включены' if STOCK_CHANGE_NOTIFICATIONS else '❌ выключены'}\n"
        f"Магазин: <code>{shop_id}</code>\n"
        f"Проверка каждые: <b>{max(60, STOCK_CHANGE_CHECK_INTERVAL_SECONDS)}</b> сек.\n"
        f"Состояние: {'остатки уже запомнены' if initialized else 'инициализация при следующей проверке'}\n\n"
        "Бот сравнивает <b>FBO</b>, <b>FBS/DBS</b> и <b>общий остаток</b>. "
        "Так можно поймать FBO-продажи, даже если Finance API показывает нули.",
        reply_markup=menu_for_message(message),
    )



# --- Умный раздел: что требует внимания ---
def _attention_recommendation_ru(low_count: int, zero_count: int, dead_count: int, missing_cost: int, low_margin: int, cancels_today: int) -> str:
    if zero_count > 0:
        return "Начните с товаров, которые закончились: они не смогут продаваться, пока не пополнятся остатки."
    if low_count > 0:
        return "Сначала проверьте товары, которые скоро закончатся, особенно если они хорошо продаются."
    if missing_cost > 0:
        return "Загрузите себестоимость через Excel, чтобы бот точнее считал прибыль и маржу."
    if low_margin > 0:
        return "Проверьте товары с низкой маржой: возможно, цена или себестоимость указаны невыгодно."
    if dead_count > 0:
        return "Посмотрите товары без продаж: возможно, стоит изменить цену, фото или вывести товар из оборота."
    if cancels_today > 0:
        return "Проверьте сегодняшние отмены и товары, по которым они произошли."
    return "Критичных проблем не видно. Можно посмотреть продажи и прибыль за 30 дней."


def _attention_recommendation_uz(low_count: int, zero_count: int, dead_count: int, missing_cost: int, low_margin: int, cancels_today: int) -> str:
    if zero_count > 0:
        return "Avval qoldig‘i tugagan tovarlarni tekshiring: qoldiq bo‘lmasa, savdo ham bo‘lmaydi."
    if low_count > 0:
        return "Avval tez tugayotgan tovarlarni tekshiring, ayniqsa ular yaxshi sotilayotgan bo‘lsa."
    if missing_cost > 0:
        return "Foyda va marjani aniq hisoblash uchun tannarxni Excel orqali yuklang."
    if low_margin > 0:
        return "Past marjali tovarlarni tekshiring: narx yoki tannarx foydasiz bo‘lishi mumkin."
    if dead_count > 0:
        return "30 kun sotilmagan tovarlarni ko‘rib chiqing: narx, rasm yoki aylanmani tekshirish kerak."
    if cancels_today > 0:
        return "Bugungi bekor qilishlarni tekshiring."
    return "Jiddiy muammo ko‘rinmayapti. 30 kunlik savdo va foydani ko‘rishingiz mumkin."


async def _build_attention_summary(message: Message) -> str | None:
    req = await require_connection(message)
    if req is None:
        return None
    telegram_id, client, shop_id = req
    lang = get_user_language(telegram_id)

    # 1) Остатки
    stock_rows = await load_sku_rows(client, shop_id, max_pages=50)
    low_count = 0
    zero_count = 0
    for row in stock_rows:
        total = _stock_row_total(row)
        if total <= 0:
            zero_count += 1
        elif total <= LOW_STOCK_THRESHOLD:
            low_count += 1

    # 2) Потерянные товары по quantityMissing
    try:
        products = await load_products(client, shop_id, max_pages=50, page_size=100)
        missing_count = sum(1 for p in products if isinstance(p, dict) and _product_missing_qty(p) > 0)
    except Exception:
        missing_count = 0

    # 3) Продажи/отмены сегодня и товары без продаж за DEAD_STOCK_DAYS
    date_today_from, date_today_to = _today_range_ms()
    today_rows, _ = await _load_finance_orders(client, shop_id, date_from_ms=date_today_from, date_to_ms=date_today_to, max_pages=3, page_size=100)
    cancels_today = sum(1 for item in today_rows if _is_cancelled_status(_finance_status(item)))

    date_from, date_to = _days_range_ms(DEAD_STOCK_DAYS)
    sales_rows, _ = await _load_finance_orders(client, shop_id, date_from_ms=date_from, date_to_ms=date_to, max_pages=5, page_size=100)
    sold_keys: set[str] = set()
    for item in sales_rows:
        if not _is_cancelled_status(_finance_status(item)):
            sold_keys.update(_sale_match_keys(item))
    dead_count = 0
    for row in stock_rows:
        if _stock_row_total(row) <= 0:
            continue
        keys = _stock_match_keys(row)
        if keys and not keys.intersection(sold_keys):
            dead_count += 1

    # 4) Юнит-экономика: нет себестоимости и низкая маржа
    try:
        unit_rows, _stats, _saved_costs = await _unit_economy_for_shop(client, telegram_id, shop_id, days=30)
    except Exception:
        unit_rows = []
    missing_cost = sum(1 for r in unit_rows if r.get("cost_per_unit") is None)
    low_margin = sum(
        1
        for r in unit_rows
        if r.get("cost_per_unit") is not None
        and r.get("profit") is not None
        and float(r.get("margin") or 0) < LOW_MARGIN_THRESHOLD_PERCENT
    )

    if lang == "uz":
        rec = _attention_recommendation_uz(low_count, zero_count, dead_count, missing_cost, low_margin, cancels_today)
        return (
            "🧠 <b>Nimani tekshirish kerak</b>\n\n"
            f"🏪 Do‘kon: <code>{shop_id}</code>\n\n"
            f"⚠️ Tez tugayotgan tovarlar: <b>{low_count}</b>\n"
            f"❌ Qoldig‘i tugagan: <b>{zero_count}</b>\n"
            f"🐢 {DEAD_STOCK_DAYS} kun sotilmagan: <b>{dead_count}</b>\n"
            f"🧾 Tannarxi kiritilmagan: <b>{missing_cost}</b>\n"
            f"📉 Past marja: <b>{low_margin}</b>\n"
            f"❌ Bugungi bekor qilishlar: <b>{cancels_today}</b>\n"
            f"🧭 Yo‘qolganlar: <b>{missing_count}</b>\n\n"
            f"💡 <b>Tavsiya:</b> {escape(rec)}\n\n"
            "Pastdagi tugmalar orqali kerakli bo‘limni oching 👇"
        )

    rec = _attention_recommendation_ru(low_count, zero_count, dead_count, missing_cost, low_margin, cancels_today)
    return (
        "🧠 <b>Что требует внимания</b>\n\n"
        f"🏪 Магазин: <code>{shop_id}</code>\n\n"
        f"⚠️ Скоро закончится: <b>{low_count}</b>\n"
        f"❌ Закончились: <b>{zero_count}</b>\n"
        f"🐢 Без продаж {DEAD_STOCK_DAYS} дней: <b>{dead_count}</b>\n"
        f"🧾 Без себестоимости: <b>{missing_cost}</b>\n"
        f"📉 Низкая маржа: <b>{low_margin}</b>\n"
        f"❌ Отмены сегодня: <b>{cancels_today}</b>\n"
        f"🧭 Потерянные: <b>{missing_count}</b>\n\n"
        f"💡 <b>Рекомендация:</b> {escape(rec)}\n\n"
        "Ниже можете сразу открыть нужный раздел 👇"
    )


@dp.message(Command("attention", "check_attention", "what_check"))
async def attention_report(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    wait_text = "⌛ Do‘konni tekshiryapman..." if lang == "uz" else "⌛ Проверяю магазин..."
    await message.answer(wait_text, reply_markup=attention_menu_for_message(message))
    try:
        text = await _build_attention_summary(message)
        if text:
            await message.answer(text, reply_markup=attention_menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)

# --- Главное меню в стиле Noorza Bot ---
@dp.message(F.text == "👑 Admin")
@dp.message(F.text == "👑 Админ")
async def button_admin_panel(message: Message) -> None:
    await admin_panel(message)


@dp.message(F.text == "👥 Foydalanuvchilar")
@dp.message(F.text == "👥 Пользователи")
async def button_admin_users(message: Message) -> None:
    await admin_users(message)


@dp.message(F.text == "💳 To‘lovlar")
@dp.message(F.text == "💳 Оплаты")
async def button_admin_payments(message: Message) -> None:
    await admin_payments(message)


@dp.message(F.text == "⏳ Tugayotganlar")
@dp.message(F.text == "⏳ Скоро заканчиваются")
async def button_admin_expiring(message: Message) -> None:
    await admin_expiring(message)


@dp.message(F.text == "⛔ Bloklanganlar")
@dp.message(F.text == "⛔ Заблокированные")
async def button_admin_blocked(message: Message) -> None:
    await admin_blocked_users(message)


@dp.message(F.text == "💰 Savdo")
@dp.message(F.text == "💰 Продажи")
async def button_sales_section_simple(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    text = "💰 <b>Savdo bo‘limi</b>\nKerakli davr yoki hisobotni tanlang 👇" if lang == "uz" else "💰 <b>Продажи</b>\nВыберите, что посмотреть 👇"
    await message.answer(text, reply_markup=sales_menu_for_message(message))


@dp.message(F.text == "📦 Ombor")
@dp.message(F.text == "📦 Склад")
async def button_stock_section_simple(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    text = "📦 <b>Ombor</b>\nQoldiq, prognoz yoki FBO yuk xatlarini tanlang 👇" if lang == "uz" else "📦 <b>Склад</b>\nОстатки, прогноз и FBO-накладные 👇"
    await message.answer(text, reply_markup=stock_menu_for_message(message))


@dp.message(F.text == "🔔 Xabarnomalar")
@dp.message(F.text == "🔔 Уведомления")
async def button_notifications_section_simple(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    text = "🔔 <b>Xabarnomalar</b>\nKerakli holatni tanlang 👇" if lang == "uz" else "🔔 <b>Уведомления</b>\nПроверьте, что включено 👇"
    await message.answer(text, reply_markup=notify_menu_for_message(message))


@dp.message(F.text == "📊 Hisobotlar")
@dp.message(F.text == "📊 Отчёты")
async def button_reports_section_simple(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    text = "📊 <b>Hisobotlar</b>\nExcel, foyda va tayyor hisobotlar 👇" if lang == "uz" else "📊 <b>Отчёты</b>\nExcel, прибыль и готовые отчёты 👇"
    await message.answer(text, reply_markup=report_menu_for_message(message))


@dp.message(F.text == "🧠 Tekshirish")
@dp.message(F.text == "🧠 Что проверить")
async def button_attention_section(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    if lang == "uz":
        text = (
            "🧠 <b>Nimani tekshirish kerak</b>\n\n"
            "Bot do‘koningizdagi muhim joylarni tekshiradi:\n"
            "⚠️ tez tugayotgan qoldiqlar, 🐢 sotilmayotgan tovarlar, "
            "🧾 tannarxi kiritilmagan SKU va 📉 past marja.\n\n"
            "Boshlash uchun <b>🔍 Hozir tekshirish</b> tugmasini bosing 👇"
        )
    else:
        text = (
            "🧠 <b>Что проверить</b>\n\n"
            "Бот быстро покажет, где есть проблемы:\n"
            "⚠️ заканчивающиеся остатки, 🐢 товары без продаж, "
            "🧾 SKU без себестоимости и 📉 низкую маржу.\n\n"
            "Чтобы начать, нажмите <b>🔍 Проверить сейчас</b> 👇"
        )
    await message.answer(text, reply_markup=attention_menu_for_message(message))


@dp.message(F.text == "🔍 Hozir tekshirish")
@dp.message(F.text == "🔍 Проверить сейчас")
async def button_attention_now(message: Message) -> None:
    await attention_report(message)


@dp.message(F.text == "⚠️ Qoldiqlar")
@dp.message(F.text == "⚠️ Остатки")
async def button_attention_stock(message: Message) -> None:
    await smart_lowstock(message)


@dp.message(F.text == "🐢 Sotuv yo‘q")
@dp.message(F.text == "🐢 Без продаж")
async def button_attention_dead_stock(message: Message) -> None:
    await dead_stock(message)


@dp.message(F.text == "🧾 Tannarx yo‘q")
@dp.message(F.text == "🧾 Нет себестоимости")
async def button_attention_missing_cost(message: Message) -> None:
    await unit_economy(message)


@dp.message(F.text == "📉 Past foyda")
@dp.message(F.text == "📉 Низкая прибыль")
async def button_attention_low_margin(message: Message) -> None:
    await profit_report(message)


@dp.message(F.text == "❌ Bekor qilishlar")
@dp.message(F.text == "❌ Отмены")
async def button_attention_cancel(message: Message) -> None:
    await sales_30(message)


@dp.message(F.text == "🔐 Xavfsizlik")
@dp.message(F.text == "🔐 Безопасность")
async def button_security_simple(message: Message) -> None:
    await security(message)


@dp.message(F.text == "💸 Yangi savdolar")
async def button_sales_notify_status_uz(message: Message) -> None:
    await sales_notify_status(message)


@dp.message(F.text == "📉 Kam qoldiq")
async def button_lowstock_notify_status_uz(message: Message) -> None:
    await lowstock_notify_status(message)


@dp.message(F.text == "❌ Qoldiq tugagan")
async def button_outofstock_notify_status_uz(message: Message) -> None:
    await outofstock_notify_status(message)


@dp.message(F.text == "⚙️ Holat")
async def button_status_uz(message: Message) -> None:
    await status(message)


@dp.message(F.text == "📦 Baza zaxirasi")
@dp.message(F.text == "📦 Бэкап базы")
async def button_admin_backup(message: Message) -> None:
    await admin_backup_db(message)


@dp.message(F.text == "📢 Xabar yuborish")
@dp.message(F.text == "📢 Рассылка")
async def button_admin_broadcast_help(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    await message.answer(
        "📢 <b>Рассылка</b>\n\n"
        "Чтобы отправить сообщение всем пользователям, напишите:\n"
        "<code>/broadcast ваш текст</code>\n\n"
        "Пример:\n"
        "<code>/broadcast Завтра в 09:00 будет обновление бота.</code>",
        reply_markup=admin_menu_for_message(message),
    )


@dp.message(F.text == "✅ Ulanishni tekshirish")
@dp.message(F.text == "✅ Проверить подключение")
async def button_check_connection(message: Message) -> None:
    await check_connection(message)


@dp.message(F.text == "⬅️ Asosiy menyu")
@dp.message(F.text == "⬅️ Главное меню")
@dp.message(F.text == "Menyu")
@dp.message(F.text == "Меню")
async def button_main_menu(message: Message) -> None:
    await message.answer(tr_user(upsert_from_message(message), "main_menu"), reply_markup=menu_for_message(message))


@dp.message(F.text == "💰 Balans")
@dp.message(F.text == "💰 Баланс")
async def button_balance(message: Message) -> None:
    await balance(message)


@dp.message(F.text == "📊 Bugun")
@dp.message(F.text == "📊 Сегодня")
async def button_today(message: Message) -> None:
    await today_sales(message)


@dp.message(F.text == "📆 Kecha")
@dp.message(F.text == "📆 Вчера")
async def button_yesterday(message: Message) -> None:
    await yesterday_sales(message)


@dp.message(F.text == "🗓 7 kun")
@dp.message(F.text == "🗓 7 дней")
async def button_week(message: Message) -> None:
    await week_sales(message)


@dp.message(F.text == "📅 30 kun")
@dp.message(F.text == "📅 30 дней")
async def button_30_days(message: Message) -> None:
    await sales_30(message)


@dp.message(F.text == "📦 Qoldiq")
@dp.message(F.text == "📦 Остатки")
async def button_stock_short(message: Message) -> None:
    await stock(message)


@dp.message(F.text == "⚠️ Заканчивается")
@dp.message(F.text == "⚠️ Заканчиваются")
async def button_lowstock_short(message: Message) -> None:
    await lowstock(message)


@dp.message(F.text == "🧭 Yo‘qolganlar")
@dp.message(F.text == "🧭 Потерянные")
async def button_lost(message: Message) -> None:
    await lost_goods(message)


@dp.message(F.text == "📄 FBO yuk xatlari")
@dp.message(F.text == "📄 Накладные FBO")
async def button_fbo_invoices(message: Message) -> None:
    await fbo_invoices(message)


@dp.message(F.text == "💎 Obuna")
@dp.message(F.text == "💎 Подписка")
async def button_subscription(message: Message) -> None:
    await subscribe(message)


@dp.message(F.text == "ℹ️ Yordam")
@dp.message(F.text == "ℹ️ Помощь")
@dp.message(F.text == "❓ Помощь")
async def button_help(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    if lang == "uz":
        text = (
            "ℹ️ <b>Yordam</b>\n\n"
            "Botni ulash uchun:\n"
            "1. <code>/video</code> — videoqo‘llanmani ko‘ring.\n"
            "2. <code>/connect</code> — API-kalitni yuboring.\n"
            "3. <b>💰 Savdo</b> yoki <b>📦 Ombor</b> bo‘limidan foydalaning.\n\n"
            "Foydali buyruqlar:\n"
            "• <code>/api_token</code> — API-kalit bo‘yicha yozma yo‘riqnoma\n"
            "• <code>/check</code> — ulanishni tekshirish\n"
            "• <code>/support</code> — yordam va administrator bilan aloqa"
        )
    else:
        text = (
            "ℹ️ <b>Помощь</b>\n\n"
            "Чтобы подключить бота:\n"
            "1. <code>/video</code> — посмотрите видеоинструкцию.\n"
            "2. <code>/connect</code> — отправьте API-ключ.\n"
            "3. Пользуйтесь разделами <b>💰 Продажи</b> и <b>📦 Склад</b>.\n\n"
            "Полезные команды:\n"
            "• <code>/api_token</code> — текстовая инструкция по API-ключу\n"
            "• <code>/check</code> — проверить подключение\n"
            "• <code>/support</code> — поддержка и связь с администратором"
        )
    await message.answer(text, reply_markup=help_links_markup(lang) or menu_for_message(message))


@dp.message(F.text == "🎥 Видеоинструкция")
@dp.message(F.text == "🎥 API ulash videosi")
async def button_video_instruction(message: Message) -> None:
    await video_instruction(message)


# Старые красивые кнопки оставлены для совместимости, если они остались у пользователя в Telegram.
@dp.message(F.text == "📊 Аналитика")
async def section_analytics(message: Message) -> None:
    await message.answer(tr_user(upsert_from_message(message), "main_menu"), reply_markup=menu_for_message(message))


@dp.message(F.text == "📦 Товары")
async def section_products(message: Message) -> None:
    await stock(message)


@dp.message(F.text == "🛒 Заказы/продажи")
async def section_orders(message: Message) -> None:
    await orders(message)


@dp.message(F.text == "🔔 Уведомления старое")
async def section_notifications(message: Message) -> None:
    await notify_status(message)


@dp.message(F.text == "⚙️ Настройки")
async def section_settings(message: Message) -> None:
    await status(message)


@dp.message(F.text == "⭐ Отзывы")
async def button_reviews(message: Message) -> None:
    await reviews(message)


@dp.message(F.text == "📈 Сводка")
@dp.message(F.text == "📈 Сводка FBO/FBS")
async def button_dashboard(message: Message) -> None:
    await dashboard(message)


@dp.message(F.text == "📊 Сводка заказов")
@dp.message(F.text == "📊 Сводка FBS/DBS")
async def button_orders_summary(message: Message) -> None:
    await orders_summary(message)


@dp.message(F.text == "💰 Продажи Finance")
async def button_sales(message: Message) -> None:
    await sales(message)


@dp.message(F.text == "📦 Все товары")
async def button_products(message: Message) -> None:
    await products(message)


@dp.message(F.text == "📊 Все остатки")
async def button_stock(message: Message) -> None:
    await stock(message)


@dp.message(F.text == "🏬 Остатки FBO")
async def button_fbo_stock(message: Message) -> None:
    await fbo(message)


@dp.message(F.text == "🚚 Остатки FBS/DBS")
async def button_fbs_stock(message: Message) -> None:
    await fbs(message)


@dp.message(F.text == "🛒 Новые заказы")
@dp.message(F.text == "🛒 FBS/DBS заказы")
async def button_orders(message: Message) -> None:
    await orders(message)


@dp.message(F.text == "📊 Excel hisobot")
@dp.message(F.text == "📊 Excel отчёт")
@dp.message(F.text == "📄 Excel-отчёт")
async def button_excel_report(message: Message) -> None:
    await report_excel(message)


@dp.message(F.text == "⚙️ Статус")
async def button_status(message: Message) -> None:
    await status(message)


@dp.message(F.text == "🏪 Do‘konlar")
@dp.message(F.text == "🏪 Магазины")
async def button_shops(message: Message) -> None:
    await shops(message)


@dp.message(F.text == "🔔 Новые заказы")
@dp.message(F.text == "🔔 FBS/DBS новые заказы")
async def button_notify_status(message: Message) -> None:
    await notify_status(message)


@dp.message(F.text == "💸 Новые продажи")
@dp.message(F.text == "💸 Продажи Finance")
async def button_sales_notify_status(message: Message) -> None:
    await sales_notify_status(message)


@dp.message(F.text == "📦 Изменение остатков")
@dp.message(F.text == "📦 Изменение FBO/FBS")
@dp.message(F.text == "📦 FBO/FBS движение")
@dp.message(F.text == "📦 FBO/FBS движение остатков")
async def button_stock_change_notify_status(message: Message) -> None:
    await stock_change_notify_status(message)


@dp.message(F.text == "📉 Низкие остатки")
@dp.message(F.text == "📉 Низкие остатки FBO/FBS")
async def button_lowstock_notify_status(message: Message) -> None:
    await lowstock_notify_status(message)


@dp.message(F.text == "❌ Нет в наличии")
async def button_outofstock_notify_status(message: Message) -> None:
    await outofstock_notify_status(message)





# --- PRO FEATURES: multi-shop, analytics, reports, reminders ---
def _shop_id_from_any(shop: Any) -> int | None:
    if isinstance(shop, dict):
        for key in ("shopId", "shop_id", "id", "value"):
            value = shop.get(key)
            try:
                if value not in (None, ""):
                    return int(value)
            except Exception:
                pass
        for value in shop.values():
            if isinstance(value, dict):
                found = _shop_id_from_any(value)
                if found is not None:
                    return found
    else:
        for attr in ("shop_id", "shopId", "id"):
            value = getattr(shop, attr, None)
            try:
                if value not in (None, ""):
                    return int(value)
            except Exception:
                pass
    return None


def _shop_name_from_any(shop: Any) -> str:
    if isinstance(shop, dict):
        value = pick(shop, "title", "name", "shopName", "storeName", "legalName", "displayName")
        if value not in (None, ""):
            return str(value)
        for v in shop.values():
            if isinstance(v, dict):
                nested = _shop_name_from_any(v)
                if nested != "":
                    return nested
    return ""


async def _user_shop_list(telegram_id: int, client: UzumClient | None = None) -> list[dict[str, Any]]:
    shops_raw = db.list_shops(telegram_id) or []
    if not shops_raw and client is not None:
        try:
            data = await client.get_shops()
            items = extract_items(data)
            encrypted = db.get_encrypted_token(telegram_id)
            if encrypted and items:
                db.save_connection(telegram_id, encrypted, items)
            shops_raw = db.list_shops(telegram_id) or items
        except Exception:
            shops_raw = []

    result: list[dict[str, Any]] = []
    seen: set[int] = set()
    for item in shops_raw:
        sid = _shop_id_from_any(item)
        if sid is None or sid in seen:
            continue
        seen.add(sid)
        result.append({"shop_id": sid, "name": _shop_name_from_any(item), "raw": item})
    return result


def _stock_row_title(row: dict[str, Any]) -> str:
    value = pick(row, "skuTitle", "sku_title", "title", "productTitle", "product_title", "name")
    if value in (None, ""):
        value = _deep_pick_value(row, ("skuTitle", "productTitle", "title", "name"))
    return str(value or "Без названия")


def _stock_row_sku(row: dict[str, Any]) -> str:
    value = pick(row, "sku", "skuId", "sku_id", "barcode", "offerId", "shopSku")
    if value in (None, ""):
        value = _deep_pick_value(row, ("sku", "skuId", "barcode", "offerId"))
    return str(value or "")


def _stock_row_total(row: dict[str, Any]) -> int:
    value = _num_from_value(pick(row, "total", "quantity", "available", "stock", "qty"))
    if value is None:
        value = _deep_pick_number(row, ("total", "quantity", "available", "stock", "qty"))
    return int(value or 0)


def _stock_row_price(row: dict[str, Any]) -> float:
    value = _num_from_value(pick(row, "sellPrice", "price", "purchasePrice", "oldPrice"))
    if value is None:
        value = _deep_pick_number(row, ("sellPrice", "price", "purchasePrice", "oldPrice"))
    return float(value or 0.0)


def _sale_match_keys(item: dict[str, Any]) -> set[str]:
    keys: set[str] = set()
    for value in (_finance_sku_title(item), _finance_title(item), str(_deep_pick_value(item, ("skuId", "sku", "barcode", "offerId")) or "")):
        value = str(value or "").strip().lower()
        if value and value != "-":
            keys.add(value)
    return keys


def _stock_match_keys(row: dict[str, Any]) -> set[str]:
    keys: set[str] = set()
    for value in (_stock_row_sku(row), _stock_row_title(row)):
        value = str(value or "").strip().lower()
        if value:
            keys.add(value)
    return keys


def _merge_noorza_stats(stats_list: list[dict[str, Any]]) -> dict[str, Any]:
    result = {
        "rows": 0.0,
        "units": 0.0,
        "returns": 0.0,
        "revenue": 0.0,
        "commission": 0.0,
        "logistics": 0.0,
        "payout_total": 0.0,
        "withdrawn": 0.0,
        "left_to_withdraw": 0.0,
        "statuses": {},
    }
    statuses: dict[str, int] = {}
    for stats in stats_list:
        for key in ("rows", "units", "returns", "revenue", "commission", "logistics", "payout_total", "withdrawn", "left_to_withdraw"):
            result[key] = float(result.get(key) or 0) + float(stats.get(key) or 0)
        for status, count in (stats.get("statuses") or {}).items():
            statuses[str(status)] = statuses.get(str(status), 0) + int(count or 0)
    result["statuses"] = statuses
    return result


def _format_all_shops_balance(days_title: str, shops_count: int, stats: dict[str, Any], per_shop: list[str]) -> str:
    text = (
        f"🌐 <b>Баланс по всем магазинам {escape(days_title)}</b>\n\n"
        f"🏪 Магазинов: <b>{shops_count}</b>\n"
        f"🛒 Продаж: <b>{int(stats['rows'])}</b>\n"
        f"📦 Товаров продано: <b>{float(stats['units']):.0f} шт.</b>\n"
        f"↩️ Возвратов: <b>{float(stats['returns']):.0f} шт.</b>\n\n"
        f"💵 Выручка: <b>{_format_money(float(stats['revenue']))}</b>\n"
        f"🏷 Комиссия Uzum: <b>{_format_money(float(stats['commission']))}</b>\n"
        f"🚚 Логистика: <b>{_format_money(float(stats['logistics']))}</b>\n\n"
        f"✅ К выплате: <b>{_format_money(float(stats['payout_total']))}</b>\n"
        f"💳 Уже выведено: <b>{_format_money(float(stats['withdrawn']))}</b>\n"
        f"🧾 Остаток к выплате: <b>{_format_money(float(stats['left_to_withdraw']))}</b>"
    )
    if per_shop:
        text += "\n\n<b>🏪 По магазинам:</b>\n" + "\n".join(per_shop[:20])
    return text


async def _all_shops_finance_stats(telegram_id: int, client: UzumClient, date_from: int, date_to: int) -> tuple[dict[str, Any], list[str], int]:
    shops_list = await _user_shop_list(telegram_id, client)
    if not shops_list:
        default_shop = db.get_default_shop_id(telegram_id)
        if default_shop:
            shops_list = [{"shop_id": int(default_shop), "name": "", "raw": {}}]
    stats_list: list[dict[str, Any]] = []
    per_shop: list[str] = []
    for shop in shops_list:
        sid = int(shop["shop_id"])
        try:
            rows, _ = await _load_finance_orders(client, sid, date_from_ms=date_from, date_to_ms=date_to, max_pages=5, page_size=100)
            stats = _build_noorza_today_stats(rows)
            stats_list.append(stats)
            name = f" — {escape(shop['name'])}" if shop.get("name") else ""
            per_shop.append(
                f"• <code>{sid}</code>{name}: {_format_money(float(stats['revenue']))}, "
                f"{float(stats['units']):.0f} шт., к выплате {_format_money(float(stats['payout_total']))}"
            )
            await asyncio.sleep(0.2)
        except Exception as e:
            logging.exception("All shops balance: failed for shop=%s", sid)
            per_shop.append(f"• <code>{sid}</code>: ошибка API — {escape(str(e)[:80])}")
    return _merge_noorza_stats(stats_list), per_shop, len(shops_list)


@dp.message(Command("balance_all", "all_balance", "allshops", "all_shops"))
async def balance_all_shops(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    if not await require_active_subscription(message, telegram_id):
        return
    client = get_uzum_for_user(telegram_id)
    if client is None:
        await message.answer("Сначала подключите Uzum API-токен: <code>/connect</code>", reply_markup=menu_for_message(message))
        return
    await message.answer("⌛ Считаю баланс по всем магазинам за 30 дней...", reply_markup=menu_for_message(message))
    try:
        date_from, date_to = _days_range_ms(30)
        stats, per_shop, shops_count = await _all_shops_finance_stats(telegram_id, client, date_from, date_to)
        await message.answer(_format_all_shops_balance("за 30 дней", shops_count, stats, per_shop), reply_markup=menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


async def _top_products_for_shop(client: UzumClient, shop_id: int, days: int) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    date_from, date_to = _days_range_ms(days)
    rows, _ = await _load_finance_orders(client, shop_id, date_from_ms=date_from, date_to_ms=date_to, max_pages=5, page_size=100)
    groups: dict[str, dict[str, Any]] = {}
    for item in rows:
        if _is_cancelled_status(_finance_status(item)):
            continue
        title = _finance_title(item)
        sku = _finance_sku_title(item)
        key = (sku or title or "-").strip().lower()
        if not key:
            key = title.strip().lower()
        entry = groups.setdefault(key, {"title": title, "sku": sku, "qty": 0.0, "revenue": 0.0, "payout": 0.0})
        gross = _finance_gross_revenue(item)
        commission = _finance_commission(item)
        logistics = _finance_logistics(item)
        payout_direct = _finance_payout_direct(item)
        payout = payout_direct if payout_direct is not None else max(0.0, gross - commission - logistics)
        entry["qty"] += _finance_qty(item)
        entry["revenue"] += gross
        entry["payout"] += max(0.0, payout)
    top = sorted(groups.values(), key=lambda x: float(x.get("revenue") or 0), reverse=True)
    return top, _build_noorza_today_stats(rows)


@dp.message(Command("top", "top_products"))
async def top_products(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    telegram_id = message.from_user.id if message.from_user else 0
    lang = get_user_language(telegram_id)
    days = TOP_PRODUCTS_DAYS
    await message.answer(f"⌛ {days} kunlik top tovarlarni hisoblayapman..." if lang == "uz" else f"⌛ Считаю топ товаров за {days} дней...", reply_markup=sales_menu_for_message(message))
    try:
        top, stats = await _top_products_for_shop(client, shop_id, days)
        if not top:
            text = f"🏆 <b>{days} kunlik top tovarlar</b>\n🏪 Do‘kon: <code>{shop_id}</code>\n\nSavdolar topilmadi." if lang == "uz" else f"🏆 <b>Топ товаров за {days} дней</b>\n🏪 Магазин: <code>{shop_id}</code>\n\nПродаж не найдено."
            await message.answer(text, reply_markup=sales_menu_for_message(message))
            return
        title = f"🏆 <b>{days} kunlik top tovarlar</b>" if lang == "uz" else f"🏆 <b>Топ товаров за {days} дней</b>"
        summary = [
            f"🏪 Do‘kon: <code>{shop_id}</code>" if lang == "uz" else f"🏪 Магазин: <code>{shop_id}</code>",
            f"📦 Sotilgan: <b>{float(stats['units']):.0f} dona</b>" if lang == "uz" else f"📦 Всего продано: <b>{float(stats['units']):.0f} шт.</b>",
            f"💵 Tushum: <b>{_format_money(float(stats['revenue']))}</b>" if lang == "uz" else f"💵 Выручка: <b>{_format_money(float(stats['revenue']))}</b>",
        ]
        items: list[str] = []
        for idx, item in enumerate(top, start=1):
            title_item = escape(_short_text(item.get("title"), 85))
            sku = escape(_short_text(item.get("sku"), 60))
            sku_line = f"\n🔖 SKU: <code>{sku}</code>" if sku and sku != "-" else ""
            if lang == "uz":
                items.append(
                    f"{idx}. <b>{title_item}</b>{sku_line}\n"
                    f"🔢 Soni: <b>{float(item.get('qty') or 0):.0f} dona</b> | "
                    f"💵 Tushum: <b>{_format_money(float(item.get('revenue') or 0))}</b> | "
                    f"✅ To‘lovga: <b>{_format_money(float(item.get('payout') or 0))}</b>"
                )
            else:
                items.append(
                    f"{idx}. <b>{title_item}</b>{sku_line}\n"
                    f"🔢 Продано: <b>{float(item.get('qty') or 0):.0f} шт.</b> | "
                    f"💵 Выручка: <b>{_format_money(float(item.get('revenue') or 0))}</b> | "
                    f"✅ К выплате: <b>{_format_money(float(item.get('payout') or 0))}</b>"
                )
        await send_paginated_list(message, kind="top", title=title, summary=summary, items=items, section="sales", reply_markup=sales_menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("deadstock", "no_sales", "stuck"))
async def dead_stock(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    telegram_id = message.from_user.id if message.from_user else 0
    lang = get_user_language(telegram_id)
    days = DEAD_STOCK_DAYS
    await message.answer(f"⌛ {days} kun sotilmagan tovarlarni qidiryapman..." if lang == "uz" else f"⌛ Ищу товары с остатком, но без продаж за {days} дней...", reply_markup=sales_menu_for_message(message))
    try:
        date_from, date_to = _days_range_ms(days)
        sales_rows, _ = await _load_finance_orders(client, shop_id, date_from_ms=date_from, date_to_ms=date_to, max_pages=5, page_size=100)
        sold_keys: set[str] = set()
        for item in sales_rows:
            if not _is_cancelled_status(_finance_status(item)):
                sold_keys.update(_sale_match_keys(item))
        stock_rows = await load_sku_rows(client, shop_id, max_pages=50)
        candidates: list[dict[str, Any]] = []
        for row in stock_rows:
            total = _stock_row_total(row)
            if total <= 0:
                continue
            keys = _stock_match_keys(row)
            if keys and not keys.intersection(sold_keys):
                price = _stock_row_price(row)
                candidates.append({"row": row, "total": total, "price": price, "value": total * price})
        candidates.sort(key=lambda x: float(x.get("value") or 0), reverse=True)
        if not candidates:
            text = f"🐢 <b>{days} kun sotilmagan tovarlar</b>\n🏪 Do‘kon: <code>{shop_id}</code>\n\nQoldiqda turib sotilmayotgan tovarlar topilmadi." if lang == "uz" else f"🐢 <b>Товары без продаж за {days} дней</b>\n🏪 Магазин: <code>{shop_id}</code>\n\nНе нашёл товаров с остатком и нулевыми продажами."
            await message.answer(text, reply_markup=sales_menu_for_message(message))
            return
        total_value = sum(float(x.get("value") or 0) for x in candidates)
        title = f"🐢 <b>{days} kun sotilmagan tovarlar</b>" if lang == "uz" else f"🐢 <b>Товары без продаж за {days} дней</b>"
        summary = [
            f"🏪 Do‘kon: <code>{shop_id}</code>" if lang == "uz" else f"🏪 Магазин: <code>{shop_id}</code>",
            f"📦 Pozitsiyalar: <b>{len(candidates)}</b>" if lang == "uz" else f"📦 Позиций: <b>{len(candidates)}</b>",
            f"💰 Taxminan muzlagan summa: <b>{_format_money(total_value)}</b>" if lang == "uz" else f"💰 Примерно заморожено: <b>{_format_money(total_value)}</b>",
        ]
        items: list[str] = []
        for idx, item in enumerate(candidates, start=1):
            row = item["row"]
            title_item = escape(_short_text(_stock_row_title(row), 85))
            sku = escape(_short_text(_stock_row_sku(row), 60))
            sku_line = f"\n🔖 SKU: <code>{sku}</code>" if sku else ""
            if lang == "uz":
                items.append(f"{idx}. <b>{title_item}</b>{sku_line}\n📦 Qoldiq: <b>{int(item['total'])} dona</b> | 💵 Narx: {_format_money(float(item['price']))} | 💰 Summa: <b>{_format_money(float(item['value']))}</b>")
            else:
                items.append(f"{idx}. <b>{title_item}</b>{sku_line}\n📦 Остаток: <b>{int(item['total'])} шт.</b> | 💵 Цена: {_format_money(float(item['price']))} | 💰 Сумма: <b>{_format_money(float(item['value']))}</b>")
        await send_paginated_list(message, kind="dead", title=title, summary=summary, items=items, section="sales", reply_markup=sales_menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("smart_lowstock", "forecast_stock"))
async def smart_lowstock(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    telegram_id = message.from_user.id if message.from_user else 0
    lang = get_user_language(telegram_id)
    await message.answer("⌛ Qoldiq necha kunga yetishini hisoblayapman..." if lang == "uz" else "⌛ Считаю, на сколько дней хватит остатков...", reply_markup=stock_menu_for_message(message))
    try:
        date_from, date_to = _days_range_ms(7)
        sales_rows, _ = await _load_finance_orders(client, shop_id, date_from_ms=date_from, date_to_ms=date_to, max_pages=5, page_size=100)
        sold_qty: dict[str, float] = {}
        for item in sales_rows:
            if _is_cancelled_status(_finance_status(item)):
                continue
            keys = _sale_match_keys(item)
            for key in keys:
                sold_qty[key] = sold_qty.get(key, 0.0) + _finance_qty(item)
        stock_rows = await load_sku_rows(client, shop_id, max_pages=50)
        alerts: list[dict[str, Any]] = []
        for row in stock_rows:
            total = _stock_row_total(row)
            if total <= 0:
                continue
            keys = _stock_match_keys(row)
            qty_7 = max([sold_qty.get(k, 0.0) for k in keys] or [0.0])
            avg_day = qty_7 / 7.0
            days_left = 9999.0 if avg_day <= 0 else total / avg_day
            if total <= LOW_STOCK_THRESHOLD or days_left <= SMART_LOW_STOCK_DAYS:
                alerts.append({"row": row, "total": total, "qty_7": qty_7, "days_left": days_left})
        alerts.sort(key=lambda x: (float(x.get("days_left") or 9999), int(x.get("total") or 0)))
        if not alerts:
            text = (
                f"⚠️ <b>Qoldiq prognozi</b>\n🏪 Do‘kon: <code>{shop_id}</code>\n\nKritik tovarlar topilmadi."
                if lang == "uz"
                else f"⚠️ <b>Прогноз остатков</b>\n🏪 Магазин: <code>{shop_id}</code>\n\nКритичных товаров не нашёл."
            )
            await message.answer(text, reply_markup=stock_menu_for_message(message))
            return
        title = "⚠️ <b>Qoldiq prognozi</b>" if lang == "uz" else "⚠️ <b>Прогноз остатков</b>"
        summary = [
            f"🏪 Do‘kon: <code>{shop_id}</code>" if lang == "uz" else f"🏪 Магазин: <code>{shop_id}</code>",
            f"Chegara: ≤ {LOW_STOCK_THRESHOLD} dona yoki {SMART_LOW_STOCK_DAYS} kundan kam" if lang == "uz" else f"Порог: ≤ {LOW_STOCK_THRESHOLD} шт. или хватит меньше чем на {SMART_LOW_STOCK_DAYS} дня",
        ]
        items: list[str] = []
        for idx, item in enumerate(alerts, start=1):
            row = item["row"]
            title_item = escape(_short_text(_stock_row_title(row), 85))
            sku = escape(_short_text(_stock_row_sku(row), 60))
            days_left = float(item["days_left"])
            if lang == "uz":
                days_text = "7 kunda savdo yo‘q" if days_left > 9000 else f"taxminan {days_left:.1f} kun"
                sku_line = f"\n🔖 SKU: <code>{sku}</code>" if sku else ""
                items.append(f"{idx}. <b>{title_item}</b>{sku_line}\n📦 Qoldiq: <b>{int(item['total'])} dona</b> | 7 kunda sotildi: <b>{float(item['qty_7']):.0f}</b> | Yetadi: <b>{days_text}</b>")
            else:
                days_text = "нет продаж за 7 дней" if days_left > 9000 else f"примерно на {days_left:.1f} дн."
                sku_line = f"\n🔖 SKU: <code>{sku}</code>" if sku else ""
                items.append(f"{idx}. <b>{title_item}</b>{sku_line}\n📦 Остаток: <b>{int(item['total'])} шт.</b> | Продажи за 7 дней: <b>{float(item['qty_7']):.0f}</b> | Хватит: <b>{days_text}</b>")
        await send_paginated_list(message, kind="forecast", title=title, summary=summary, items=items, section="stock", reply_markup=stock_menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


async def _build_morning_report_text(telegram_id: int, client: UzumClient) -> str:
    now_uzt = datetime.now(UZT)
    end = now_uzt.replace(hour=0, minute=0, second=0, microsecond=0)
    start = end - timedelta(days=1)
    date_from = int(start.timestamp() * 1000)
    date_to = int(end.timestamp() * 1000)
    stats, per_shop, shops_count = await _all_shops_finance_stats(telegram_id, client, date_from, date_to)
    return (
        "🌙 <b>Утренний отчёт Uzum</b>\n"
        f"За вчера: <b>{start.strftime('%d.%m.%Y')}</b>\n\n"
        + _format_all_shops_balance("за вчера", shops_count, stats, per_shop).replace("🌐 <b>Баланс по всем магазинам за вчера</b>\n\n", "")
        + "\n\n<i>Автоотчёт можно включить переменной DAILY_REPORTS=1.</i>"
    )


@dp.message(Command("morning_report", "daily_report"))
async def morning_report(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    if not await require_active_subscription(message, telegram_id):
        return
    client = get_uzum_for_user(telegram_id)
    if client is None:
        await message.answer("Сначала подключите Uzum API-токен: <code>/connect</code>", reply_markup=menu_for_message(message))
        return
    await message.answer("⌛ Готовлю утренний отчёт за вчера...", reply_markup=menu_for_message(message))
    try:
        await message.answer(await _build_morning_report_text(telegram_id, client), reply_markup=menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("extend1", "extend_month"))
async def admin_extend_1_month(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split()
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/extend1 TELEGRAM_ID</code>", reply_markup=menu_for_message(message))
        return
    target = int(arg[0])
    new_until = extend_subscription_days(target, 30)
    await message.answer(f"✅ Продлено на 1 месяц для <code>{target}</code>. До: <b>{_fmt_dt(new_until)}</b>", reply_markup=menu_for_message(message))


@dp.message(Command("extend3"))
async def admin_extend_3_months(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split()
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/extend3 TELEGRAM_ID</code>", reply_markup=menu_for_message(message))
        return
    target = int(arg[0])
    new_until = extend_subscription_days(target, 90)
    await message.answer(f"✅ Продлено на 3 месяца для <code>{target}</code>. До: <b>{_fmt_dt(new_until)}</b>", reply_markup=menu_for_message(message))


@dp.message(Command("extend6"))
async def admin_extend_6_months(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split()
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/extend6 TELEGRAM_ID</code>", reply_markup=menu_for_message(message))
        return
    target = int(arg[0])
    new_until = extend_subscription_days(target, 180)
    await message.answer(f"✅ Продлено на 6 месяцев для <code>{target}</code>. До: <b>{_fmt_dt(new_until)}</b>", reply_markup=menu_for_message(message))




# --- Юнит-экономика ---
def _unit_group_key(item: dict[str, Any]) -> str:
    sku = _finance_sku_title(item)
    if sku and sku != "-":
        return _unit_sku_key(sku)
    return _unit_sku_key(_finance_title(item))


def _unit_cost_lookup(costs: dict[str, dict[str, Any]], item: dict[str, Any]) -> float | None:
    for key in (_finance_sku_title(item), _finance_title(item), str(_deep_pick_value(item, ("skuId", "sku", "barcode", "offerId")) or "")):
        sku_key = _unit_sku_key(key)
        if sku_key and sku_key in costs:
            try:
                return float(costs[sku_key].get("cost") or 0)
            except Exception:
                return None
    return None


async def _unit_economy_for_shop(client: UzumClient, telegram_id: int, shop_id: int, days: int = 30) -> tuple[list[dict[str, Any]], dict[str, Any], int]:
    date_from, date_to = _days_range_ms(days)
    rows, _ = await _load_finance_orders(client, shop_id, date_from_ms=date_from, date_to_ms=date_to, max_pages=5, page_size=100)
    costs = get_unit_cost_map(telegram_id, shop_id)
    groups: dict[str, dict[str, Any]] = {}
    for item in rows:
        if _is_cancelled_status(_finance_status(item)):
            continue
        key = _unit_group_key(item)
        if not key:
            continue
        qty = _finance_qty(item)
        revenue = _finance_gross_revenue(item)
        commission = _finance_commission(item)
        logistics = _finance_logistics(item)
        payout_direct = _finance_payout_direct(item)
        payout = payout_direct if payout_direct is not None else max(0.0, revenue - commission - logistics)
        cost_per_unit = _unit_cost_lookup(costs, item)
        entry = groups.setdefault(key, {
            "sku": _finance_sku_title(item),
            "title": _finance_title(item),
            "qty": 0.0,
            "revenue": 0.0,
            "commission": 0.0,
            "logistics": 0.0,
            "payout": 0.0,
            "cost_per_unit": cost_per_unit,
            "cost_total": 0.0,
            "profit": None,
        })
        if entry.get("cost_per_unit") is None and cost_per_unit is not None:
            entry["cost_per_unit"] = cost_per_unit
        entry["qty"] += qty
        entry["revenue"] += revenue
        entry["commission"] += commission
        entry["logistics"] += logistics
        entry["payout"] += max(0.0, payout)
        if cost_per_unit is not None:
            entry["cost_total"] += float(cost_per_unit) * qty
    for entry in groups.values():
        if entry.get("cost_per_unit") is not None:
            entry["profit"] = float(entry.get("payout") or 0) - float(entry.get("cost_total") or 0)
            revenue = float(entry.get("revenue") or 0)
            entry["margin"] = (float(entry["profit"]) / revenue * 100.0) if revenue > 0 else 0.0
        else:
            entry["margin"] = None
    top = sorted(groups.values(), key=lambda x: float(x.get("revenue") or 0), reverse=True)
    return top, _build_noorza_today_stats(rows), len(costs)


def _format_unit_economy(shop_id: int, days: int, rows: list[dict[str, Any]], stats: dict[str, Any], saved_costs: int, lang: str = "ru") -> str:
    lang = normalize_lang(lang)
    if lang == "uz":
        if not rows:
            return (
                f"🧾 <b>Unit iqtisodiyot</b>\n🏪 Do‘kon: <code>{shop_id}</code>\n\n"
                "Savdolar topilmadi. Avval 30 kunlik savdolar bo‘yicha ma’lumot bo‘lishi kerak."
            )
        total_known_profit = sum(float(x.get("profit") or 0) for x in rows if x.get("profit") is not None)
        known_items = sum(1 for x in rows if x.get("profit") is not None)
        lines = [
            f"🧾 <b>Unit iqtisodiyot — {days} kun</b>",
            f"🏪 Do‘kon: <code>{shop_id}</code>",
            f"📦 Sotilgan: <b>{float(stats.get('units') or 0):.0f} dona</b>",
            f"💵 Tushum: <b>{_format_money(float(stats.get('revenue') or 0))}</b>",
            f"✅ To‘lovga: <b>{_format_money(float(stats.get('payout_total') or 0))}</b>",
            f"💾 Kiritilgan tannarxlar: <b>{saved_costs}</b>",
        ]
        if known_items:
            lines.append(f"💰 Taxminiy sof foyda: <b>{_format_money(total_known_profit)}</b>")
        lines.append("\n<b>Top tovarlar:</b>")
        for idx, item in enumerate(rows[:10], start=1):
            sku = escape(_short_text(item.get("sku"), 55))
            title = escape(_short_text(item.get("title"), 70))
            cost = item.get("cost_per_unit")
            if cost is None:
                hint = f"\n   ⚠️ Tannarx kiritilmagan: <code>/cost {sku} 60000</code>"
            else:
                profit = float(item.get("profit") or 0)
                margin = float(item.get("margin") or 0)
                hint = f"\n   🧾 Tannarx: <b>{_format_money(float(cost))}</b> | 💰 Foyda: <b>{_format_money(profit)}</b> | 📈 Marja: <b>{margin:.1f}%</b>"
            lines.append(
                f"{idx}. <b>{title}</b>\n"
                f"   🔖 SKU: <code>{sku}</code>\n"
                f"   🔢 Soni: <b>{float(item.get('qty') or 0):.0f} dona</b> | "
                f"💵 Tushum: <b>{_format_money(float(item.get('revenue') or 0))}</b> | "
                f"✅ To‘lovga: <b>{_format_money(float(item.get('payout') or 0))}</b>" + hint
            )
        lines.append("\nTannarx qo‘shish: <code>/cost SKU 60000</code>")
        return "\n\n".join(lines)

    if not rows:
        return (
            f"🧾 <b>Юнит-экономика</b>\n🏪 Магазин: <code>{shop_id}</code>\n\n"
            "Продаж не найдено. Сначала должны быть продажи за выбранный период."
        )
    total_known_profit = sum(float(x.get("profit") or 0) for x in rows if x.get("profit") is not None)
    known_items = sum(1 for x in rows if x.get("profit") is not None)
    lines = [
        f"🧾 <b>Юнит-экономика за {days} дней</b>",
        f"🏪 Магазин: <code>{shop_id}</code>",
        f"📦 Продано: <b>{float(stats.get('units') or 0):.0f} шт.</b>",
        f"💵 Выручка: <b>{_format_money(float(stats.get('revenue') or 0))}</b>",
        f"✅ К выплате: <b>{_format_money(float(stats.get('payout_total') or 0))}</b>",
        f"💾 Себестоимостей сохранено: <b>{saved_costs}</b>",
    ]
    if known_items:
        lines.append(f"💰 Примерная чистая прибыль: <b>{_format_money(total_known_profit)}</b>")
    lines.append("\n<b>Топ товаров:</b>")
    for idx, item in enumerate(rows[:10], start=1):
        sku = escape(_short_text(item.get("sku"), 55))
        title = escape(_short_text(item.get("title"), 70))
        cost = item.get("cost_per_unit")
        if cost is None:
            hint = f"\n   ⚠️ Себестоимость не указана: <code>/cost {sku} 60000</code>"
        else:
            profit = float(item.get("profit") or 0)
            margin = float(item.get("margin") or 0)
            hint = f"\n   🧾 Себестоимость: <b>{_format_money(float(cost))}</b> | 💰 Прибыль: <b>{_format_money(profit)}</b> | 📈 Маржа: <b>{margin:.1f}%</b>"
        lines.append(
            f"{idx}. <b>{title}</b>\n"
            f"   🔖 SKU: <code>{sku}</code>\n"
            f"   🔢 Кол-во: <b>{float(item.get('qty') or 0):.0f} шт.</b> | "
            f"💵 Выручка: <b>{_format_money(float(item.get('revenue') or 0))}</b> | "
            f"✅ К выплате: <b>{_format_money(float(item.get('payout') or 0))}</b>" + hint
        )
    lines.append("\nДобавить себестоимость: <code>/cost SKU 60000</code>")
    return "\n\n".join(lines)


@dp.message(Command("unit", "unit_economy", "profit"))
async def unit_economy(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    lang = get_user_language(telegram_id)
    await message.answer("⌛ Hisoblayapman..." if lang == "uz" else "⌛ Считаю юнит-экономику...", reply_markup=sales_menu_for_message(message))
    try:
        rows, stats, saved_costs = await _unit_economy_for_shop(client, telegram_id, shop_id, days=30)
        if not rows:
            text = f"🧾 <b>Unit iqtisodiyot</b>\n🏪 Do‘kon: <code>{shop_id}</code>\n\nSavdolar topilmadi." if lang == "uz" else f"🧾 <b>Юнит-экономика</b>\n🏪 Магазин: <code>{shop_id}</code>\n\nПродаж не найдено."
            await message.answer(text, reply_markup=sales_menu_for_message(message))
            return
        total_known_profit = sum(float(x.get("profit") or 0) for x in rows if x.get("profit") is not None)
        known_items = sum(1 for x in rows if x.get("profit") is not None)
        title = "🧾 <b>Unit iqtisodiyot — 30 kun</b>" if lang == "uz" else "🧾 <b>Юнит-экономика за 30 дней</b>"
        summary = [
            f"🏪 Do‘kon: <code>{shop_id}</code>" if lang == "uz" else f"🏪 Магазин: <code>{shop_id}</code>",
            f"📦 Sotilgan: <b>{float(stats.get('units') or 0):.0f} dona</b> | 💵 Tushum: <b>{_format_money(float(stats.get('revenue') or 0))}</b>" if lang == "uz" else f"📦 Продано: <b>{float(stats.get('units') or 0):.0f} шт.</b> | 💵 Выручка: <b>{_format_money(float(stats.get('revenue') or 0))}</b>",
            f"💾 Tannarxlar: <b>{saved_costs}</b>" if lang == "uz" else f"💾 Себестоимостей: <b>{saved_costs}</b>",
        ]
        if known_items:
            summary.append(f"💰 Taxminiy sof foyda: <b>{_format_money(total_known_profit)}</b>" if lang == "uz" else f"💰 Примерная чистая прибыль: <b>{_format_money(total_known_profit)}</b>")
        items: list[str] = []
        for idx, item in enumerate(rows, start=1):
            sku = escape(_short_text(item.get("sku"), 55))
            title_item = escape(_short_text(item.get("title"), 70))
            cost = item.get("cost_per_unit")
            if cost is None:
                hint = f"\n⚠️ Tannarx kiritilmagan: <code>/cost {sku} 60000</code>" if lang == "uz" else f"\n⚠️ Себестоимость не указана: <code>/cost {sku} 60000</code>"
            else:
                profit = float(item.get("profit") or 0)
                margin = float(item.get("margin") or 0)
                hint = f"\n🧾 Tannarx: <b>{_format_money(float(cost))}</b> | 💰 Foyda: <b>{_format_money(profit)}</b> | 📈 Marja: <b>{margin:.1f}%</b>" if lang == "uz" else f"\n🧾 Себестоимость: <b>{_format_money(float(cost))}</b> | 💰 Прибыль: <b>{_format_money(profit)}</b> | 📈 Маржа: <b>{margin:.1f}%</b>"
            if lang == "uz":
                items.append(f"{idx}. <b>{title_item}</b>\n🔖 SKU: <code>{sku}</code>\n🔢 Soni: <b>{float(item.get('qty') or 0):.0f} dona</b> | 💵 Tushum: <b>{_format_money(float(item.get('revenue') or 0))}</b> | ✅ To‘lovga: <b>{_format_money(float(item.get('payout') or 0))}</b>{hint}")
            else:
                items.append(f"{idx}. <b>{title_item}</b>\n🔖 SKU: <code>{sku}</code>\n🔢 Кол-во: <b>{float(item.get('qty') or 0):.0f} шт.</b> | 💵 Выручка: <b>{_format_money(float(item.get('revenue') or 0))}</b> | ✅ К выплате: <b>{_format_money(float(item.get('payout') or 0))}</b>{hint}")
        await send_paginated_list(message, kind="unit", title=title, summary=summary, items=items, section="sales", reply_markup=sales_menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("cost", "setcost", "set_cost"))
async def set_cost_command(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    req = await require_connection(message)
    if req is None:
        return
    _, _, shop_id = req
    lang = get_user_language(telegram_id)
    parsed = _parse_cost_command_args(message.text or "")
    if not parsed:
        text = (
            "Tannarxni shunday kiriting:\n<code>/cost SKU 60000</code>\n\nMasalan:\n<code>/cost NOORZA-NR751-BEJEV-XXL 60000</code>"
            if lang == "uz" else
            "Укажите себестоимость так:\n<code>/cost SKU 60000</code>\n\nНапример:\n<code>/cost NOORZA-NR751-БЕЖЕВ-XXL 60000</code>"
        )
        await message.answer(text, reply_markup=menu_for_message(message))
        return
    sku, cost = parsed
    save_unit_cost(telegram_id, shop_id, sku, cost, title=sku)
    if lang == "uz":
        await message.answer(f"✅ Saqlandi\n🔖 SKU: <code>{escape(sku)}</code>\n🧾 Tannarx: <b>{_format_money(cost)}</b>", reply_markup=menu_for_message(message))
    else:
        await message.answer(f"✅ Сохранено\n🔖 SKU: <code>{escape(sku)}</code>\n🧾 Себестоимость: <b>{_format_money(cost)}</b>", reply_markup=menu_for_message(message))


@dp.message(Command("costs"))
async def costs_command(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    req = await require_connection(message)
    if req is None:
        return
    _, _, shop_id = req
    lang = get_user_language(telegram_id)
    rows = list_unit_costs(telegram_id, shop_id, limit=50)
    if not rows:
        text = "Hali tannarxlar kiritilmagan. Qo‘shish: <code>/cost SKU 60000</code>" if lang == "uz" else "Себестоимость ещё не указана. Добавить: <code>/cost SKU 60000</code>"
        await message.answer(text, reply_markup=menu_for_message(message))
        return
    title = "🧾 <b>Saqlangan tannarxlar</b>" if lang == "uz" else "🧾 <b>Сохранённая себестоимость</b>"
    lines = [title, f"🏪 <code>{shop_id}</code>"]
    for r in rows:
        lines.append(f"• <code>{escape(str(r.get('sku_key') or ''))}</code> — <b>{_format_money(float(r.get('cost') or 0))}</b>")
    await message.answer("\n".join(lines), reply_markup=menu_for_message(message))


@dp.message(Command("delcost", "delete_cost"))
async def delete_cost_command(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    req = await require_connection(message)
    if req is None:
        return
    _, _, shop_id = req
    sku = parse_args(message.text or "").strip()
    lang = get_user_language(telegram_id)
    if not sku:
        await message.answer("Напишите так: <code>/delcost SKU</code>", reply_markup=menu_for_message(message))
        return
    ok = delete_unit_cost(telegram_id, shop_id, sku)
    if lang == "uz":
        await message.answer("✅ O‘chirildi" if ok else "Topilmadi", reply_markup=menu_for_message(message))
    else:
        await message.answer("✅ Удалено" if ok else "Не найдено", reply_markup=menu_for_message(message))



def _build_cost_template(path: str, lang: str = "ru") -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Себестоимость" if lang != "uz" else "Tannarx"
    headers = ["SKU", "Себестоимость", "Название (необязательно)"] if lang != "uz" else ["SKU", "Tannarx", "Nomi (ixtiyoriy)"]
    ws.append(headers)
    ws.append(["NOORZA-NR751-BEJEV-XXL", 60000, "Пример товара" if lang != "uz" else "Tovar namunasi"])
    ws.append(["NOORZA-KOR101-SERIY", 45000, ""])
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 38
    ws.freeze_panes = "A2"
    wb.save(path)


def _normalize_header(value: Any) -> str:
    text = str(value or "").strip().lower().replace("ё", "е")
    text = text.replace(" ", "_").replace("-", "_")
    return text


def _parse_excel_cost_value(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value) if float(value) >= 0 else None
    s = str(value).strip()
    if not s:
        return None
    s = s.replace("сум", "").replace("so'm", "").replace("so‘m", "").replace("uzs", "")
    s = s.replace(" ", "").replace("\u00a0", "").replace(",", ".")
    try:
        val = float(s)
        return val if val >= 0 else None
    except Exception:
        return None


def _detect_cost_columns(ws) -> tuple[int, int, int | None, int]:
    header_values = [cell.value for cell in ws[1]]
    normalized = [_normalize_header(v) for v in header_values]
    sku_col = cost_col = title_col = None
    for idx, h in enumerate(normalized, start=1):
        if h in {"sku", "артикул", "sku_id", "sku_title", "шк", "barcode", "offerid", "offer_id"}:
            sku_col = idx
        if h in {"себестоимость", "sebestoimost", "tannarx", "cost", "purchase_price", "закуп", "закупочная_цена", "tan_narx"}:
            cost_col = idx
        if h in {"название", "name", "title", "nomi", "товар", "product", "product_title", "название_(необязательно)", "nomi_(ixtiyoriy)"}:
            title_col = idx
    if sku_col and cost_col:
        return sku_col, cost_col, title_col, 2
    return 1, 2, 3, 1


def _parse_costs_workbook(path: str) -> tuple[list[dict[str, Any]], list[str]]:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    sku_col, cost_col, title_col, start_row = _detect_cost_columns(ws)
    imported: list[dict[str, Any]] = []
    errors: list[str] = []
    seen: set[str] = set()
    for row_idx, row in enumerate(ws.iter_rows(min_row=start_row, values_only=True), start=start_row):
        sku_val = row[sku_col - 1] if len(row) >= sku_col else None
        cost_val = row[cost_col - 1] if len(row) >= cost_col else None
        title_val = row[title_col - 1] if title_col and len(row) >= title_col else ""
        sku = str(sku_val or "").strip()
        if not sku:
            continue
        cost = _parse_excel_cost_value(cost_val)
        if cost is None:
            errors.append(f"Строка {row_idx}: неверная себестоимость для {sku}")
            continue
        key = _unit_sku_key(sku)
        if not key:
            errors.append(f"Строка {row_idx}: неверный SKU")
            continue
        if key in seen:
            continue
        seen.add(key)
        imported.append({"sku": sku, "cost": float(cost), "title": str(title_val or "").strip()})
    return imported, errors


@dp.message(Command("cost_template", "costs_template", "template_costs"))
async def cost_template_command(message: Message, state: FSMContext | None = None) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    if state is not None:
        await state.set_state(CostImportStates.waiting_for_file)
    filename = f"unit_costs_template_{telegram_id}.xlsx"
    file_path = str(Path(tempfile.gettempdir()) / filename)
    _build_cost_template(file_path, lang=lang)
    if lang == "uz":
        text = (
            "📥 <b>Tannarxlarni Excel orqali yuklash</b>\n\n"
            "1. Shablonni yuklab oling.\n"
            "2. SKU va tannarxlarni to‘ldiring.\n"
            "3. Tayyor Excel faylni shu chatga yuboring.\n\n"
            "Ustunlar: <b>SKU</b> va <b>Tannarx</b>.\n"
            "Bekor qilish: <code>/cancel</code>"
        )
    else:
        text = (
            "📥 <b>Загрузка себестоимости через Excel</b>\n\n"
            "1. Скачайте шаблон.\n"
            "2. Заполните SKU и себестоимость.\n"
            "3. Отправьте готовый Excel-файл сюда в чат.\n\n"
            "Обязательные колонки: <b>SKU</b> и <b>Себестоимость</b>.\n"
            "Отмена: <code>/cancel</code>"
        )
    await message.answer(text, reply_markup=menu_for_message(message))
    await message.answer_document(FSInputFile(file_path, filename="unit_costs_template.xlsx"))


@dp.message(Command("import_costs", "upload_costs"))
async def import_costs_command(message: Message, state: FSMContext) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    await state.set_state(CostImportStates.waiting_for_file)
    text = (
        "📥 Excel faylni yuboring. Shablon kerak bo‘lsa: <code>/cost_template</code>\nBekor qilish: <code>/cancel</code>"
        if lang == "uz"
        else "📥 Отправьте Excel-файл с себестоимостью. Шаблон: <code>/cost_template</code>\nОтмена: <code>/cancel</code>"
    )
    await message.answer(text, reply_markup=menu_for_message(message))


@dp.message(CostImportStates.waiting_for_file, F.document)
async def receive_costs_excel(message: Message, state: FSMContext) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    req = await require_connection(message)
    if req is None:
        return
    _, _, shop_id = req
    document = message.document
    if document is None:
        return
    filename = (document.file_name or "").lower()
    if not (filename.endswith(".xlsx") or filename.endswith(".xlsm")):
        await message.answer(
            "Excel fayl .xlsx formatida bo‘lishi kerak." if lang == "uz" else "Нужен Excel-файл в формате .xlsx.",
            reply_markup=menu_for_message(message),
        )
        return
    tmp_path = str(Path(tempfile.gettempdir()) / f"costs_{telegram_id}_{int(datetime.now().timestamp())}.xlsx")
    try:
        await bot.download(document, destination=tmp_path)
        items, errors = _parse_costs_workbook(tmp_path)
        if not items:
            await message.answer(
                "Faylda saqlash uchun SKU va tannarx topilmadi." if lang == "uz" else "В файле не нашёл SKU и себестоимость для сохранения.",
                reply_markup=menu_for_message(message),
            )
            return
        for item in items:
            save_unit_cost(telegram_id, int(shop_id), item["sku"], float(item["cost"]), title=item.get("title") or item["sku"])
        await state.clear()
        preview = "\n".join([f"• <code>{escape(i['sku'])}</code> — <b>{_format_money(float(i['cost']))}</b>" for i in items[:10]])
        more = max(0, len(items) - 10)
        if lang == "uz":
            text = f"✅ <b>Tannarxlar saqlandi</b>\n\n🏪 Do‘kon: <code>{shop_id}</code>\n📦 Saqlandi: <b>{len(items)}</b> SKU"
            if preview:
                text += "\n\n" + preview
            if more:
                text += f"\n...yana {more} ta"
            text += "\n\nEndi <b>🧾 Unit iqtisodiyot</b> yoki <b>💰 Foyda</b> bo‘limini tekshiring."
        else:
            text = f"✅ <b>Себестоимость сохранена</b>\n\n🏪 Магазин: <code>{shop_id}</code>\n📦 Сохранено: <b>{len(items)}</b> SKU"
            if preview:
                text += "\n\n" + preview
            if more:
                text += f"\n...ещё {more} шт."
            text += "\n\nТеперь проверьте <b>🧾 Юнит-экономику</b> или <b>💰 Прибыль</b>."
        if errors:
            text += ("\n\n⚠️ Ошибки: " if lang != "uz" else "\n\n⚠️ Xatolar: ") + str(len(errors))
            text += "\n" + "\n".join(escape(e) for e in errors[:5])
        await message.answer(text, reply_markup=sales_menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)
    finally:
        try:
            Path(tmp_path).unlink(missing_ok=True)
        except Exception:
            pass


@dp.message(CostImportStates.waiting_for_file)
async def receive_costs_excel_wrong(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    await message.answer(
        "Excel faylni yuboring yoki <code>/cancel</code> bosing." if lang == "uz" else "Отправьте Excel-файл или нажмите <code>/cancel</code>.",
        reply_markup=menu_for_message(message),
    )


def _profit_summary_from_unit_rows(rows: list[dict[str, Any]], stats: dict[str, Any]) -> dict[str, Any]:
    cost_total = sum(float(r.get("cost_total") or 0) for r in rows if r.get("cost_per_unit") is not None)
    known_profit = sum(float(r.get("profit") or 0) for r in rows if r.get("cost_per_unit") is not None)
    known_revenue = sum(float(r.get("revenue") or 0) for r in rows if r.get("cost_per_unit") is not None)
    missing = [r for r in rows if r.get("cost_per_unit") is None]
    margin = (known_profit / known_revenue * 100.0) if known_revenue > 0 else 0.0
    return {
        "cost_total": cost_total,
        "profit": known_profit,
        "margin": margin,
        "missing_count": len(missing),
        "known_revenue": known_revenue,
        "missing": missing,
    }


def _format_profit_report(shop_id: int, rows: list[dict[str, Any]], stats: dict[str, Any], lang: str = "ru") -> str:
    summary = _profit_summary_from_unit_rows(rows, stats)
    top_profit = sorted([r for r in rows if r.get("cost_per_unit") is not None], key=lambda r: float(r.get("profit") or 0), reverse=True)[:10]
    if lang == "uz":
        lines = [
            "💰 <b>30 kunlik foyda</b>",
            f"🏪 Do‘kon: <code>{shop_id}</code>",
            "",
            f"💵 Tushum: <b>{_format_money(float(stats.get('revenue') or 0))}</b>",
            f"🏷 Uzum komissiyasi: <b>{_format_money(float(stats.get('commission') or 0))}</b>",
            f"🚚 Logistika: <b>{_format_money(float(stats.get('logistics') or 0))}</b>",
            f"📦 Tannarx: <b>{_format_money(float(summary['cost_total']))}</b>",
            "",
            f"✅ To‘lovga: <b>{_format_money(float(stats.get('payout_total') or 0))}</b>",
            f"💰 Sof foyda: <b>{_format_money(float(summary['profit']))}</b>",
            f"📈 Marja: <b>{float(summary['margin']):.1f}%</b>",
        ]
        if summary["missing_count"]:
            lines.append(f"\n⚠️ Tannarx kiritilmagan SKU: <b>{summary['missing_count']}</b>")
            lines.append("Tannarx yuklash: <code>/cost_template</code>")
        if top_profit:
            lines.append("\n🏆 <b>Foyda bo‘yicha top tovarlar:</b>")
            for idx, r in enumerate(top_profit, start=1):
                lines.append(f"{idx}. {escape(_short_text(str(r.get('title') or r.get('sku') or '-'), 55))} — <b>{_format_money(float(r.get('profit') or 0))}</b>")
        return "\n".join(lines)
    lines = [
        "💰 <b>Прибыль за 30 дней</b>",
        f"🏪 Магазин: <code>{shop_id}</code>",
        "",
        f"💵 Выручка: <b>{_format_money(float(stats.get('revenue') or 0))}</b>",
        f"🏷 Комиссия Uzum: <b>{_format_money(float(stats.get('commission') or 0))}</b>",
        f"🚚 Логистика: <b>{_format_money(float(stats.get('logistics') or 0))}</b>",
        f"📦 Себестоимость: <b>{_format_money(float(summary['cost_total']))}</b>",
        "",
        f"✅ К выплате: <b>{_format_money(float(stats.get('payout_total') or 0))}</b>",
        f"💰 Чистая прибыль: <b>{_format_money(float(summary['profit']))}</b>",
        f"📈 Маржа: <b>{float(summary['margin']):.1f}%</b>",
    ]
    if summary["missing_count"]:
        lines.append(f"\n⚠️ Без себестоимости: <b>{summary['missing_count']}</b> SKU")
        lines.append("Загрузить себестоимость: <code>/cost_template</code>")
    if top_profit:
        lines.append("\n🏆 <b>Топ товаров по прибыли:</b>")
        for idx, r in enumerate(top_profit, start=1):
            lines.append(f"{idx}. {escape(_short_text(str(r.get('title') or r.get('sku') or '-'), 55))} — <b>{_format_money(float(r.get('profit') or 0))}</b>")
    return "\n".join(lines)


@dp.message(Command("profit", "unit_profit", "profit_report"))
async def profit_report(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    telegram_id, client, shop_id = req
    lang = get_user_language(telegram_id)
    await message.answer("⌛ Foydani hisoblayapman..." if lang == "uz" else "⌛ Считаю прибыль...", reply_markup=sales_menu_for_message(message))
    try:
        rows, stats, _ = await _unit_economy_for_shop(client, telegram_id, shop_id, days=30)
        summary_stats = _profit_summary_from_unit_rows(rows, stats)
        known = sorted([r for r in rows if r.get("cost_per_unit") is not None], key=lambda r: float(r.get("profit") or 0), reverse=True)
        title = "💰 <b>30 kunlik foyda</b>" if lang == "uz" else "💰 <b>Прибыль за 30 дней</b>"
        summary = [
            f"🏪 Do‘kon: <code>{shop_id}</code>" if lang == "uz" else f"🏪 Магазин: <code>{shop_id}</code>",
            f"💵 Tushum: <b>{_format_money(float(stats.get('revenue') or 0))}</b> | 📦 Tannarx: <b>{_format_money(float(summary_stats['cost_total']))}</b>" if lang == "uz" else f"💵 Выручка: <b>{_format_money(float(stats.get('revenue') or 0))}</b> | 📦 Себестоимость: <b>{_format_money(float(summary_stats['cost_total']))}</b>",
            f"💰 Sof foyda: <b>{_format_money(float(summary_stats['profit']))}</b> | 📈 Marja: <b>{float(summary_stats['margin']):.1f}%</b>" if lang == "uz" else f"💰 Чистая прибыль: <b>{_format_money(float(summary_stats['profit']))}</b> | 📈 Маржа: <b>{float(summary_stats['margin']):.1f}%</b>",
        ]
        if summary_stats["missing_count"]:
            summary.append(f"⚠️ Tannarx kiritilmagan SKU: <b>{summary_stats['missing_count']}</b>" if lang == "uz" else f"⚠️ Без себестоимости: <b>{summary_stats['missing_count']}</b> SKU")
        items: list[str] = []
        for idx, r in enumerate(known, start=1):
            title_item = escape(_short_text(str(r.get("title") or r.get("sku") or "-"), 70))
            sku = escape(_short_text(str(r.get("sku") or ""), 55))
            profit = float(r.get("profit") or 0)
            margin = float(r.get("margin") or 0)
            items.append((f"{idx}. <b>{title_item}</b>\n🔖 SKU: <code>{sku}</code>\n💰 Foyda: <b>{_format_money(profit)}</b> | 📈 Marja: <b>{margin:.1f}%</b>" if lang == "uz" else f"{idx}. <b>{title_item}</b>\n🔖 SKU: <code>{sku}</code>\n💰 Прибыль: <b>{_format_money(profit)}</b> | 📈 Маржа: <b>{margin:.1f}%</b>"))
        if not items:
            items = ["Tannarx kiritilgan savdolar hali yo‘q. Tannarx yuklash: /cost_template" if lang == "uz" else "Пока нет продаж с указанной себестоимостью. Загрузите себестоимость: /cost_template"]
        await send_paginated_list(message, kind="profit", title=title, summary=summary, items=items, section="sales", reply_markup=sales_menu_for_message(message))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(F.text == "💰 Foyda")
@dp.message(F.text == "💰 Прибыль")
async def button_profit_report_near_unit(message: Message) -> None:
    await profit_report(message)


@dp.message(F.text == "📥 Tannarx Excel")
@dp.message(F.text == "📥 Себестоимость Excel")
async def button_costs_excel_near_unit(message: Message, state: FSMContext) -> None:
    await cost_template_command(message, state)


@dp.message(F.text == "🧾 Unit iqtisodiyot")
@dp.message(F.text == "🧾 Юнит-экономика")
async def button_unit_economy(message: Message) -> None:
    await unit_economy(message)

@dp.message(F.text == "🌐 Barcha do‘konlar")
@dp.message(F.text == "🌐 Все магазины")
async def button_all_shops(message: Message) -> None:
    await balance_all_shops(message)


@dp.message(F.text == "🏆 Top tovarlar")
@dp.message(F.text == "🏆 Топ товаров")
async def button_top_products(message: Message) -> None:
    await top_products(message)


@dp.message(F.text == "🐢 Sotilmayapti")
@dp.message(F.text == "🐢 Не продаётся")
async def button_dead_stock(message: Message) -> None:
    await dead_stock(message)


@dp.message(F.text == "🌙 Ertalabki hisobot")
@dp.message(F.text == "🌙 Утренний отчёт")
async def button_morning_report(message: Message) -> None:
    await morning_report(message)


@dp.message(F.text == "⚠️ Qoldiq prognozi")
@dp.message(F.text == "⚠️ Прогноз остатков")
async def button_smart_lowstock(message: Message) -> None:
    await smart_lowstock(message)


_daily_report_sent: set[tuple[int, str]] = set()
_subscription_reminder_sent: set[tuple[int, str]] = set()


def _connected_users_basic() -> list[dict[str, Any]]:
    with db.connect() as conn:
        rows = conn.execute(
            """
            SELECT telegram_id, uzum_token_encrypted, default_shop_id
            FROM users
            WHERE uzum_token_encrypted IS NOT NULL
            """
        ).fetchall()
    return [dict(row) for row in rows if has_active_subscription(int(row["telegram_id"]))]


async def daily_report_loop() -> None:
    await asyncio.sleep(90)
    logging.info("Daily report loop started. Enabled: %s. Hour UZT: %s", DAILY_REPORTS, DAILY_REPORT_HOUR_UZT)
    while True:
        try:
            now = datetime.now(UZT)
            today_key = now.strftime("%Y-%m-%d")
            if DAILY_REPORTS and now.hour == DAILY_REPORT_HOUR_UZT:
                for row in _connected_users_basic():
                    telegram_id = int(row["telegram_id"])
                    key = (telegram_id, today_key)
                    if key in _daily_report_sent:
                        continue
                    try:
                        token = cipher.decrypt(row["uzum_token_encrypted"])
                        client = UzumClient(token, UZUM_API_BASE_URL)
                        text = await _build_morning_report_text(telegram_id, client)
                        await bot.send_message(telegram_id, text, reply_markup=main_menu_for_user(telegram_id))
                        _daily_report_sent.add(key)
                        await asyncio.sleep(0.5)
                    except Exception:
                        logging.exception("Daily report: failed to send for user=%s", telegram_id)
            # Чистим старые ключи раз в сутки, чтобы память не росла.
            if now.hour == 0:
                _daily_report_sent.intersection_update({k for k in _daily_report_sent if k[1] == today_key})
        except Exception:
            logging.exception("Daily report loop error")
        await asyncio.sleep(1800)


async def subscription_reminder_loop() -> None:
    await asyncio.sleep(120)
    logging.info("Subscription reminder loop started. Enabled: %s", SUBSCRIPTION_REMINDERS)
    while True:
        try:
            if SUBSCRIPTION_REMINDERS:
                now = _utc_now()
                today_key = now.strftime("%Y-%m-%d")
                with db.connect() as conn:
                    rows = conn.execute(
                        """
                        SELECT telegram_id, trial_until, subscription_until, blocked
                        FROM subscriptions
                        WHERE blocked = 0
                        """
                    ).fetchall()
                for row in rows:
                    telegram_id = int(row["telegram_id"])
                    if is_admin(telegram_id):
                        continue
                    until = subscription_active_until(dict(row))
                    if not until:
                        continue
                    delta = until - now
                    if timedelta(0) < delta <= timedelta(days=SUBSCRIPTION_REMINDER_DAYS):
                        key = (telegram_id, today_key)
                        if key in _subscription_reminder_sent:
                            continue
                        try:
                            await bot.send_message(
                                telegram_id,
                                "⏳ <b>Скоро закончится доступ</b>\n\n"
                                f"Подписка/trial активны до: <b>{_fmt_dt(until)}</b>\n\n"
                                "Чтобы бот продолжил работать без остановки, продлите подписку заранее.\n"
                                "Оплата: <code>/subscribe</code>",
                                reply_markup=main_menu_for_user(telegram_id),
                            )
                            _subscription_reminder_sent.add(key)
                            await asyncio.sleep(0.2)
                        except Exception:
                            logging.exception("Subscription reminder: failed for user=%s", telegram_id)
        except Exception:
            logging.exception("Subscription reminder loop error")
        await asyncio.sleep(3600)


# --- OPTIMIZED WATCHERS: защита от 429 Too Many Requests ---
# Если один и тот же Uzum API-токен / магазин подключён у нескольких Telegram-пользователей
# (например, владелец и жена), старые watcher-функции делали одинаковый запрос для каждого пользователя.
# Ниже мы переопределяем check_*_once: один запрос на связку token+shop, потом рассылка всем пользователям группы.
def connected_watch_groups() -> list[dict[str, Any]]:
    groups: dict[str, dict[str, Any]] = {}
    for row in connected_users_for_order_watch():
        telegram_id = int(row["telegram_id"])
        shop_id = int(row["default_shop_id"])
        encrypted_token = row["uzum_token_encrypted"]
        key = hashlib.sha1(f"{shop_id}:{encrypted_token}".encode("utf-8")).hexdigest()
        if key not in groups:
            groups[key] = {
                "shop_id": shop_id,
                "uzum_token_encrypted": encrypted_token,
                "telegram_ids": [],
            }
        if telegram_id not in groups[key]["telegram_ids"]:
            groups[key]["telegram_ids"].append(telegram_id)
    return list(groups.values())


async def check_new_orders_once() -> None:
    for group in connected_watch_groups():
        shop_id = int(group["shop_id"])
        encrypted_token = group["uzum_token_encrypted"]
        telegram_ids = [int(x) for x in group["telegram_ids"]]

        try:
            token = cipher.decrypt(encrypted_token)
            client = UzumClient(token, UZUM_API_BASE_URL)
            data = await client.get_fbs_orders(shop_id, status="CREATED", page=0, size=20)
            items = extract_items(data)
        except Exception:
            logging.exception("Order watcher optimized: failed to check shop=%s users=%s", shop_id, telegram_ids)
            await asyncio.sleep(3)
            continue

        keys_now = [order_key(item) for item in items]
        for telegram_id in telegram_ids:
            known = _seen_order_keys_by_user.setdefault(telegram_id, set())
            if telegram_id not in _orders_watch_initialized:
                known.update(keys_now)
                _orders_watch_initialized.add(telegram_id)
                logging.info(
                    "Order watcher initialized for user=%s shop=%s orders=%s",
                    telegram_id, shop_id, len(keys_now),
                )
                continue

            new_items = [item for item, key in zip(items, keys_now) if key not in known]
            known.update(keys_now)
            if len(known) > 1000:
                _seen_order_keys_by_user[telegram_id] = set(keys_now)
            if not new_items:
                continue

            lines = [format_order_line(item) for item in new_items[:5]]
            more = "" if len(new_items) <= 5 else f"\n\nЕщё новых заказов: {len(new_items) - 5}"
            text = (
                f"🔔 <b>Новый заказ CREATED</b>\n"
                f"Магазин: <code>{shop_id}</code>\n"
                f"Новых заказов: <b>{len(new_items)}</b>\n\n"
                + "\n\n".join(lines)
                + more
                + "\n\nОткрыть список: <code>/orders</code>"
            )
            try:
                await bot.send_message(telegram_id, text, reply_markup=main_menu_for_user(telegram_id))
                await asyncio.sleep(0.15)
            except Exception:
                logging.exception("Order watcher: failed to send notification to %s", telegram_id)
        await asyncio.sleep(0.5)


async def check_low_stock_once() -> None:
    threshold = max(0, LOW_STOCK_THRESHOLD)
    for group in connected_watch_groups():
        shop_id = int(group["shop_id"])
        encrypted_token = group["uzum_token_encrypted"]
        telegram_ids = [int(x) for x in group["telegram_ids"]]

        try:
            token = cipher.decrypt(encrypted_token)
            client = UzumClient(token, UZUM_API_BASE_URL)
            rows = await load_sku_rows(client, shop_id, max_pages=50)
        except Exception:
            logging.exception("Low stock watcher optimized: failed to check shop=%s users=%s", shop_id, telegram_ids)
            await asyncio.sleep(3)
            continue

        low_rows = [r for r in rows if r.get("total") is not None and int(r.get("total") or 0) <= threshold]
        low_keys_now = [stock_row_key(r) for r in low_rows]

        for telegram_id in telegram_ids:
            known = _seen_low_stock_keys_by_user.setdefault(telegram_id, set())
            if telegram_id not in _low_stock_watch_initialized:
                known.update(low_keys_now)
                _low_stock_watch_initialized.add(telegram_id)
                logging.info(
                    "Low stock watcher initialized for user=%s shop=%s low_skus=%s threshold=%s",
                    telegram_id, shop_id, len(low_keys_now), threshold,
                )
                continue

            new_low_rows = [r for r, key in zip(low_rows, low_keys_now) if key not in known]
            _seen_low_stock_keys_by_user[telegram_id] = set(low_keys_now)
            if not new_low_rows:
                continue

            lines = [format_sku_stock_line(item, mode="all") for item in new_low_rows[:10]]
            more = "" if len(new_low_rows) <= 10 else f"\n\nЕщё SKU с низким остатком: {len(new_low_rows) - 10}"
            text = (
                f"📉 <b>Товар заканчивается</b>\n"
                f"Магазин: <code>{shop_id}</code>\n"
                f"Порог: ≤ <b>{threshold}</b> шт.\n"
                f"Новых позиций с низким остатком: <b>{len(new_low_rows)}</b>\n\n"
                + "\n\n".join(lines)
                + more
                + f"\n\nПоказать все низкие остатки: <code>/lowstock {threshold}</code>"
            )
            try:
                await bot.send_message(telegram_id, text, reply_markup=main_menu_for_user(telegram_id))
                await asyncio.sleep(0.15)
            except Exception:
                logging.exception("Low stock watcher: failed to send notification to %s", telegram_id)
        await asyncio.sleep(0.5)


async def check_out_of_stock_once() -> None:
    for group in connected_watch_groups():
        shop_id = int(group["shop_id"])
        encrypted_token = group["uzum_token_encrypted"]
        telegram_ids = [int(x) for x in group["telegram_ids"]]

        try:
            token = cipher.decrypt(encrypted_token)
            client = UzumClient(token, UZUM_API_BASE_URL)
            rows = await load_sku_rows(client, shop_id, max_pages=50)
        except Exception:
            logging.exception("Out of stock watcher optimized: failed to check shop=%s users=%s", shop_id, telegram_ids)
            await asyncio.sleep(3)
            continue

        zero_rows = [r for r in rows if r.get("total") is not None and int(r.get("total") or 0) == 0]
        zero_keys_now = [stock_row_key(r) for r in zero_rows]

        for telegram_id in telegram_ids:
            known = _seen_out_of_stock_keys_by_user.setdefault(telegram_id, set())
            if telegram_id not in _out_of_stock_watch_initialized:
                known.update(zero_keys_now)
                _out_of_stock_watch_initialized.add(telegram_id)
                logging.info(
                    "Out of stock watcher initialized for user=%s shop=%s zero_skus=%s",
                    telegram_id, shop_id, len(zero_keys_now),
                )
                continue

            new_zero_rows = [r for r, key in zip(zero_rows, zero_keys_now) if key not in known]
            _seen_out_of_stock_keys_by_user[telegram_id] = set(zero_keys_now)
            if not new_zero_rows:
                continue

            lines = [format_sku_stock_line(item, mode="all") for item in new_zero_rows[:10]]
            more = "" if len(new_zero_rows) <= 10 else f"\n\nЕщё SKU с нулевым остатком: {len(new_zero_rows) - 10}"
            text = (
                "❌ <b>Товар закончился</b>\n"
                f"Магазин: <code>{shop_id}</code>\n"
                f"Новых позиций с остатком 0: <b>{len(new_zero_rows)}</b>\n\n"
                + "\n\n".join(lines)
                + more
                + "\n\nПоказать остатки: <code>/stock</code>"
            )
            try:
                await bot.send_message(telegram_id, text, reply_markup=main_menu_for_user(telegram_id))
                await asyncio.sleep(0.15)
            except Exception:
                logging.exception("Out of stock watcher: failed to send notification to %s", telegram_id)
        await asyncio.sleep(0.5)


async def check_new_sales_once() -> None:
    date_from, date_to = _today_range_ms()
    for group in connected_watch_groups():
        shop_id = int(group["shop_id"])
        encrypted_token = group["uzum_token_encrypted"]
        telegram_ids = [int(x) for x in group["telegram_ids"]]

        try:
            token = cipher.decrypt(encrypted_token)
            client = UzumClient(token, UZUM_API_BASE_URL)
            rows, first_response = await _load_finance_orders(
                client,
                shop_id,
                date_from_ms=date_from,
                date_to_ms=date_to,
                max_pages=5,
                page_size=100,
            )
        except Exception:
            logging.exception("Sales watcher optimized: failed to check shop=%s users=%s", shop_id, telegram_ids)
            await asyncio.sleep(3)
            continue

        keys_now = [sale_key(item) for item in rows]
        for telegram_id in telegram_ids:
            known = _seen_sale_keys_by_user.setdefault(telegram_id, set())
            if telegram_id not in _sales_watch_initialized:
                known.update(keys_now)
                _sales_watch_initialized.add(telegram_id)
                logging.info(
                    "Sales watcher initialized for user=%s shop=%s sales_rows=%s",
                    telegram_id, shop_id, len(keys_now),
                )
                continue

            new_rows = [item for item, key in zip(rows, keys_now) if key not in known]
            known.update(keys_now)
            if len(known) > 3000:
                _seen_sale_keys_by_user[telegram_id] = set(keys_now)
            if not new_rows:
                continue

            for item in new_rows[:10]:
                try:
                    await bot.send_message(telegram_id, build_new_sale_message(item, shop_id=shop_id, lang=get_user_language(telegram_id)), reply_markup=main_menu_for_user(telegram_id))
                    await asyncio.sleep(0.15)
                except Exception:
                    logging.exception("Sales watcher: failed to send sale notification to %s", telegram_id)

            if len(new_rows) > 10:
                try:
                    await bot.send_message(
                        telegram_id,
                        f"➕ Ещё новых строк продаж: <b>{len(new_rows) - 10}</b>\nПодробно: <code>/balance</code>",
                        reply_markup=main_menu_for_user(telegram_id),
                    )
                except Exception:
                    logging.exception("Sales watcher: failed to send summary notification to %s", telegram_id)
        await asyncio.sleep(0.5)


async def check_stock_change_once() -> None:
    for group in connected_watch_groups():
        shop_id = int(group["shop_id"])
        encrypted_token = group["uzum_token_encrypted"]
        telegram_ids = [int(x) for x in group["telegram_ids"]]

        try:
            token = cipher.decrypt(encrypted_token)
            client = UzumClient(token, UZUM_API_BASE_URL)
            rows = await load_sku_rows(client, shop_id, max_pages=20)
            snapshot_now = _stock_change_snapshot(rows)
        except Exception:
            logging.exception("Stock change watcher optimized: failed to check shop=%s users=%s", shop_id, telegram_ids)
            await asyncio.sleep(3)
            continue

        for telegram_id in telegram_ids:
            previous = _stock_snapshot_by_user.setdefault(telegram_id, {})
            if telegram_id not in _stock_change_watch_initialized:
                _stock_snapshot_by_user[telegram_id] = snapshot_now
                _stock_change_watch_initialized.add(telegram_id)
                logging.info(
                    "Stock change watcher initialized for user=%s shop=%s skus=%s",
                    telegram_id, shop_id, len(snapshot_now),
                )
                continue

            decreased: list[tuple[str, dict[str, Any], dict[str, Any]]] = []
            for key, after in snapshot_now.items():
                before = previous.get(key)
                if not before:
                    continue
                before_total = int(before.get("total") or 0)
                after_total = int(after.get("total") or 0)
                before_fbo = int(before.get("fbo") or 0)
                after_fbo = int(after.get("fbo") or 0)
                before_fbs = int(before.get("fbs") or 0)
                after_fbs = int(after.get("fbs") or 0)
                if after_total < before_total or after_fbo < before_fbo or after_fbs < before_fbs:
                    decreased.append((key, before, after))

            _stock_snapshot_by_user[telegram_id] = snapshot_now
            if not decreased:
                continue

            lines = [_format_stock_change_line(key, before, after) for key, before, after in decreased[:10]]
            more = "" if len(decreased) <= 10 else f"\n\nЕщё изменений: {len(decreased) - 10}"
            text = (
                "📦 <b>Изменение остатков</b>\n"
                f"Магазин: <code>{shop_id}</code>\n"
                "Уменьшился остаток по SKU. Это может быть продажа, резерв, списание или изменение склада.\n\n"
                + "\n\n".join(lines)
                + more
                + "\n\nПроверить остатки: <code>/stock</code>"
            )
            try:
                await bot.send_message(telegram_id, text, reply_markup=main_menu_for_user(telegram_id))
                await asyncio.sleep(0.15)
            except Exception:
                logging.exception("Stock change watcher: failed to send notification to %s", telegram_id)
        await asyncio.sleep(0.5)


# --- Уведомления об отменах через Finance API ---
# Отмена в Uzum чаще всего не появляется как новый заказ, а меняет статус уже существующей
# финансовой строки на CANCELED / PARTIALLY_CANCELED. Поэтому обычный watcher новых продаж
# может её не прислать. Этот блок отслеживает именно изменение статуса.
CANCEL_NOTIFICATIONS = os.getenv("CANCEL_NOTIFICATIONS", "1").strip().lower() in {"1", "true", "yes", "on", "да"}
_sale_status_by_user: dict[int, dict[str, str]] = {}


def finance_identity_key(item: dict[str, Any]) -> str:
    """Стабильный ключ строки продажи без количества/суммы/статуса.

    Нужен, чтобы увидеть изменение статуса PROCESSING -> CANCELED у той же продажи.
    """
    parts: list[str] = []
    order_id = _finance_order_id(item)
    sale_id = _finance_sale_id(item)
    sku = str(_deep_pick_value(item, ("skuId", "sku_id", "skuTitle", "skuName", "barcode")) or "-")
    title = _finance_title(item)
    for label, value in (("order", order_id), ("sale", sale_id), ("sku", sku), ("title", title)):
        if value not in (None, "", "-"):
            parts.append(f"{label}:{value}")
    if parts:
        return "|".join(parts)
    raw = json.dumps(item, ensure_ascii=False, sort_keys=True, default=str)
    return "hash:" + hashlib.sha1(raw.encode("utf-8")).hexdigest()


def build_cancel_message(item: dict[str, Any], shop_id: int | None = None, lang: str = "ru") -> str:
    lang = normalize_lang(lang)
    title = escape(_finance_title(item))
    sku = escape(_finance_sku_title(item))
    qty = _finance_qty(item)
    unit_price = _deep_pick_number(item, ("sellPrice", "soldPrice", "price", "skuPrice", "productPrice"))
    status = escape(_finance_status(item))
    order_id = escape(_finance_order_id(item))
    sale_id = escape(_finance_sale_id(item))
    date_text = escape(_format_finance_date(_finance_date_value(item)))

    if lang == "uz":
        shop_line = f"🏪 Do‘kon: <code>{shop_id}</code>\n" if shop_id is not None else ""
        price_line = f"💵 Summa: <b>{_format_money(float(unit_price or 0))}</b>\n" if unit_price is not None else ""
        return (
            "❌ <b>Buyurtma bekor qilindi</b>\n\n"
            + shop_line +
            f"📦 Tovar: <b>{title}</b>\n"
            f"🔖 SKU: <code>{sku}</code>\n"
            f"🔢 Soni: <b>{qty:g} dona</b>\n\n"
            + price_line +
            f"🆔 Buyurtma: <code>{order_id}</code>\n"
            f"📌 Status: <code>{status}</code>\n"
            f"🕒 Sana: {date_text}"
        )

    shop_line = f"🏪 Магазин: <code>{shop_id}</code>\n" if shop_id is not None else ""
    price_line = f"💵 Сумма: <b>{_format_money(float(unit_price or 0))}</b>\n" if unit_price is not None else ""
    return (
        "❌ <b>Отмена заказа</b>\n\n"
        + shop_line +
        f"📦 Товар: <b>{title}</b>\n"
        f"🔖 SKU: <code>{sku}</code>\n"
        f"🔢 Кол-во: <b>{qty:g} шт.</b>\n\n"
        + price_line +
        f"🆔 Заказ: <code>{order_id}</code>\n"
        f"📌 Статус: <code>{status}</code>\n"
        f"🕒 Дата: {date_text}"
    )


# Переопределяем watcher продаж: теперь он присылает и новые продажи, и новые отмены.
async def check_new_sales_once() -> None:
    date_from, date_to = _today_range_ms()
    for group in connected_watch_groups():
        shop_id = int(group["shop_id"])
        encrypted_token = group["uzum_token_encrypted"]
        telegram_ids = [int(x) for x in group["telegram_ids"]]

        try:
            token = cipher.decrypt(encrypted_token)
            client = UzumClient(token, UZUM_API_BASE_URL)
            rows, first_response = await _load_finance_orders(
                client,
                shop_id,
                date_from_ms=date_from,
                date_to_ms=date_to,
                max_pages=5,
                page_size=100,
            )
        except Exception:
            logging.exception("Sales watcher optimized: failed to check shop=%s users=%s", shop_id, telegram_ids)
            await asyncio.sleep(3)
            continue

        keys_now = [sale_key(item) for item in rows]
        identity_status_now = {finance_identity_key(item): _finance_status(item) for item in rows}

        for telegram_id in telegram_ids:
            known = _seen_sale_keys_by_user.setdefault(telegram_id, set())
            status_memory = _sale_status_by_user.setdefault(telegram_id, {})

            # Первый проход: запоминаем текущее состояние, чтобы не прислать старые продажи/отмены.
            if telegram_id not in _sales_watch_initialized:
                known.update(keys_now)
                status_memory.update(identity_status_now)
                _sales_watch_initialized.add(telegram_id)
                logging.info(
                    "Sales watcher initialized for user=%s shop=%s sales_rows=%s",
                    telegram_id, shop_id, len(keys_now),
                )
                continue

            cancel_rows: list[dict[str, Any]] = []
            if CANCEL_NOTIFICATIONS:
                for item in rows:
                    ident = finance_identity_key(item)
                    current_status = _finance_status(item)
                    previous_status = status_memory.get(ident)
                    # Уведомляем, если строка стала отменённой после предыдущей проверки.
                    if _is_cancelled_status(current_status) and not _is_cancelled_status(str(previous_status or "")):
                        cancel_rows.append(item)
                status_memory.update(identity_status_now)

            # Новые строки продаж. Отменённые строки не отправляем как "новая продажа".
            new_rows = [
                item for item, key in zip(rows, keys_now)
                if key not in known and not _is_cancelled_status(_finance_status(item))
            ]
            known.update(keys_now)
            if len(known) > 3000:
                _seen_sale_keys_by_user[telegram_id] = set(keys_now)

            for item in new_rows[:10]:
                try:
                    await bot.send_message(
                        telegram_id,
                        build_new_sale_message(item, shop_id=shop_id, lang=get_user_language(telegram_id)),
                        reply_markup=main_menu_for_user(telegram_id),
                    )
                    await asyncio.sleep(0.15)
                except Exception:
                    logging.exception("Sales watcher: failed to send sale notification to %s", telegram_id)

            for item in cancel_rows[:10]:
                try:
                    await bot.send_message(
                        telegram_id,
                        build_cancel_message(item, shop_id=shop_id, lang=get_user_language(telegram_id)),
                        reply_markup=main_menu_for_user(telegram_id),
                    )
                    await asyncio.sleep(0.15)
                except Exception:
                    logging.exception("Sales watcher: failed to send cancel notification to %s", telegram_id)

            total_extra = max(0, len(new_rows) - 10) + max(0, len(cancel_rows) - 10)
            if total_extra:
                try:
                    await bot.send_message(
                        telegram_id,
                        f"➕ Ещё новых событий: <b>{total_extra}</b>\nПодробно: <code>/balance</code>",
                        reply_markup=main_menu_for_user(telegram_id),
                    )
                except Exception:
                    logging.exception("Sales watcher: failed to send summary notification to %s", telegram_id)
        await asyncio.sleep(0.5)


@dp.message(Command("cancel_notify_status", "cancellations_notify_status"))
async def cancel_notify_status(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    req = await require_connection(message)
    if req is None:
        return
    _, _, shop_id = req
    lang = get_user_language(telegram_id)
    if normalize_lang(lang) == "uz":
        text = (
            "❌ <b>Bekor qilingan buyurtmalar xabarnomasi</b>\n\n"
            f"Holat: {'✅ yoqilgan' if CANCEL_NOTIFICATIONS else '❌ o‘chirilgan'}\n"
            f"Do‘kon: <code>{shop_id}</code>\n"
            f"Tekshiruv: har <b>{max(60, SALE_CHECK_INTERVAL_SECONDS)}</b> soniyada\n\n"
            "Bot yangi bekor qilingan buyurtmalarni Finance API orqali kuzatadi."
        )
    else:
        text = (
            "❌ <b>Уведомления об отменах</b>\n\n"
            f"Статус: {'✅ включены' if CANCEL_NOTIFICATIONS else '❌ выключены'}\n"
            f"Магазин: <code>{shop_id}</code>\n"
            f"Проверка каждые: <b>{max(60, SALE_CHECK_INTERVAL_SECONDS)}</b> сек.\n\n"
            "Бот отслеживает новые отмены через Finance API."
        )
    await message.answer(text, reply_markup=main_menu_for_user(telegram_id))


@dp.message(F.text)
async def friendly_auto_start(message: Message, state: FSMContext) -> None:
    """Показывает понятное стартовое меню, если новый клиент написал любое сообщение вместо /start."""
    telegram_id = upsert_from_message(message)
    lang = get_user_language(telegram_id)
    current_state = await state.get_state()

    # Если пользователь был в процессе подключения/импорта, не сбрасываем состояние без причины.
    if current_state:
        if lang == "uz":
            await message.answer(
                "Men sizni tushundim. Agar jarayonni bekor qilmoqchi bo‘lsangiz, <code>/cancel</code> yuboring.\n"
                "Asosiy menyu quyida 👇",
                reply_markup=menu_for_message(message),
            )
        else:
            await message.answer(
                "Я вас понял. Если хотите отменить текущий шаг, отправьте <code>/cancel</code>.\n"
                "Главное меню ниже 👇",
                reply_markup=menu_for_message(message),
            )
        return

    ensure_subscription(telegram_id)
    if lang == "uz":
        text = (
            "👋 <b>Uzum Seller Assistant</b>\n\n"
            "Botdan foydalanishni boshlash uchun pastdagi menyudan kerakli bo‘limni tanlang.\n\n"
            "🔌 Agar do‘kon hali ulanmagan bo‘lsa — <b>Ulash</b> tugmasini bosing.\n"
            "🎥 API ulash videosi ham menyuda bor."
        )
    else:
        text = (
            "👋 <b>Uzum Seller Assistant</b>\n\n"
            "Чтобы начать пользоваться ботом, выберите нужный раздел в меню ниже.\n\n"
            "🔌 Если магазин ещё не подключён — нажмите <b>Подключить</b>.\n"
            "🎥 Видеоинструкция по API тоже есть в меню."
        )
    await message.answer(text, reply_markup=menu_for_message(message))


# --- Финальная чистка узбекских сообщений для клиентов ---
# Убирает оставшиеся русские фразы в разделах уведомлений, подписки и помощи.
_FINAL_UZ_TRANSLATE_RUNTIME_TEXT_TO_UZ = translate_runtime_text_to_uz


def translate_runtime_text_to_uz(text: str) -> str:
    if not isinstance(text, str) or not text:
        return text
    text = _FINAL_UZ_TRANSLATE_RUNTIME_TEXT_TO_UZ(text)

    fixes = [
        # Уведомления о низких остатках
        ("📉 <b>Уведомления о низких остатках</b>", "📉 <b>Kam qoldiq xabarnomalari</b>"),
        ("Уведомления о низких остатках", "Kam qoldiq xabarnomalari"),
        ("📉 <b>Низкие остатки</b>", "📉 <b>Kam qoldiq</b>"),
        ("Низкие остатки", "Kam qoldiq"),
        ("Порог:", "Chegara:"),
        ("Chegara: ≤", "Chegara: ≤"),
        ("шт.", "dona"),
        ("Проверка каждые:", "Tekshiruv har"),
        ("Tekshiruv har: <b>", "Tekshiruv har <b>"),
        ("Tekshiruv har <b>1800</b> soniya", "Tekshiruv har <b>1800</b> soniyada"),
        ("Tekshiruv har <b>300</b> soniya", "Tekshiruv har <b>300</b> soniyada"),
        ("Tekshiruv har <b>60</b> soniya", "Tekshiruv har <b>60</b> soniyada"),
        ("Состояние: остатки уже запомнены", "Holat: qoldiqlar allaqachon eslab qolingan"),
        ("Состояние: нулевые остатки уже запомнены", "Holat: nol qoldiqlar allaqachon eslab qolingan"),
        ("Состояние: инициализация при следующей проверке", "Holat: keyingi tekshiruvda ishga tushadi"),
        ("Holat: остатки уже запомнены", "Holat: qoldiqlar allaqachon eslab qolingan"),
        ("Holat: нулевые остатки уже запомнены", "Holat: nol qoldiqlar allaqachon eslab qolingan"),
        ("Holat: инициализация при следующей проверке", "Holat: keyingi tekshiruvda ishga tushadi"),
        ("Бот уведомит, когда товар впервые опустится до порога или ниже.", "Tovar birinchi marta belgilangan chegaragacha yoki undan pastga tushganda bot xabar beradi."),
        ("Бот уведомит, когда товар впервые опустится до остатка <b>0</b>.", "Tovar qoldig‘i birinchi marta <b>0</b> bo‘lganda bot xabar beradi."),

        # Подписка / тарифы
        ("💎 <b>Uzum Seller Assistant obunasi</b>", "💎 <b>Uzum Seller Assistant obunasi</b>"),
        ("Nimalar kiradi:", "Nimalar kiradi:"),
        ("1 месяц", "1 oy"),
        ("3 месяца", "3 oy"),
        ("6 месяцев", "6 oy"),
        ("1 oy — 250 000 so‘m | 3 oy — 650 000 so‘m | 6 oy — 1 200 000 so‘m | Без ограничений по количеству магазинов", "1 oy — 250 000 so‘m | 3 oy — 650 000 so‘m | 6 oy — 1 200 000 so‘m | Do‘konlar soni cheklanmagan"),
        ("Без ограничений по количеству магазинов", "Do‘konlar soni cheklanmagan"),
        ("Без ограничений по количеству do‘konlar", "Do‘konlar soni cheklanmagan"),
        ("To‘lov uchun administratorga yozing:", "To‘lov uchun administratorga yozing:"),
        ("Нажмите кнопку ниже, напишите администратору и отправьте чек. После проверки доступ будет продлён.", "Quyidagi tugmani bosing, administratorga yozing va chekni yuboring. Tekshiruvdan keyin kirish uzaytiriladi."),
        ("Нажмите кнопку ниже, напишите администратору и отправьте чек.", "Quyidagi tugmani bosing, administratorga yozing va chekni yuboring."),
        ("После проверки доступ будет продлён.", "Tekshiruvdan keyin kirish uzaytiriladi."),
        ("Чек tekshirilgach, administrator kirishni uzaytiradi.", "Chek tekshirilgach, administrator kirishni uzaytiradi."),
        ("Подписка", "Obuna"),
        ("подписка", "obuna"),

        # Остатки / продажи / статусы
        ("Статус:", "Holat:"),
        ("Status:", "Holat:"),
        ("Holat: ✅ включены", "Holat: ✅ yoqilgan"),
        ("Holat: ❌ выключены", "Holat: ❌ o‘chirilgan"),
        ("включены", "yoqilgan"),
        ("выключены", "o‘chirilgan"),
        ("Магазин:", "Do‘kon:"),
        ("Do‘kon:", "Do‘kon:"),
        ("сум", "so‘m"),
        ("сек", "soniya"),

        # Команды и подсказки
        ("Нажмите кнопку ниже", "Quyidagi tugmani bosing"),
        ("напишите администратору", "administratorga yozing"),
        ("отправьте чек", "chekni yuboring"),
        ("Проверка", "Tekshiruv"),
        ("Состояние", "Holat"),
        ("Порог", "Chegara"),
    ]
    for old, new in fixes:
        text = text.replace(old, new)

    # Небольшая нормализация после замен
    text = text.replace("soniya.", "soniya")
    text = text.replace("soniyaiya", "soniya")
    text = text.replace("Holat::", "Holat:")
    text = text.replace("Chegara::", "Chegara:")
    text = text.replace("Tekshiruv har: har", "Tekshiruv har")
    text = text.replace("Tekshiruv har har", "Tekshiruv har")
    return text


# --- FULL CHECK: финальная узбекская чистка перед запуском ---
# Этот слой специально стоит самым последним: исправляет смешанные русско-узбекские
# фразы, которые появляются из старых русских отчётов после автоматической замены.
_FULL_CHECK_TRANSLATE_RUNTIME_TEXT_TO_UZ = translate_runtime_text_to_uz


def translate_runtime_text_to_uz(text: str) -> str:
    if not isinstance(text, str) or not text:
        return text
    text = _FULL_CHECK_TRANSLATE_RUNTIME_TEXT_TO_UZ(text)

    fixes = [
        # Частые артефакты от замен "день/дней" внутри слова "сегодня/средняя"
        ("сегоkun", "bugun"),
        ("Сегоkun", "Bugun"),
        ("сегодня", "bugun"),
        ("Сегодня", "Bugun"),
        ("Среkunя строка", "O‘rtacha savdo"),
        ("Средняя строка", "O‘rtacha savdo"),
        ("O‘rtacha строка", "O‘rtacha savdo"),
        ("Отменённых qatorlar", "Bekor qilinganlar"),
        ("Отмененных qatorlar", "Bekor qilinganlar"),
        ("Отменённых строк", "Bekor qilinganlar"),
        ("Отмененных строк", "Bekor qilinganlar"),

        # Подключение API
        ("🔌 <b>Подключение магазина</b>", "🔌 <b>Do‘konni ulash</b>"),
        ("Чтобы подключить магазин, создайте API-ключ в кабинете Uzum Seller и отправьте его сюда.", "Do‘konni ulash uchun Uzum Seller kabinetida API-kalit yarating va shu yerga yuboring."),
        ("🎥 Видеоинструкция", "🎥 Videoqo‘llanma"),
        ("📌 Текстовая инструкция", "📌 Matnli qo‘llanma"),
        ("Отправьте API-ключ следующим сообщением.", "API-kalitni keyingi xabarda yuboring."),
        ("Отмена:", "Bekor qilish:"),
        ("Похоже, это не Uzum API-токен.", "Bu Uzum API-kalitiga o‘xshamaydi."),
        ("Отправьте полный токен или нажмите", "To‘liq kalitni yuboring yoki bosing"),
        ("✅ <b>Магазин уже подключён</b>", "✅ <b>Do‘kon allaqachon ulangan</b>"),
        ("Если вы случайно нажали <b>🔌 Подключить</b>, ничего страшного — старый API-ключ не удалён и не слетит.", "Agar <b>🔌 Ulash</b> tugmasini tasodifan bosgan bo‘lsangiz, xavotir olmang — eski API-kalit o‘chirilmaydi."),
        ("Чтобы заменить API-ключ, используйте только команду", "API-kalitni almashtirish uchun faqat quyidagi buyruqdan foydalaning:"),
        ("Чтобы полностью удалить подключение:", "Ulanishni butunlay o‘chirish uchun:"),
        ("🔐 <b>Uzum API-ключ не принят</b>", "🔐 <b>Uzum API-kalit qabul qilinmadi</b>"),
        ("Возможно, ключ неверный, удалён или истёк.", "Kalit noto‘g‘ri, o‘chirilgan yoki muddati tugagan bo‘lishi mumkin."),
        ("Создайте новый ключ в кабинете Uzum Seller и подключите его через", "Uzum Seller kabinetida yangi kalit yarating va uni ulang:"),
        ("✅ Подключение к Uzum API удалено. Можно подключить заново через", "✅ Uzum API ulanishi o‘chirildi. Qayta ulash uchun:"),

        # Видео / инструкция / безопасность / поддержка
        ("🎥 <b>Видеоинструкция по подключению API</b>", "🎥 <b>API ulash bo‘yicha videoqo‘llanma</b>"),
        ("В видео коротко показано:", "Videoda qisqa ko‘rsatilgan:"),
        ("Где в кабинете Uzum Seller находятся ключи API.", "Uzum Seller kabinetida API kalitlari qayerda joylashgani."),
        ("Как создать новый ключ.", "Yangi kalitni qanday yaratish."),
        ("Как подключить ключ к боту через", "Kalitni botga qanday ulash:"),
        ("Нажмите кнопку ниже, чтобы открыть видео", "Videoni ochish uchun quyidagi tugmani bosing"),
        ("▶️ Смотреть видео", "▶️ Videoni ko‘rish"),
        ("🔑 <b>Как подключить Uzum API к боту</b>", "🔑 <b>Uzum API-kalitni botga qanday ulash mumkin</b>"),
        ("API-ключ создаётся только в вашем кабинете Uzum Seller.", "API-kalit faqat sizning Uzum Seller kabinetingizda yaratiladi."),
        ("Это не пароль от кабинета, ключ можно удалить в любой момент.", "Bu kabinet paroli emas, kalitni istalgan vaqtda o‘chirishingiz mumkin."),
        ("Где взять API-ключ:", "API-kalitni qayerdan olish mumkin:"),
        ("Нажмите на профиль / аватарку в правом верхнем углу.", "O‘ng yuqori burchakdagi profil / avatarkani bosing."),
        ("Откройте", "Oching:"),
        ("Ключи API", "API kalitlari"),
        ("Создать ключ", "Kalit yaratish"),
        ("Скопируйте ключ.", "Kalitni nusxalang."),
        ("Вернитесь в бот и отправьте ключ через", "Botga qayting va kalitni yuboring:"),
        ("🔐 <b>Безопасность API-ключа</b>", "🔐 <b>API-kalit xavfsizligi</b>"),
        ("Ваш Uzum API-ключ не показывается в боте и не отправляется обратно сообщением.", "Uzum API-kalitingiz botda ko‘rsatilmaydi va qayta xabar qilib yuborilmaydi."),
        ("После подключения бот старается удалить сообщение, где был отправлен ключ.", "Ulangandan keyin bot kalit yuborilgan xabarni o‘chirishga harakat qiladi."),
        ("В базе хранится только защищённая версия ключа.", "Bazaga kalitning faqat himoyalangan ko‘rinishi saqlanadi."),
        ("Вы можете в любой момент удалить подключение командой", "Ulanishni istalgan vaqtda quyidagi buyruq bilan o‘chirishingiz mumkin:"),
        ("🛟 <b>Поддержка Uzum Seller Assistant</b>", "🛟 <b>Uzum Seller Assistant yordami</b>"),
        ("Ваш Telegram ID:", "Telegram ID’ingiz:"),
        ("Если бот не показывает данные, проверьте:", "Agar bot ma’lumot ko‘rsatmasa, tekshiring:"),
        ("API-ключ активен в кабинете Uzum Seller.", "API-kalit Uzum Seller kabinetida faol."),
        ("У ключа есть доступ к нужному магазину.", "Kalit kerakli do‘konga kirish huquqiga ega."),
        ("В кабинете Uzum есть продажи за выбранный период.", "Uzum kabinetida tanlangan davr uchun savdolar bor."),
        ("Если меняли API-ключ — нажмите", "Agar API-kalitni almashtirgan bo‘lsangiz, bosing:"),

        # Подписка
        ("💎 <b>Подписка Uzum Seller Assistant</b>", "💎 <b>Uzum Seller Assistant obunasi</b>"),
        ("Что входит:", "Nimalar kiradi:"),
        ("продажи FBO/FBS за сегодня, вчера, 7 и 30 дней", "bugun, kecha, 7 va 30 kunlik FBO/FBS savdolar"),
        ("остатки и товары, которые заканчиваются", "qoldiqlar va tugab borayotgan tovarlar"),
        ("потерянные товары, если Uzum отдаёт их в API", "Uzum API bersa, yo‘qolgan tovarlar"),
        ("уведомления о новых продажах", "yangi savdolar haqida xabarnomalar"),
        ("работа с несколькими магазинами", "bir nechta do‘kon bilan ishlash"),
        ("Excel-отчёт и утренний отчёт", "Excel hisobot va ertalabki hisobot"),
        ("для нового пользователя", "yangi foydalanuvchi uchun"),
        ("💰 <b>Тарифы</b>", "💰 <b>Tariflar</b>"),
        ("Для оплаты напишите администратору", "To‘lov uchun administratorga yozing"),
        ("После проверки чекa администратор продлит доступ.", "Chek tekshirilgach, administrator kirishni uzaytiradi."),
        ("Проверить статус", "Holatni tekshirish"),
        ("Trial до:", "Trial muddati:"),
        ("Оплачено до:", "To‘langan sana:"),
        ("Тарифы:", "Tariflar:"),
        ("История оплат:", "To‘lovlar tarixi:"),
        ("Поддержка:", "Yordam:"),
        ("Заменить API-ключ:", "API-kalitni almashtirish:"),
        ("Удалить API-ключ:", "API-kalitni o‘chirish:"),
        ("⛔ Obuna закончилась", "⛔ Obuna muddati tugagan"),
        ("⛔ Подписка закончилась", "⛔ Obuna muddati tugagan"),
        ("👑 Админ-доступ: без ограничений", "👑 Admin kirish: cheklovsiz"),
        ("⛔ Пользователь заблокирован", "⛔ Foydalanuvchi bloklangan"),

        # Разделы / меню / общее
        ("Русский", "Rus tili"),
        ("Выберите действие", "Amalni tanlang"),
        ("Выберите раздел", "Bo‘limni tanlang"),
        ("Главное меню", "Asosiy menyu"),
        ("Действие отменено.", "Amal bekor qilindi."),
        ("Язык интерфейса", "Interfeys tili"),
        ("Выберите язык, на котором бот будет показывать меню и основные подсказки.", "Bot menyu va asosiy ko‘rsatmalarni qaysi tilda ko‘rsatishini tanlang."),
        ("Язык изменён", "Til o‘zgartirildi"),
        ("Админ-панель доступна только владельцу бота.", "Admin panel faqat bot egasi uchun."),
        ("Доступ ограничен", "Kirish cheklangan"),
        ("Trial или подписка закончились.", "Trial yoki obuna muddati tugagan."),
        ("Ваш Uzum-токен и настройки сохранены — после продления всё снова заработает.", "Uzum tokeningiz va sozlamalaringiz saqlanadi — obuna uzaytirilgach hammasi yana ishlaydi."),
        ("Проверить подписку", "Obunani tekshirish"),
        ("Оплата", "To‘lov"),
        ("Сначала подключите магазин", "Avval do‘konni ulang"),
        ("Сначала подключите Uzum API-токен", "Avval Uzum API-kalitini ulang"),
        ("Продажи", "Savdo"),
        ("Склад", "Ombor"),
        ("Уведомления", "Xabarnomalar"),
        ("Отчёты", "Hisobotlar"),
        ("Что проверить", "Tekshirish"),
        ("Помощь", "Yordam"),
        ("Подключить", "Ulash"),
        ("Магазины", "Do‘konlar"),
        ("Пользователь", "Foydalanuvchi"),
        ("Доступ", "Kirish"),
        ("Проверить подключение", "Ulanishni tekshirish"),
        ("Что делать", "Nima qilish kerak"),
        ("отправьте Uzum API-ключ", "Uzum API-kalitini yuboring"),

        # Продажи / финансы / отчёты
        ("💰 <b>Продажи за bugun</b>", "💰 <b>Bugungi savdo</b>"),
        ("💰 <b>Продажи за Bugun</b>", "💰 <b>Bugungi savdo</b>"),
        ("💰 <b>Продажи ", "💰 <b>Savdo "),
        ("Продаж:", "Savdolar:"),
        ("Товаров продано:", "Sotilgan tovarlar:"),
        ("Возвратов:", "Qaytarilganlar:"),
        ("Продаж не найдено", "Savdolar topilmadi"),
        ("Пока продаж tanlangan davr uchun не найдено.", "Tanlangan davr uchun hali savdolar topilmadi."),
        ("Если продажа только появилась в кабинете, она может отобразиться чуть позже.", "Agar savdo kabinetda endi paydo bo‘lgan bo‘lsa, botda biroz keyin ko‘rinishi mumkin."),
        ("Ответ Finance API пришёл, но строки продаж не найдены.", "Finance API javob berdi, lekin savdo qatorlari topilmadi."),
        ("Фрагмент ответа", "Javobdan parcha"),
        ("Показаны первые", "Birinchi"),
        ("позиций из", "pozitsiya ko‘rsatildi, jami"),
        ("Новая продажа", "Yangi savdo"),
        ("Заказ", "Buyurtma"),
        ("Цена:", "Narx:"),
        ("Чистая прибыль", "Sof foyda"),
        ("Топ товаров по прибыли", "Foyda bo‘yicha top tovarlar"),
        ("Top tovarlar по прибыли", "Foyda bo‘yicha top tovarlar"),
        ("Загрузить себестоимость", "Tannarxni yuklash"),
        ("Себестоимостей сохранено", "Saqlangan tannarxlar"),
        ("Себестоимость ещё не указана", "Tannarx hali ko‘rsatilmagan"),
        ("Добавить", "Qo‘shish"),
        ("Укажите себестоимость так", "Tannarxni shunday kiriting"),
        ("Например", "Masalan"),
        ("Теперь проверьте", "Endi tekshiring"),
        ("Юнит-экономику", "Unit iqtisodiyot"),
        ("Unit iqtisodiyot за", "Unit iqtisodiyot"),

        # Склад / товары / накладные
        ("📦 <b>Товары магазина</b>", "📦 <b>Do‘kon tovarlari</b>"),
        ("Tovarы магазина", "Do‘kon tovarlari"),
        ("📊 <b>Остатки FBO / склад Uzum</b>", "📊 <b>FBO qoldiqlari / Uzum ombori</b>"),
        ("Qoldiq FBO / склад Uzum", "FBO qoldiqlari / Uzum ombori"),
        ("Проверяю потерянные товары", "Yo‘qolgan tovarlar tekshirilmoqda"),
        ("Раздел использует поле quantityMissing из Products API. Если в кабинете Uzum потери считаются по актам иначе, сумма может отличаться.", "Bo‘lim Products API’dagi quantityMissing maydonidan foydalanadi. Agar Uzum kabinetida yo‘qotishlar boshqacha hisoblangan bo‘lsa, summa farq qilishi mumkin."),
        ("Загружаю FBO-накладные поставки", "FBO yuk xatlari yuklanmoqda"),
        ("Чтобы посмотреть состав, отправьте", "Tarkibini ko‘rish uchun yuboring"),
        ("Например:", "Masalan:"),
        ("Заказов по основным статусам не найдено", "Asosiy statuslar bo‘yicha buyurtmalar topilmadi"),
        ("Собираю общую сводку магазина", "Do‘kon bo‘yicha umumiy xulosa tayyorlanmoqda"),
        ("остаток изменился", "qoldiq o‘zgardi"),
        ("Показать товары с низким остатком", "Kam qoldiqdagi tovarlarni ko‘rsatish"),
        ("Готовлю подробный Excel-отчёт", "Batafsil Excel hisobot tayyorlanmoqda"),
        ("Это может занять 20–60 soniyaунд", "Bu 20–60 soniya vaqt olishi mumkin"),
        ("собираю продажи, остатки и FBO-накладные", "savdolar, qoldiqlar va FBO yuk xatlari yig‘ilmoqda"),

        # Уведомления / отмены
        ("❌ <b>Уведомления об отменах</b>", "❌ <b>Bekor qilishlar xabarnomalari</b>"),
        ("Бот отслеживает новые отмены через Finance API.", "Bot yangi bekor qilishlarni Finance API orqali kuzatadi."),
        ("🛒 <b>Новая продажа</b>", "🛒 <b>Yangi savdo</b>"),

        # Что требует внимания
        ("Критичных проблем не видно. Можно посмотреть продажи и прибыль 30 kun uchun.", "Jiddiy muammo ko‘rinmayapti. 30 kunlik savdo va foydani ko‘rishingiz mumkin."),
        ("Начните с товаров, которые закончились: они не смогут продаваться, пока не пополнятся остатки.", "Avval tugagan tovarlardan boshlang: qoldiq to‘ldirilmaguncha ular sotilmaydi."),
        ("Сначала проверьте товары, которые скоро закончатся, особенно если они хорошо продаются.", "Avval tez tugaydigan tovarlarni tekshiring, ayniqsa ular yaxshi sotilayotgan bo‘lsa."),
        ("Загрузите себестоимость через Excel, чтобы бот точнее считал прибыль и маржу.", "Bot foyda va marjani aniqroq hisoblashi uchun tannarxni Excel orqali yuklang."),
        ("Проверьте товары с низкой маржой: возможно, цена или себестоимость указаны невыгодно.", "Past marjali tovarlarni tekshiring: narx yoki tannarx foydasiz bo‘lishi mumkin."),
        ("Посмотрите товары без продаж: возможно, стоит изменить цену, фото или вывести товар из оборота.", "Sotilmayotgan tovarlarni ko‘ring: narxni, rasmlarni o‘zgartirish yoki tovarni chiqarish kerak bo‘lishi mumkin."),
        ("Проверьте сегодняшние отмены и товары, по которым они произошли.", "Bugungi bekor qilishlar va ular bo‘lgan tovarlarni tekshiring."),
        ("Ниже можете сразу открыть нужный раздел", "Quyida kerakli bo‘limni darhol ochishingiz mumkin"),
        ("Проверяю магазин", "Do‘kon tekshirilmoqda"),
        ("Скоро закончится", "Tez tugaydi"),
        ("Закончились", "Tugagan"),
        ("Без продаж", "Sotuv yo‘q"),
        ("Без себестоимости", "Tannarx yo‘q"),
        ("Низкая маржа", "Past marja"),
        ("Низкая прибыль", "Past foyda"),
        ("Отмены сегодня", "Bugungi bekor qilishlar"),
        ("Потерянные", "Yo‘qolganlar"),
        ("Рекомендация", "Tavsiya"),

        # Excel / выгрузки
        ("Сводка", "Xulosa"),
        ("Показатель", "Ko‘rsatkich"),
        ("Значение", "Qiymat"),
        ("Дата создания отчёта", "Hisobot yaratilgan sana"),
        ("Период продаж в деталях", "Savdolar davri batafsil"),
        ("SKU в остатках", "Qoldiqdagi SKU"),
        ("SKU заканчиваются", "Tugayotgan SKU"),
        ("SKU с потерями", "Yo‘qotishli SKU"),
        ("FBO накладных найдено", "FBO yuk xatlari topildi"),
        ("Состав накладных загружен", "Yuk xatlari tarkibi yuklandi"),
        ("Период", "Davr"),
        ("Позиций", "Pozitsiyalar"),
        ("Товаров, dona", "Tovarlar, dona"),
        ("ID заказа/операции", "Buyurtma/operatsiya ID"),
        ("SKU/код", "SKU/kod"),
        ("Выведено", "Chiqarilgan"),
        ("Сырой фрагмент", "Xom parcha"),
        ("Код продавца", "Sotuvchi kodi"),
        ("Категория", "Kategoriya"),
        ("Цена", "Narx"),
        ("Активно", "Faol"),
        ("Потеряно", "Yo‘qolgan"),
        ("Брак", "Yaroqsiz"),
        ("Ожидает", "Kutilmoqda"),
        ("Примерная so‘mма", "Taxminiy summa"),
        ("Номер", "Raqam"),
        ("Создана", "Yaratilgan"),
        ("Создан", "Yaratilgan"),
        ("Окно от", "Oyna boshi"),
        ("Окно до", "Oyna oxiri"),
        ("Принята", "Qabul qilingan"),
        ("По накладной", "Yuk xati bo‘yicha"),
        ("Закупочная цена", "Xarid narxi"),
        ("Сумма по накладной", "Yuk xati summasi"),

        # Админка — чтобы в узбекском режиме тоже не было каши, но команды оставляем как есть
        ("👑 Админ", "👑 Admin"),
        ("👥 Пользователи", "👥 Foydalanuvchilar"),
        ("💳 Оплаты", "💳 To‘lovlar"),
        ("⏳ Скоро заканчиваются", "⏳ Tugayotganlar"),
        ("⛔ Заблокированные", "⛔ Bloklanganlar"),
        ("📦 Бэкап базы", "📦 Baza zaxirasi"),
        ("📢 Рассылка", "📢 Xabar yuborish"),
        ("⬅️ Главное меню", "⬅️ Asosiy menyu"),
        ("Чтобы отправить сообщение всем пользователям, напишите", "Barcha foydalanuvchilarga xabar yuborish uchun yozing"),
        ("Пример", "Masalan"),
        ("Отправляю резервную копию базы. Храните файл аккуратно — там данные пользователей.", "Baza zaxirasi yuborilmoqda. Faylni ehtiyot saqlang — unda foydalanuvchilar ma’lumotlari bor."),
    ]

    for old, new in fixes:
        text = text.replace(old, new)

    # Yakuniy normalizatsiya: eng ko‘p uchraydigan aralashmalar
    import re
    text = re.sub(r"(?<=\d)\s*сум\b", " so‘m", text)
    text = re.sub(r"(?<=\d)\s*шт\.?\b", " dona", text)
    text = text.replace("so‘mма", "summa")
    text = text.replace("so‘mмы", "summaning")
    text = text.replace("Tovarов", "Tovarlar")
    text = text.replace("Qaytarilganов", "Qaytarilganlar")
    text = text.replace("Savdo 30 kun uchun", "30 kunlik savdo")
    text = text.replace("Savdo 7 kun uchun", "7 kunlik savdo")
    text = text.replace("Savdo bugun uchun", "Bugungi savdo")
    text = text.replace("Savdo kecha uchun", "Kechagi savdo")
    text = text.replace("за bugun", "bugun uchun")
    text = text.replace("за Kecha", "kecha uchun")
    text = text.replace("за 7 kun", "7 kun uchun")
    text = text.replace("за 30 kun", "30 kun uchun")
    text = text.replace(" | 1 oy", " | 1 oy")
    text = text.replace("soniyaунд", "soniya")
    text = text.replace("soniyaия", "soniya")
    text = text.replace("Holat::", "Holat:")
    text = text.replace("Chegara::", "Chegara:")
    text = text.replace("Tekshiruv har: ", "Tekshiruv har ")
    text = text.replace("Tekshiruv har <b>300</b> soniya", "Tekshiruv har <b>300</b> soniyada")
    text = text.replace("Tekshiruv har <b>1800</b> soniya", "Tekshiruv har <b>1800</b> soniyada")
    return text


# --- FINAL AUDIT LAYER: исправления после полной проверки примеров ---
_AUDIT_TRANSLATE_RUNTIME_TEXT_TO_UZ = translate_runtime_text_to_uz


def translate_runtime_text_to_uz(text: str) -> str:
    if not isinstance(text, str) or not text:
        return text
    text = _AUDIT_TRANSLATE_RUNTIME_TEXT_TO_UZ(text)
    fixes = [
        ("💎 <b>Obuna Uzum Seller Assistant</b>", "💎 <b>Uzum Seller Assistant obunasi</b>"),
        ("✅ продажи FBO/FBS bugun uchun, вчера, 7 и 30 kun", "✅ bugun, kecha, 7 va 30 kunlik FBO/FBS savdolar"),
        ("✅ продажи FBO/FBS bugun uchun, kecha, 7 va 30 kun", "✅ bugun, kecha, 7 va 30 kunlik FBO/FBS savdolar"),
        ("продажи FBO/FBS bugun uchun, вчера, 7 и 30 kun", "bugun, kecha, 7 va 30 kunlik FBO/FBS savdolar"),
        ("Для оплаты administratorga yozing", "To‘lov uchun administratorga yozing"),
        ("Для оплаты", "To‘lov uchun"),
        ("Tovarlar продано", "Sotilgan tovarlar"),
        ("Tovarlar sotildi", "Sotilgan tovarlar"),
        ("Sotilgan tovarlar:", "Sotilgan tovarlar:"),
        ("✅ <b>Do‘kon уже подключён</b>", "✅ <b>Do‘kon allaqachon ulangan</b>"),
        ("Do‘kon уже подключён", "Do‘kon allaqachon ulangan"),
        ("Savdo bugun uchun", "Bugungi savdo"),
        ("Savdo kecha uchun", "Kechagi savdo"),
        ("Savdo 7 kun uchun", "7 kunlik savdo"),
        ("Savdo 30 kun uchun", "30 kunlik savdo"),
        ("Продажи bugun uchun", "Bugungi savdo"),
        ("Продажи kecha uchun", "Kechagi savdo"),
        ("Продажи 7 kun uchun", "7 kunlik savdo"),
        ("Продажи 30 kun uchun", "30 kunlik savdo"),
        ("Sotuvlar bugun uchun", "Bugungi savdo"),
        ("Sotuvlar kecha uchun", "Kechagi savdo"),
        ("Sotuvlar 7 kun uchun", "7 kunlik savdo"),
        ("Sotuvlar 30 kun uchun", "30 kunlik savdo"),
        ("вчера", "kecha"),
        ("7 и 30", "7 va 30"),
        ("и 30", "va 30"),
        ("уже", "allaqachon"),
        ("подключён", "ulangan"),
        ("подключен", "ulangan"),
        ("продано", "sotilgan"),
        ("продажи", "savdolar"),
        ("Продажи", "Savdolar"),
        ("товары", "tovarlar"),
        ("Товары", "Tovarlar"),
        ("остатки", "qoldiqlar"),
        ("Остатки", "Qoldiqlar"),
        ("которые заканчиваются", "tugab borayotgan"),
        ("уведомления", "xabarnomalar"),
        ("Уведомления", "Xabarnomalar"),
        ("новых", "yangi"),
        ("несколькими магазинами", "bir nechta do‘kon"),
        ("работа с", "ishlash:"),
        ("потерянные", "yo‘qolgan"),
        ("Потерянные", "Yo‘qolgan"),
        ("если Uzum отдаёт их в API", "agar Uzum API’da bersa"),
        ("для", "uchun"),
        ("нового пользователя", "yangi foydalanuvchi"),
    ]
    for old, new in fixes:
        text = text.replace(old, new)
    # Последняя нормализация заголовков и фраз после всех замен
    text = text.replace("💰 <b>Savdo bugun uchun</b>", "💰 <b>Bugungi savdo</b>")
    text = text.replace("💰 <b>Savdo kecha uchun</b>", "💰 <b>Kechagi savdo</b>")
    text = text.replace("💰 <b>Savdo 7 kun uchun</b>", "💰 <b>7 kunlik savdo</b>")
    text = text.replace("💰 <b>Savdo 30 kun uchun</b>", "💰 <b>30 kunlik savdo</b>")
    text = text.replace("💰 <b>Bugungi savdo uchun</b>", "💰 <b>Bugungi savdo</b>")
    text = text.replace("💰 <b>Kechagi savdo uchun</b>", "💰 <b>Kechagi savdo</b>")
    text = text.replace("savdolar FBO/FBS bugun uchun, kecha, 7 va 30 kun", "bugun, kecha, 7 va 30 kunlik FBO/FBS savdolar")
    text = text.replace("Tovarlar sotilgan", "Sotilgan tovarlar")
    return text

async def main() -> None:
    logging.info("INTUITIVE_ATTENTION_INTERFACE_LOADED: simple sections + attention report + full stock list")
    init_language_tables()
    await bot.delete_webhook(drop_pending_updates=True)
    if NEW_ORDER_NOTIFICATIONS:
        asyncio.create_task(order_watch_loop())
    if LOW_STOCK_NOTIFICATIONS:
        asyncio.create_task(low_stock_watch_loop())
    if OUT_OF_STOCK_NOTIFICATIONS:
        asyncio.create_task(out_of_stock_watch_loop())
    if SALE_NOTIFICATIONS:
        asyncio.create_task(sales_watch_loop())
    if STOCK_CHANGE_NOTIFICATIONS:
        asyncio.create_task(stock_change_watch_loop())
    if DAILY_REPORTS:
        asyncio.create_task(daily_report_loop())
    if SUBSCRIPTION_REMINDERS:
        asyncio.create_task(subscription_reminder_loop())
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())



