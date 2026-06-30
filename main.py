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
from aiogram.types import FSInputFile, InlineKeyboardButton, InlineKeyboardMarkup, KeyboardButton, Message, ReplyKeyboardMarkup
from dotenv import load_dotenv
from openpyxl import Workbook
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
REPORT_INVOICE_PRODUCT_LIMIT = int(os.getenv("REPORT_INVOICE_PRODUCT_LIMIT", "10") or "10")
SMART_LOW_STOCK_DAYS = int(os.getenv("SMART_LOW_STOCK_DAYS", "3") or "3")
TOP_PRODUCTS_DAYS = int(os.getenv("TOP_PRODUCTS_DAYS", "30") or "30")
DEAD_STOCK_DAYS = int(os.getenv("DEAD_STOCK_DAYS", "30") or "30")
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
        "⛔ <b>Доступ ограничен</b>\n\n"
        "Trial или подписка закончились.\n"
        "Ваш Uzum-токен и настройки сохранены — после продления всё снова заработает.\n\n"
        "Проверить подписку: <code>/my_subscription</code>\n"
        "Оплата: <code>/subscribe</code>",
        reply_markup=MAIN_MENU,
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


init_subscription_tables()
init_business_tables()

# Минималистичное меню в стиле Noorza Bot.
# Главное меню оставляем коротким: только самые важные разделы.
# Все старые команды остаются рабочими: /products, /stock, /orders, /reviews и т.д.
MAIN_MENU = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="💰 Баланс"), KeyboardButton(text="🌐 Все магазины")],
        [KeyboardButton(text="📊 Сегодня"), KeyboardButton(text="📆 Вчера")],
        [KeyboardButton(text="🗓 7 дней"), KeyboardButton(text="📅 30 дней")],
        [KeyboardButton(text="📦 Остатки"), KeyboardButton(text="⚠️ Прогноз остатков")],
        [KeyboardButton(text="🏆 Топ товаров"), KeyboardButton(text="🐢 Не продаётся")],
        [KeyboardButton(text="🏪 Магазины"), KeyboardButton(text="🧭 Потерянные")],
        [KeyboardButton(text="📄 Накладные FBO"), KeyboardButton(text="📊 Excel отчёт")],
        [KeyboardButton(text="🌙 Утренний отчёт"), KeyboardButton(text="💎 Подписка")],
        [KeyboardButton(text="ℹ️ Помощь")],
    ],
    resize_keyboard=True,
    input_field_placeholder="Выберите действие",
)

# Для совместимости: старые обработчики разделов могут ссылаться на эти переменные.
ANALYTICS_MENU = MAIN_MENU
PRODUCTS_MENU = MAIN_MENU
ORDERS_MENU = MAIN_MENU
NOTIFICATIONS_MENU = MAIN_MENU
SETTINGS_MENU = MAIN_MENU

class ConnectStates(StatesGroup):
    waiting_for_token = State()


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


async def require_connection(message: Message) -> tuple[int, UzumClient, int] | None:
    telegram_id = upsert_from_message(message)
    if not await require_active_subscription(message, telegram_id):
        return None
    client = get_uzum_for_user(telegram_id)
    shop_id = db.get_default_shop_id(telegram_id)

    if client is None:
        await message.answer(
            "Сначала подключите ваш Uzum Seller API-токен.\n\n"
            "Команда: <code>/connect</code>",
            reply_markup=MAIN_MENU,
        )
        return None

    if shop_id is None:
        await message.answer(
            "Токен подключён, но основной магазин не выбран.\n"
            "Напишите <code>/shops</code>, потом <code>/setshop SHOP_ID</code>.",
            reply_markup=MAIN_MENU,
        )
        return None

    return telegram_id, client, int(shop_id)


async def send_api_error(message: Message, error: Exception) -> None:
    text = escape(str(error))
    if len(text) > 3500:
        text = text[:3500] + "\n..."
    await message.answer(f"⚠️ Ошибка API:\n<code>{text}</code>", reply_markup=MAIN_MENU)


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
            reply_markup=MAIN_MENU,
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
                reply_markup=MAIN_MENU,
            )
            return

        encrypted = cipher.encrypt(token)
        default_shop_id = db.save_connection(telegram_id, encrypted, shops)

        try:
            await message.delete()
        except Exception:
            pass

        lines = [format_shop_line(shop) for shop in shops[:20]]
        await message.answer(
            "✅ Uzum API подключён.\n\n"
            "Найденные магазины:\n\n"
            + "\n".join(lines)
            + "\n\nОсновной магазин: "
            + (f"<code>{default_shop_id}</code>" if default_shop_id else "не выбран")
            + "\n\nПроверка остатков: <code>/stock</code> или <code>/lowstock</code>",
            reply_markup=MAIN_MENU,
        )
        if state:
            await state.clear()
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("start", "help"))
async def start(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    ensure_subscription(telegram_id)
    connected = "✅ подключён" if db.has_uzum_connection(telegram_id) else "❌ не подключён"
    sub_line = subscription_status_text(telegram_id)

    admin_part = ""
    if is_admin(telegram_id):
        admin_part = (
            "\n👑 <b>Админ-команды</b>\n"
            "• <code>/users</code> — список пользователей\n"
            "• <code>/extend ID 30</code> — продлить доступ\n"
            "• <code>/block ID</code> / <code>/unblock ID</code> — блокировка\n"
            "• <code>/paid ID сумма дни</code> — записать оплату и продлить\n"
            "• <code>/payments</code> — история оплат\n"
            "• <code>/backup_db</code> — резервная копия базы\n"
            "• <code>/broadcast текст</code> — рассылка\n"
        )

    await message.answer(
        "👋 <b>Добро пожаловать в Uzum Seller Assistant</b>\n\n"
        "Бот помогает селлерам Uzum быстро смотреть важные данные прямо в Telegram:\n"
        "📊 продажи за сегодня, вчера, 7 и 30 дней\n"
        "📦 остатки FBO/FBS/DBS\n"
        "⚠️ товары, которые заканчиваются\n"
        "🧭 потерянные товары по данным Uzum\n"
        "📄 FBO-накладные и состав поставки\n"
        "📊 подробный Excel-отчёт по продажам, остаткам и накладным\n"
        "🔔 уведомления о продажах и изменении остатков\n\n"
        f"Uzum API: {connected}\n"
        f"Доступ: {sub_line}\n\n"
        "🚀 <b>Как начать</b>\n"
        "1. Нажмите <code>/connect</code>\n"
        "2. Отправьте свой Uzum Seller OpenAPI token\n"
        "3. Нажмите <b>📊 Сегодня</b> или <b>📦 Остатки</b>\n\n"
        "Не знаете, где взять токен? Напишите <code>/api_token</code>.\n"
        "Подписка и оплата: <code>/subscribe</code>.\n"
        "Поддержка: <code>/support</code>.\n"
        "Заменить API-ключ: <code>/reconnect</code>. Удалить API-ключ: <code>/disconnect</code>.\n"
        "Подробный Excel-отчёт: <code>/report_excel</code>."
        + admin_part,
        reply_markup=MAIN_MENU,
    )


@dp.message(Command("menu"))
async def menu(message: Message) -> None:
    upsert_from_message(message)
    await message.answer("Выберите раздел 👇", reply_markup=MAIN_MENU)


@dp.message(Command("cancel"))
async def cancel(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer("Действие отменено.", reply_markup=MAIN_MENU)


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
        reply_markup=MAIN_MENU,
    )

@dp.message(Command("connect", "reconnect"))
async def connect(message: Message, state: FSMContext) -> None:
    token = parse_args(message.text or "")
    if token:
        await connect_token(message, token, state)
        return

    upsert_from_message(message)
    await state.set_state(ConnectStates.waiting_for_token)
    await message.answer(
        "🔑 <b>Подключение Uzum API</b>\n\n"
        "Отправьте ваш Uzum Seller OpenAPI token следующим сообщением.\n\n"
        "Где взять токен: <code>/api_token</code>\n\n"
        "Важно:\n"
        "• токен будет сохранён в зашифрованном виде;\n"
        "• после проверки я постараюсь удалить сообщение с токеном;\n"
        "• отменить: <code>/cancel</code>.",
        reply_markup=MAIN_MENU,
    )


@dp.message(ConnectStates.waiting_for_token, F.text)
async def connect_waiting_token(message: Message, state: FSMContext) -> None:
    await connect_token(message, message.text or "", state)


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
        reply_markup=MAIN_MENU,
    )


@dp.message(Command("pinguzum"))
async def ping_uzum(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    client = get_uzum_for_user(telegram_id)
    if client is None:
        await message.answer(
            "Сначала подключите Uzum API-токен: <code>/connect</code>",
            reply_markup=MAIN_MENU,
        )
        return
    try:
        data = await client.get_shops()
        shops = extract_items(data)
        await message.answer(f"✅ Uzum API отвечает. Найдено магазинов: {len(shops)}", reply_markup=MAIN_MENU)
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("shops"))
async def shops(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    client = get_uzum_for_user(telegram_id)
    if client is None:
        await message.answer("Сначала подключите Uzum API-токен: <code>/connect</code>", reply_markup=MAIN_MENU)
        return

    try:
        data = await client.get_shops()
        items = extract_items(data)
        if not items:
            await message.answer(
                "Ответ получен, но список магазинов не найден:\n<code>"
                + escape(compact_json_preview(data))
                + "</code>",
                reply_markup=MAIN_MENU,
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
            reply_markup=MAIN_MENU,
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
            reply_markup=MAIN_MENU,
        )
        return

    shop_id = int(arg)
    ok = db.set_default_shop_id(telegram_id, shop_id)
    if not ok:
        await message.answer("Этот магазин не найден среди подключённых. Сначала обновите список: <code>/shops</code>", reply_markup=MAIN_MENU)
        return

    await message.answer(f"✅ Основной магазин выбран: <code>{shop_id}</code>", reply_markup=MAIN_MENU)


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
                reply_markup=MAIN_MENU,
            )
            return

        title = f"📦 <b>Товары магазина</b> <code>{shop_id}</code>"
        if search_query:
            title += f" по запросу “{escape(search_query)}”"
        lines = [format_product_line(item) for item in items[:10]]
        await message.answer(title + ":\n\n" + "\n\n".join(lines), reply_markup=MAIN_MENU)
    except Exception as e:
        await send_api_error(message, e)


async def send_stock_list(message: Message, mode: str = "all") -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    search_query = parse_args(message.text or "")

    try:
        rows = await load_sku_rows(client, shop_id, search_query=search_query, max_pages=10)
        if not rows:
            await message.answer("SKU-остатки не найдены.", reply_markup=MAIN_MENU)
            return

        if mode == "fbo":
            rows = [r for r in rows if (r.get("fbo") or 0) > 0]
            title = "📊 <b>Остатки FBO / склад Uzum</b>"
        elif mode == "fbs":
            rows = [r for r in rows if (r.get("fbs") or 0) > 0]
            title = "📊 <b>Остатки FBS/DBS / склад продавца</b>"
        else:
            title = "📊 <b>Остатки по SKU: FBO + FBS/DBS + итого</b>"

        if search_query:
            title += f"\nПоиск: {escape(search_query)}"
        if not rows:
            await message.answer(title + "\n\nНичего не найдено.", reply_markup=MAIN_MENU)
            return

        lines = [format_sku_stock_line(row, mode=mode) for row in rows[:25]]
        await message.answer(
            title + f"\nПоказано: {min(len(rows), 25)} из {len(rows)}\n\n" + "\n\n".join(lines),
            reply_markup=MAIN_MENU,
        )
    except Exception as e:
        await send_api_error(message, e)


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
    arg = parse_args(message.text or "")
    threshold = int(arg) if arg.isdigit() else 5

    try:
        rows = await load_sku_rows(client, shop_id, max_pages=20)
        if not rows:
            await message.answer("SKU-остатки не найдены.", reply_markup=MAIN_MENU)
            return

        low = [r for r in rows if r.get("total") is not None and r["total"] <= threshold]
        if not low:
            await message.answer(f"✅ В первых {len(rows)} SKU нет общего остатка ≤ {threshold}.", reply_markup=MAIN_MENU)
            return

        lines = [format_sku_stock_line(row, mode="all") for row in low[:30]]
        await message.answer(
            f"⚠️ <b>Низкие остатки по общему количеству ≤ {threshold}:</b>\n"
            f"Показано: {min(len(low), 30)} из {len(low)}\n\n"
            + "\n\n".join(lines),
            reply_markup=MAIN_MENU,
        )
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
            await message.answer(f"Заказы со статусом <code>{escape(status)}</code> не найдены.", reply_markup=MAIN_MENU)
            return

        lines = [format_order_line(item) for item in items[:10]]
        await message.answer(
            f"🛒 <b>Заказы {escape(status)} для магазина</b> <code>{shop_id}</code>:\n\n"
            + "\n".join(lines),
            reply_markup=MAIN_MENU,
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

    statuses = stats.get("statuses") or {}
    if statuses:
        lines.append("")
        lines.append("<b>Статусы:</b>")
        for status, count in sorted(statuses.items(), key=lambda x: str(x[0]))[:8]:
            lines.append(f"• <code>{escape(str(status))}</code>: {count}")

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
    statuses = stats.get("statuses") or {}
    status_lines = []
    for key in ("PROCESSING", "TO_WITHDRAW"):
        status_lines.append(f"{key}: <b>{int(statuses.get(key, 0))}</b>")
    for k, v in sorted(statuses.items()):
        if k not in {"PROCESSING", "TO_WITHDRAW"}:
            status_lines.append(f"{escape(str(k))}: <b>{int(v)}</b>")
        if len(status_lines) >= 7:
            break
    extra = ""
    if not rows:
        extra = (
            "\n\n<i>Finance API пока не вернул строки продаж за сегодня. "
            "Если в кабинете продажа уже есть, она может появиться здесь позже.</i>"
        )
    return (
        "💰 <b>Продажи Uzum FBO/FBS за сегодня</b>\n"
        f"Магазин: <code>{shop_id}</code>\n\n"
        f"<b>Позиции продаж:</b> {int(stats['rows'])}\n"
        f"<b>Кол-во товаров:</b> {float(stats['units']):.0f} шт.\n"
        f"<b>Возвраты:</b> {float(stats['returns']):.0f} шт.\n\n"
        f"<b>Выручка:</b> {_format_money(float(stats['revenue']))}\n"
        f"<b>Комиссия Uzum:</b> {_format_money(float(stats['commission']))}\n"
        f"<b>Логистика:</b> {_format_money(float(stats['logistics']))}\n\n"
        f"<b>К выплате всего:</b> {_format_money(float(stats['payout_total']))}\n"
        f"<b>Уже выведено:</b> {_format_money(float(stats['withdrawn']))}\n"
        f"<b>Остаток к выплате:</b> {_format_money(float(stats['left_to_withdraw']))}\n\n"
        "<b>Статусы:</b>\n" + "\n".join(status_lines) +
        "\n\n<i>Расчёт по данным /v1/finance/orders. Если в кабинете Uzum есть корректировки/расходы, итог может отличаться.</i>" + extra
    )


@dp.message(Command("today"))
async def today_sales(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⌛ Считаю продажи за сегодня...", reply_markup=MAIN_MENU)
    try:
        rows, _, _ = await _load_today_finance_flexible(client, shop_id)
        stats = _build_noorza_today_stats(rows)
        await message.answer(_format_noorza_today(shop_id, stats, rows), reply_markup=MAIN_MENU)
    except Exception as e:
        await send_api_error(message, e)


def _format_noorza_period(title: str, shop_id: int, stats: dict[str, Any], rows: list[dict[str, Any]]) -> str:
    text = _format_noorza_today(shop_id, stats, rows)
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
    await message.answer("⌛ Считаю продажи за вчера...", reply_markup=MAIN_MENU)
    try:
        date_from, date_to = _yesterday_range_ms()
        rows, _, _ = await _load_finance_range_flexible(client, shop_id, date_from, date_to)
        stats = _build_noorza_today_stats(rows)
        await message.answer(_format_noorza_period("Продажи Uzum FBO/FBS за вчера", shop_id, stats, rows), reply_markup=MAIN_MENU)
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("week"))
@dp.message(Command("last7"))
async def week_sales(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⌛ Считаю продажи за 7 дней...", reply_markup=MAIN_MENU)
    try:
        date_from, date_to = _last_7_days_range_ms()
        rows, _, _ = await _load_finance_range_flexible(client, shop_id, date_from, date_to)
        stats = _build_noorza_today_stats(rows)
        await message.answer(_format_noorza_period("Продажи Uzum FBO/FBS за 7 дней", shop_id, stats, rows), reply_markup=MAIN_MENU)
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("balance"))
async def balance(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⌛ Считаю баланс за 30 дней...", reply_markup=MAIN_MENU)
    try:
        date_from, date_to = _days_range_ms(30)
        rows, _, _ = await _load_finance_range_flexible(client, shop_id, date_from, date_to)
        stats = _build_noorza_today_stats(rows)
        await message.answer(
            _format_noorza_period("Баланс Uzum FBO за 30 дней", shop_id, stats, rows),
            reply_markup=MAIN_MENU,
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

    await message.answer("⌛ Проверяю потерянные товары...", reply_markup=MAIN_MENU)
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

        title = "🧭 <b>Потерянные товары Uzum FBO</b>"
        if not products:
            await message.answer(
                title
                + "\n\nПотерянных товаров не найдено.\n"
                + "Проверил поле <code>quantityMissing</code> в списке товаров Uzum.",
                reply_markup=MAIN_MENU,
            )
            return

        total_missing = sum(_product_missing_qty(p) for p in products)
        approx_value = sum(
            _product_missing_qty(p)
            * (_pick_number(p, ("price", "sellPrice", "purchasePrice", "oldPrice")) or 0)
            for p in products
        )

        lines = [
            title,
            f"Магазин: <code>{shop_id}</code>",
            f"SKU с потерями: <b>{len(products)}</b>",
            f"Всего потеряно: <b>{total_missing} шт.</b>",
            f"Примерная сумма по текущей цене: <b>{_format_money(approx_value)}</b>",
        ]

        for idx, product in enumerate(products[:80], start=1):
            lines.append(_format_lost_product_line(product, idx))

        if len(products) > 80:
            lines.append(f"Показаны первые 80 SKU из {len(products)}.")

        lines.append("<i>Раздел использует поле quantityMissing из Products API. Если в кабинете Uzum потери считаются по актам иначе, сумма может отличаться.</i>")

        for part in _split_long_message("\n\n".join(lines)):
            await message.answer(part, reply_markup=MAIN_MENU)
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

    await message.answer("⌛ Загружаю FBO-накладные поставки...", reply_markup=MAIN_MENU)
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
            await message.answer(text, reply_markup=MAIN_MENU)
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
            await message.answer(part, reply_markup=MAIN_MENU)
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
            reply_markup=MAIN_MENU,
        )
        return
    invoice_id = int(arg.split()[0])

    await message.answer(f"⌛ Загружаю состав накладной <code>{invoice_id}</code>...", reply_markup=MAIN_MENU)
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
                reply_markup=MAIN_MENU,
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
            await message.answer(part, reply_markup=MAIN_MENU)
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("sales"))
async def sales(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⏳ Считаю продажи за сегодня, 7 и 30 дней...", reply_markup=MAIN_MENU)
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
            reply_markup=MAIN_MENU,
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
            _format_sales_details(days, shop_id, stats, first), reply_markup=MAIN_MENU
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
    await message.answer("⏳ Считаю заказы по статусам...", reply_markup=MAIN_MENU)
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
            reply_markup=MAIN_MENU,
        )
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("dashboard", "summary", "report"))
async def dashboard(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⏳ Собираю общую сводку магазина...", reply_markup=MAIN_MENU)
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
            reply_markup=MAIN_MENU,
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
            reply_markup=MAIN_MENU,
        )
        return

    if not items:
        await message.answer("⭐ Отзывы не найдены.", reply_markup=MAIN_MENU)
        return

    lines = [format_review_line(item) for item in items[:10]]
    await message.answer(
        "⭐ <b>Последние отзывы</b>\n\n"
        + "\n\n".join(lines)
        + "\n\nЧтобы ответить: <code>/reply ID_ОТЗЫВА ваш ответ</code>",
        reply_markup=MAIN_MENU,
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
            reply_markup=MAIN_MENU,
        )
        return

    review_id, answer_text = arg.split(maxsplit=1)
    review_id = review_id.strip()
    answer_text = answer_text.strip()
    if not review_id or not answer_text:
        await message.answer("Не вижу ID отзыва или текст ответа.", reply_markup=MAIN_MENU)
        return
    if len(answer_text) > 1000:
        await message.answer("Ответ слишком длинный. Сделайте до 1000 символов.", reply_markup=MAIN_MENU)
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
                    reply_markup=MAIN_MENU,
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
        reply_markup=MAIN_MENU,
    )


@dp.message(Command("subscribe"))
async def subscribe(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    ensure_subscription(telegram_id)
    await message.answer(
        "💎 <b>Подписка Uzum Seller Assistant</b>\n\n"
        "Что входит:\n"
        "✅ продажи FBO/FBS за сегодня, вчера, 7 и 30 дней\n"
        "✅ остатки и товары, которые заканчиваются\n"
        "✅ потерянные товары, если Uzum отдаёт их в API\n"
        "✅ уведомления о новых продажах и остатках\n\n"
        f"🎁 Trial: <b>{TRIAL_DAYS} дня</b> для нового пользователя\n\n"
        "💰 <b>Тарифы</b>\n"
        f"{escape(SUBSCRIPTION_PLANS_TEXT)}\n\n"
        f"Для оплаты напишите администратору: <b>{admin_contact_text()}</b>\n"
        f"{escape(PAYMENT_TEXT)}\n\n"
        "После проверки чекa администратор продлит доступ.\n"
        "Проверить статус: <code>/my_subscription</code>",
        reply_markup=admin_contact_markup(),
    )


@dp.message(Command("api_token", "token_help", "how_token"))
async def api_token_help(message: Message) -> None:
    upsert_from_message(message)
    await message.answer(
        "🔑 <b>Где взять Uzum Seller API-ключ</b>\n\n"
        "Инструкция:\n"
        "1. Зайдите в кабинет продавца <b>Uzum Seller</b>.\n"
        "2. Нажмите на свой профиль / аватарку в правом верхнем углу.\n"
        "3. Откройте раздел <b>Мой профиль</b>.\n"
        "4. Нажмите <b>Ключи API</b>.\n"
        "5. Нажмите <b>Создать ключ</b>.\n"
        "6. Скопируйте созданный API-ключ.\n"
        "7. Вернитесь в этот бот и нажмите <code>/connect</code>.\n"
        "8. Отправьте API-ключ одним сообщением.\n\n"
        "⚠️ <b>Важно:</b> не отправляйте API-ключ посторонним. "
        "Бот сохранит его в зашифрованном виде и постарается удалить сообщение с ключом после проверки.",
        reply_markup=MAIN_MENU,
    )




@dp.message(Command("security", "privacy"))
async def security(message: Message) -> None:
    upsert_from_message(message)
    await message.answer(
        "🔐 <b>Безопасность API-ключа</b>\n\n"
        "Ваш Uzum API-ключ не показывается в боте и не отправляется обратно сообщением.\n"
        "После подключения бот старается удалить сообщение, где был отправлен ключ.\n"
        "В базе хранится только защищённая версия ключа.\n\n"
        "Вы можете в любой момент удалить подключение командой <code>/disconnect</code>.\n"
        "Чтобы заменить ключ, используйте <code>/reconnect</code>.",
        reply_markup=MAIN_MENU,
    )


@dp.message(Command("support"))
async def support(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    user = db.get_user(telegram_id)
    shops = db.list_shops(telegram_id)
    connected = bool(user and user["uzum_token_encrypted"])
    await message.answer(
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
        f"Связаться с администратором: <b>{admin_contact_text()}</b>",
        reply_markup=admin_contact_markup() or MAIN_MENU,
    )


@dp.message(Command("my_payments"))
async def my_payments(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    rows = list_payments(telegram_id, 10)
    if not rows:
        await message.answer("💳 История оплат пока пустая.", reply_markup=MAIN_MENU)
        return
    await message.answer(
        "💳 <b>Мои оплаты</b>\n\n" + "\n".join(payment_line(row) for row in rows),
        reply_markup=MAIN_MENU,
    )

@dp.message(Command("my_subscription", "subscription"))
async def my_subscription(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    await message.answer(subscription_full_text(telegram_id), reply_markup=MAIN_MENU)


@dp.message(Command("users"))
async def admin_users(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    if not admin_only(telegram_id):
        return
    rows = list_subscription_users(30)
    if not rows:
        await message.answer("Пользователей пока нет.", reply_markup=MAIN_MENU)
        return
    lines = [subscription_compact_line(row) for row in rows]
    await message.answer(
        "👥 <b>Пользователи</b>\n\n"
        + "\n".join(lines)
        + "\n\nКоманды: <code>/extend ID 30</code>, <code>/paid ID сумма дни</code>, <code>/payments</code>",
        reply_markup=MAIN_MENU,
    )


@dp.message(Command("user"))
async def admin_user_info(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "")
    if not arg.split() or not arg.split()[0].isdigit():
        await message.answer("Напишите так: <code>/user TELEGRAM_ID</code>", reply_markup=MAIN_MENU)
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
        reply_markup=MAIN_MENU,
    )


@dp.message(Command("extend"))
async def admin_extend(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    parts = parse_args(message.text or "").split()
    if len(parts) < 2 or not parts[0].isdigit() or not parts[1].isdigit():
        await message.answer("Напишите так: <code>/extend TELEGRAM_ID 30</code>", reply_markup=MAIN_MENU)
        return
    target = int(parts[0])
    days = int(parts[1])
    new_until = extend_subscription_days(target, days)
    await message.answer(
        f"✅ Доступ продлён для <code>{target}</code> на {days} дней.\n"
        f"Активен до: <b>{_fmt_dt(new_until)}</b>",
        reply_markup=MAIN_MENU,
    )
    try:
        await bot.send_message(
            target,
            f"✅ Ваша подписка продлена на {days} дней.\nАктивна до: <b>{_fmt_dt(new_until)}</b>",
            reply_markup=MAIN_MENU,
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
            reply_markup=MAIN_MENU,
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
        reply_markup=MAIN_MENU,
    )
    try:
        await bot.send_message(
            target,
            "✅ <b>Оплата подтверждена</b>\n\n"
            f"Подписка продлена на <b>{days}</b> дней.\n"
            f"Доступ активен до: <b>{_fmt_dt(new_until)}</b>\n\n"
            "Спасибо за оплату!",
            reply_markup=MAIN_MENU,
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
        await message.answer("Напишите так: <code>/paid1 TELEGRAM_ID</code>", reply_markup=MAIN_MENU)
        return
    target = int(arg[0])
    new_until = extend_subscription_days(target, 30)
    payment_id = record_payment(target, 250000, 30, admin_id, "1 месяц")
    await message.answer(f"✅ Оплата #{payment_id}: 250 000 сум. Доступ для <code>{target}</code> до <b>{_fmt_dt(new_until)}</b>", reply_markup=MAIN_MENU)
    try:
        await bot.send_message(target, f"✅ Оплата подтверждена. Подписка продлена на 1 месяц. Доступ до: <b>{_fmt_dt(new_until)}</b>", reply_markup=MAIN_MENU)
    except Exception:
        pass


@dp.message(Command("paid3"))
async def admin_paid_3_months(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split(maxsplit=1)
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/paid3 TELEGRAM_ID</code>", reply_markup=MAIN_MENU)
        return
    target = int(arg[0])
    new_until = extend_subscription_days(target, 90)
    payment_id = record_payment(target, 650000, 90, admin_id, "3 месяца")
    await message.answer(f"✅ Оплата #{payment_id}: 650 000 сум. Доступ для <code>{target}</code> до <b>{_fmt_dt(new_until)}</b>", reply_markup=MAIN_MENU)
    try:
        await bot.send_message(target, f"✅ Оплата подтверждена. Подписка продлена на 3 месяца. Доступ до: <b>{_fmt_dt(new_until)}</b>", reply_markup=MAIN_MENU)
    except Exception:
        pass


@dp.message(Command("paid6"))
async def admin_paid_6_months(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split(maxsplit=1)
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/paid6 TELEGRAM_ID</code>", reply_markup=MAIN_MENU)
        return
    target = int(arg[0])
    new_until = extend_subscription_days(target, 180)
    payment_id = record_payment(target, 1200000, 180, admin_id, "6 месяцев")
    await message.answer(f"✅ Оплата #{payment_id}: 1 200 000 сум. Доступ для <code>{target}</code> до <b>{_fmt_dt(new_until)}</b>", reply_markup=MAIN_MENU)
    try:
        await bot.send_message(target, f"✅ Оплата подтверждена. Подписка продлена на 6 месяцев. Доступ до: <b>{_fmt_dt(new_until)}</b>", reply_markup=MAIN_MENU)
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
        await message.answer("💳 История оплат пока пустая.", reply_markup=MAIN_MENU)
        return
    title = f"💳 <b>Оплаты пользователя <code>{target}</code></b>" if target else "💳 <b>Последние оплаты</b>"
    await message.answer(title + "\n\n" + "\n".join(payment_line(row) for row in rows), reply_markup=MAIN_MENU)


@dp.message(Command("backup_db"))
async def admin_backup_db(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    path = Path(DB_PATH)
    if not path.exists():
        await message.answer(f"❌ База не найдена: <code>{escape(str(path))}</code>", reply_markup=MAIN_MENU)
        return
    await message.answer("📦 Отправляю резервную копию базы. Храните файл аккуратно — там данные пользователей.", reply_markup=MAIN_MENU)
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
        await message.answer("Напишите так: <code>/trial TELEGRAM_ID 3</code>", reply_markup=MAIN_MENU)
        return
    target = int(parts[0])
    days = int(parts[1])
    new_until = set_trial_days(target, days)
    await message.answer(f"🎁 Trial для <code>{target}</code> до <b>{_fmt_dt(new_until)}</b>", reply_markup=MAIN_MENU)


@dp.message(Command("block"))
async def admin_block(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split()
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/block TELEGRAM_ID</code>", reply_markup=MAIN_MENU)
        return
    target = int(arg[0])
    set_blocked(target, True)
    await message.answer(f"⛔ Пользователь <code>{target}</code> заблокирован.", reply_markup=MAIN_MENU)


@dp.message(Command("unblock"))
async def admin_unblock(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split()
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/unblock TELEGRAM_ID</code>", reply_markup=MAIN_MENU)
        return
    target = int(arg[0])
    set_blocked(target, False)
    await message.answer(f"✅ Пользователь <code>{target}</code> разблокирован.", reply_markup=MAIN_MENU)


@dp.message(Command("broadcast"))
async def admin_broadcast(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    text = parse_args(message.text or "")
    if not text:
        await message.answer("Напишите так: <code>/broadcast текст рассылки</code>", reply_markup=MAIN_MENU)
        return
    rows = list_subscription_users(500)
    sent = 0
    for row in rows:
        target = int(row["telegram_id"])
        try:
            await bot.send_message(target, "📢 <b>Сообщение от администратора</b>\n\n" + text, reply_markup=MAIN_MENU)
            sent += 1
            await asyncio.sleep(0.05)
        except Exception:
            pass
    await message.answer(f"✅ Рассылка завершена. Отправлено: {sent}", reply_markup=MAIN_MENU)


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
                reply_markup=MAIN_MENU,
            )
            return

        await message.answer(
            "🧪 <b>Первый товар — сырой JSON</b>\n\n<code>"
            + escape(compact_json_preview(items[0], limit=3200))
            + "</code>",
            reply_markup=MAIN_MENU,
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
            await message.answer("SKU-остатки для экспорта не найдены.", reply_markup=MAIN_MENU)
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
        await message.answer(f"✅ Экспортировано SKU-остатков: {len(rows)}", reply_markup=MAIN_MENU)
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
        reply_markup=MAIN_MENU,
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
            reply_markup=MAIN_MENU,
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
            await bot.send_message(telegram_id, text, reply_markup=MAIN_MENU)
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
        reply_markup=MAIN_MENU,
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
            await bot.send_message(telegram_id, text, reply_markup=MAIN_MENU)
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
        reply_markup=MAIN_MENU,
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
            await bot.send_message(telegram_id, text, reply_markup=MAIN_MENU)
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
        reply_markup=MAIN_MENU,
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


def build_new_sale_message(item: dict[str, Any], shop_id: int | None = None) -> str:
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

    shop_line = f"Магазин: <code>{shop_id}</code>\n" if shop_id is not None else ""
    return (
        "🛒 <b>Новая продажа Uzum FBO</b>\n\n"
        + shop_line +
        f"<b>Товар:</b> {title}\n"
        f"<b>SKU:</b> {sku}\n"
        f"<b>Кол-во:</b> {qty:g} шт.\n\n"
        f"<b>Цена продажи:</b> {_format_money(float(unit_price or 0))}\n"
        f"<b>Комиссия:</b> {_format_money(float(commission))}\n"
        f"<b>Логистика:</b> {_format_money(float(logistics))}\n"
        f"<b>К выплате:</b> {_format_money(float(payout))}\n\n"
        f"<b>ID заказа:</b> {escape(_finance_order_id(item))}\n"
        f"<b>ID продажи:</b> {escape(_finance_sale_id(item))}\n"
        f"<b>Статус:</b> {escape(_finance_status(item))}\n"
        f"<b>Дата:</b> {escape(_format_finance_date(_finance_date_value(item)))}"
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
                    build_new_sale_message(item, shop_id=shop_id),
                    reply_markup=MAIN_MENU,
                )
                await asyncio.sleep(0.15)
            except Exception:
                logging.exception("Sales watcher: failed to send sale notification to %s", telegram_id)

        if len(new_rows) > 10:
            try:
                await bot.send_message(
                    telegram_id,
                    f"➕ Ещё новых строк продаж: <b>{len(new_rows) - 10}</b>\nПодробно: <code>/balance</code>",
                    reply_markup=MAIN_MENU,
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
        reply_markup=MAIN_MENU,
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
            await bot.send_message(telegram_id, text, reply_markup=MAIN_MENU)
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
        reply_markup=MAIN_MENU,
    )


# --- Главное меню в стиле Noorza Bot ---
@dp.message(F.text == "⬅️ Главное меню")
@dp.message(F.text == "Меню")
async def button_main_menu(message: Message) -> None:
    await message.answer("Главное меню 👇", reply_markup=MAIN_MENU)


@dp.message(F.text == "💰 Баланс")
async def button_balance(message: Message) -> None:
    await balance(message)


@dp.message(F.text == "📊 Сегодня")
async def button_today(message: Message) -> None:
    await today_sales(message)


@dp.message(F.text == "📆 Вчера")
async def button_yesterday(message: Message) -> None:
    await yesterday_sales(message)


@dp.message(F.text == "🗓 7 дней")
async def button_week(message: Message) -> None:
    await week_sales(message)


@dp.message(F.text == "📅 30 дней")
async def button_30_days(message: Message) -> None:
    await sales_30(message)


@dp.message(F.text == "📦 Остатки")
async def button_stock_short(message: Message) -> None:
    await stock(message)


@dp.message(F.text == "⚠️ Заканчивается")
@dp.message(F.text == "⚠️ Заканчиваются")
async def button_lowstock_short(message: Message) -> None:
    await lowstock(message)


@dp.message(F.text == "🧭 Потерянные")
async def button_lost(message: Message) -> None:
    await lost_goods(message)


@dp.message(F.text == "📄 Накладные FBO")
async def button_fbo_invoices(message: Message) -> None:
    await fbo_invoices(message)


@dp.message(F.text == "💎 Подписка")
async def button_subscription(message: Message) -> None:
    await subscribe(message)


@dp.message(F.text == "ℹ️ Помощь")
@dp.message(F.text == "❓ Помощь")
async def button_help(message: Message) -> None:
    await start(message)


# Старые красивые кнопки оставлены для совместимости, если они остались у пользователя в Telegram.
@dp.message(F.text == "📊 Аналитика")
async def section_analytics(message: Message) -> None:
    await message.answer("Главное меню 👇", reply_markup=MAIN_MENU)


@dp.message(F.text == "📦 Товары")
async def section_products(message: Message) -> None:
    await stock(message)


@dp.message(F.text == "🛒 Заказы/продажи")
async def section_orders(message: Message) -> None:
    await orders(message)


@dp.message(F.text == "🔔 Уведомления")
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


@dp.message(F.text == "💰 Продажи")
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


@dp.message(F.text == "📊 Excel отчёт")
@dp.message(F.text == "📄 Excel-отчёт")
async def button_excel_report(message: Message) -> None:
    await report_excel(message)


@dp.message(F.text == "⚙️ Статус")
async def button_status(message: Message) -> None:
    await status(message)


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
    statuses = stats.get("statuses") or {}
    status_lines = []
    for key in ("PROCESSING", "TO_WITHDRAW"):
        status_lines.append(f"{key}: <b>{int(statuses.get(key, 0))}</b>")
    for k, v in sorted(statuses.items()):
        if k not in {"PROCESSING", "TO_WITHDRAW"}:
            status_lines.append(f"{escape(str(k))}: <b>{int(v)}</b>")
        if len(status_lines) >= 7:
            break
    text = (
        f"🌐 <b>Баланс по всем магазинам {escape(days_title)}</b>\n\n"
        f"Магазинов: <b>{shops_count}</b>\n"
        f"Позиции продаж: <b>{int(stats['rows'])}</b>\n"
        f"Кол-во товаров: <b>{float(stats['units']):.0f} шт.</b>\n"
        f"Возвраты: <b>{float(stats['returns']):.0f} шт.</b>\n\n"
        f"Выручка: <b>{_format_money(float(stats['revenue']))}</b>\n"
        f"Комиссия Uzum: <b>{_format_money(float(stats['commission']))}</b>\n"
        f"Логистика: <b>{_format_money(float(stats['logistics']))}</b>\n\n"
        f"К выплате всего: <b>{_format_money(float(stats['payout_total']))}</b>\n"
        f"Уже выведено: <b>{_format_money(float(stats['withdrawn']))}</b>\n"
        f"Остаток к выплате: <b>{_format_money(float(stats['left_to_withdraw']))}</b>\n\n"
        "<b>Статусы:</b>\n" + "\n".join(status_lines)
    )
    if per_shop:
        text += "\n\n<b>По магазинам:</b>\n" + "\n".join(per_shop[:20])
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
        await message.answer("Сначала подключите Uzum API-токен: <code>/connect</code>", reply_markup=MAIN_MENU)
        return
    await message.answer("⌛ Считаю баланс по всем магазинам за 30 дней...", reply_markup=MAIN_MENU)
    try:
        date_from, date_to = _days_range_ms(30)
        stats, per_shop, shops_count = await _all_shops_finance_stats(telegram_id, client, date_from, date_to)
        await message.answer(_format_all_shops_balance("за 30 дней", shops_count, stats, per_shop), reply_markup=MAIN_MENU)
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
    days = TOP_PRODUCTS_DAYS
    await message.answer(f"⌛ Считаю топ товаров за {days} дней...", reply_markup=MAIN_MENU)
    try:
        top, stats = await _top_products_for_shop(client, shop_id, days)
        if not top:
            await message.answer(f"🏆 <b>Топ товаров за {days} дней</b>\nМагазин: <code>{shop_id}</code>\n\nПродаж не найдено.", reply_markup=MAIN_MENU)
            return
        lines = [
            f"🏆 <b>Топ товаров за {days} дней</b>",
            f"Магазин: <code>{shop_id}</code>",
            f"Всего продано: <b>{float(stats['units']):.0f} шт.</b>",
            f"Выручка: <b>{_format_money(float(stats['revenue']))}</b>",
        ]
        for idx, item in enumerate(top[:20], start=1):
            title = escape(_short_text(item.get("title"), 85))
            sku = escape(_short_text(item.get("sku"), 60))
            sku_line = f"\nSKU: <code>{sku}</code>" if sku and sku != "-" else ""
            lines.append(
                f"{idx}. <b>{title}</b>{sku_line}\n"
                f"Продано: <b>{float(item.get('qty') or 0):.0f} шт.</b> | "
                f"Выручка: <b>{_format_money(float(item.get('revenue') or 0))}</b> | "
                f"К выплате: <b>{_format_money(float(item.get('payout') or 0))}</b>"
            )
        for part in _split_long_message("\n\n".join(lines)):
            await message.answer(part, reply_markup=MAIN_MENU)
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("deadstock", "no_sales", "stuck"))
async def dead_stock(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    days = DEAD_STOCK_DAYS
    await message.answer(f"⌛ Ищу товары с остатком, но без продаж за {days} дней...", reply_markup=MAIN_MENU)
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
            await message.answer(f"🐢 <b>Товары без продаж за {days} дней</b>\nМагазин: <code>{shop_id}</code>\n\nНе нашёл товаров с остатком и нулевыми продажами.", reply_markup=MAIN_MENU)
            return
        total_value = sum(float(x.get("value") or 0) for x in candidates)
        lines = [
            f"🐢 <b>Товары без продаж за {days} дней</b>",
            f"Магазин: <code>{shop_id}</code>",
            f"Позиций: <b>{len(candidates)}</b>",
            f"Примерно заморожено: <b>{_format_money(total_value)}</b>",
        ]
        for idx, item in enumerate(candidates[:30], start=1):
            row = item["row"]
            title = escape(_short_text(_stock_row_title(row), 85))
            sku = escape(_short_text(_stock_row_sku(row), 60))
            sku_line = f"\nSKU: <code>{sku}</code>" if sku else ""
            lines.append(
                f"{idx}. <b>{title}</b>{sku_line}\n"
                f"Остаток: <b>{int(item['total'])} шт.</b> | "
                f"Цена: {_format_money(float(item['price']))} | "
                f"Сумма: <b>{_format_money(float(item['value']))}</b>"
            )
        lines.append("<i>Расчёт примерный: бот сопоставляет продажи и остатки по SKU/названию.</i>")
        for part in _split_long_message("\n\n".join(lines)):
            await message.answer(part, reply_markup=MAIN_MENU)
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("smart_lowstock", "forecast_stock"))
async def smart_lowstock(message: Message) -> None:
    req = await require_connection(message)
    if req is None:
        return
    _, client, shop_id = req
    await message.answer("⌛ Считаю, на сколько дней хватит остатков...", reply_markup=MAIN_MENU)
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
            await message.answer(
                f"⚠️ <b>Умное 'Заканчивается'</b>\nМагазин: <code>{shop_id}</code>\n\n"
                f"Критичных товаров не нашёл. Порог: ≤ {LOW_STOCK_THRESHOLD} шт. или хватит меньше чем на {SMART_LOW_STOCK_DAYS} дня.",
                reply_markup=MAIN_MENU,
            )
            return
        lines = [
            "⚠️ <b>Умное 'Заканчивается'</b>",
            f"Магазин: <code>{shop_id}</code>",
            f"Порог: ≤ {LOW_STOCK_THRESHOLD} шт. или хватит меньше чем на {SMART_LOW_STOCK_DAYS} дня",
        ]
        for idx, item in enumerate(alerts[:30], start=1):
            row = item["row"]
            title = escape(_short_text(_stock_row_title(row), 85))
            sku = escape(_short_text(_stock_row_sku(row), 60))
            days_left = float(item["days_left"])
            days_text = "нет продаж за 7 дней" if days_left > 9000 else f"примерно на {days_left:.1f} дн."
            sku_line = f"\nSKU: <code>{sku}</code>" if sku else ""
            lines.append(
                f"{idx}. <b>{title}</b>{sku_line}\n"
                f"Остаток: <b>{int(item['total'])} шт.</b> | "
                f"Продажи за 7 дней: <b>{float(item['qty_7']):.0f} шт.</b> | "
                f"Хватит: <b>{escape(days_text)}</b>"
            )
        for part in _split_long_message("\n\n".join(lines)):
            await message.answer(part, reply_markup=MAIN_MENU)
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
        await message.answer("Сначала подключите Uzum API-токен: <code>/connect</code>", reply_markup=MAIN_MENU)
        return
    await message.answer("⌛ Готовлю утренний отчёт за вчера...", reply_markup=MAIN_MENU)
    try:
        await message.answer(await _build_morning_report_text(telegram_id, client), reply_markup=MAIN_MENU)
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("extend1", "extend_month"))
async def admin_extend_1_month(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split()
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/extend1 TELEGRAM_ID</code>", reply_markup=MAIN_MENU)
        return
    target = int(arg[0])
    new_until = extend_subscription_days(target, 30)
    await message.answer(f"✅ Продлено на 1 месяц для <code>{target}</code>. До: <b>{_fmt_dt(new_until)}</b>", reply_markup=MAIN_MENU)


@dp.message(Command("extend3"))
async def admin_extend_3_months(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split()
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/extend3 TELEGRAM_ID</code>", reply_markup=MAIN_MENU)
        return
    target = int(arg[0])
    new_until = extend_subscription_days(target, 90)
    await message.answer(f"✅ Продлено на 3 месяца для <code>{target}</code>. До: <b>{_fmt_dt(new_until)}</b>", reply_markup=MAIN_MENU)


@dp.message(Command("extend6"))
async def admin_extend_6_months(message: Message) -> None:
    admin_id = upsert_from_message(message)
    if not admin_only(admin_id):
        return
    arg = parse_args(message.text or "").split()
    if not arg or not arg[0].isdigit():
        await message.answer("Напишите так: <code>/extend6 TELEGRAM_ID</code>", reply_markup=MAIN_MENU)
        return
    target = int(arg[0])
    new_until = extend_subscription_days(target, 180)
    await message.answer(f"✅ Продлено на 6 месяцев для <code>{target}</code>. До: <b>{_fmt_dt(new_until)}</b>", reply_markup=MAIN_MENU)


@dp.message(F.text == "🌐 Все магазины")
async def button_all_shops(message: Message) -> None:
    await balance_all_shops(message)


@dp.message(F.text == "🏆 Топ товаров")
async def button_top_products(message: Message) -> None:
    await top_products(message)


@dp.message(F.text == "🐢 Не продаётся")
async def button_dead_stock(message: Message) -> None:
    await dead_stock(message)


@dp.message(F.text == "🌙 Утренний отчёт")
async def button_morning_report(message: Message) -> None:
    await morning_report(message)


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
                        await bot.send_message(telegram_id, text, reply_markup=MAIN_MENU)
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
                                reply_markup=MAIN_MENU,
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
                await bot.send_message(telegram_id, text, reply_markup=MAIN_MENU)
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
                await bot.send_message(telegram_id, text, reply_markup=MAIN_MENU)
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
                await bot.send_message(telegram_id, text, reply_markup=MAIN_MENU)
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
                    await bot.send_message(telegram_id, build_new_sale_message(item, shop_id=shop_id), reply_markup=MAIN_MENU)
                    await asyncio.sleep(0.15)
                except Exception:
                    logging.exception("Sales watcher: failed to send sale notification to %s", telegram_id)

            if len(new_rows) > 10:
                try:
                    await bot.send_message(
                        telegram_id,
                        f"➕ Ещё новых строк продаж: <b>{len(new_rows) - 10}</b>\nПодробно: <code>/balance</code>",
                        reply_markup=MAIN_MENU,
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
                await bot.send_message(telegram_id, text, reply_markup=MAIN_MENU)
                await asyncio.sleep(0.15)
            except Exception:
                logging.exception("Stock change watcher: failed to send notification to %s", telegram_id)
        await asyncio.sleep(0.5)

async def main() -> None:
    logging.info("FINANCE_SECONDS_PATCH_LOADED: dateFrom for Finance = seconds, dateTo = milliseconds")
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


