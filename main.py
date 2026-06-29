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
from aiogram.types import FSInputFile, KeyboardButton, Message, ReplyKeyboardMarkup
from dotenv import load_dotenv
from openpyxl import Workbook

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
ORDER_CHECK_INTERVAL_SECONDS = int(os.getenv("ORDER_CHECK_INTERVAL_SECONDS", "300") or "300")
NEW_ORDER_NOTIFICATIONS = (
    os.getenv("NEW_ORDER_NOTIFICATIONS", "1").strip().lower()
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
    os.getenv("STOCK_CHANGE_NOTIFICATIONS", "1").strip().lower()
    not in {"0", "false", "no", "off"}
)
STOCK_CHANGE_CHECK_INTERVAL_SECONDS = int(
    os.getenv("STOCK_CHANGE_CHECK_INTERVAL_SECONDS", "300") or "300"
)

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

# Минималистичное меню в стиле Noorza Bot.
# Главное меню оставляем коротким: только самые важные разделы.
# Все старые команды остаются рабочими: /products, /stock, /orders, /reviews и т.д.
MAIN_MENU = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📊 Сегодня"), KeyboardButton(text="📆 Вчера")],
        [KeyboardButton(text="🗓 7 дней"), KeyboardButton(text="📅 30 дней")],
        [KeyboardButton(text="📦 Остатки"), KeyboardButton(text="⚠️ Заканчивается")],
        [KeyboardButton(text="🧭 Потерянные"), KeyboardButton(text="ℹ️ Помощь")],
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
    connected = "✅ подключён" if db.has_uzum_connection(telegram_id) else "❌ не подключён"
    await message.answer(
        "👋 <b>Uzum Seller Assistant</b>\n\n"
        f"Статус Uzum API: {connected}\n\n"
        "Основные команды:\n"
        "• <code>/menu</code> — показать кнопки\n"
        "• <code>/connect</code> — подключить Uzum API-токен\n"
        "• <code>/disconnect</code> — удалить подключение\n"
        "• <code>/shops</code> — мои магазины\n"
        "• <code>/setshop SHOP_ID</code> — выбрать основной магазин\n"
        "• <code>/products [поиск]</code> — товары, цены и общий остаток\n"
        "• <code>/stock [поиск]</code> — FBO + FBS/DBS + итого по SKU\n"
        "• <code>/fbo [поиск]</code> — остатки FBO\n"
        "• <code>/fbs [поиск]</code> — остатки FBS/DBS\n"
        "• <code>/lowstock [порог]</code> — товары, которые заканчиваются\n"
        "• <code>/orders [CREATED]</code> — FBS/DBS заказы\n"
        "• <code>/export_products</code> — Excel: FBO, FBS/DBS, итого\n"
        "• <code>/debug_product</code> — сырой JSON первого товара\n"
        "• <code>/status</code> — статус подключения\n"
        "• <code>/notify_status</code> — уведомления о новых FBS/DBS заказах\n"
        "• <code>/lowstock_notify_status</code> — низкие остатки FBO + FBS/DBS\n"
        "• <code>/outofstock_notify_status</code> — нулевые остатки FBO + FBS/DBS\n"
        "• <code>/stock_change_notify_status</code> — изменение остатков FBO + FBS/DBS\n\n"
        "Для начала нажмите: <code>/connect</code>",
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
    user = db.get_user(telegram_id)
    shops = db.list_shops(telegram_id)
    connected = bool(user and user["uzum_token_encrypted"])
    default_shop_id = user["default_shop_id"] if user else None
    await message.answer(
        "⚙️ <b>Статус</b>\n\n"
        f"Uzum API: {'✅ подключён' if connected else '❌ не подключён'}\n"
        f"Магазинов: {len(shops)}\n"
        f"Основной магазин: {f'<code>{default_shop_id}</code>' if default_shop_id else 'не выбран'}\n\n"
        "Подписки и trial добавим следующим шагом.",
        reply_markup=MAIN_MENU,
    )


@dp.message(Command("connect"))
async def connect(message: Message, state: FSMContext) -> None:
    token = parse_args(message.text or "")
    if token:
        await connect_token(message, token, state)
        return

    upsert_from_message(message)
    await state.set_state(ConnectStates.waiting_for_token)
    await message.answer(
        "Отправьте ваш Uzum Seller OpenAPI token следующим сообщением.\n\n"
        "Важно:\n"
        "• токен будет сохранён в зашифрованном виде;\n"
        "• после проверки я постараюсь удалить сообщение с токеном;\n"
        "• отменить: <code>/cancel</code>.",
        reply_markup=MAIN_MENU,
    )


@dp.message(ConnectStates.waiting_for_token, F.text)
async def connect_waiting_token(message: Message, state: FSMContext) -> None:
    await connect_token(message, message.text or "", state)


@dp.message(Command("disconnect"))
async def disconnect(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    db.disconnect_uzum(telegram_id)
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
    direct = _deep_pick_number(
        item,
        (
            "totalPrice", "totalAmount", "totalSum", "productPrice", "skuPrice",
            "sellPrice", "soldPrice", "priceWithDiscount", "orderItemPrice",
            "orderPrice", "purchasePrice", "price",
        ),
    )
    if direct is not None:
        return max(0.0, direct)
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
            "logistics", "logistic", "logisticAmount", "logisticsAmount",
            "logisticsSum", "delivery", "deliveryAmount", "deliveryPrice",
            "deliveryCost", "shipping", "shippingAmount",
        ),
    )
    return abs(value or 0.0)


def _finance_payout_direct(item: dict[str, Any]) -> float | None:
    return _deep_pick_number(
        item,
        (
            "amountToWithdraw", "toWithdraw", "withdrawAmount", "sellerPayout",
            "payout", "payoutAmount", "sellerAmount", "accrual", "accrualAmount",
        ),
    )


def _finance_withdrawn(item: dict[str, Any]) -> float:
    value = _deep_pick_number(
        item,
        (
            "withdrawn", "withdrawnAmount", "paid", "paidAmount", "transferred",
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
        statuses[status] = statuses.get(status, 0) + 1
        qty = _finance_qty(item)
        if _is_cancelled_status(status) or "RETURN" in status.upper() or "ВОЗВР" in status.upper():
            returns += abs(qty)
            continue
        active_rows += 1
        units += qty
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
    await message.answer("⌛ Считаю баланс по Finance API...", reply_markup=MAIN_MENU)
    try:
        rows, _, _ = await _load_today_finance_flexible(client, shop_id)
        stats = _build_noorza_today_stats(rows)
        await message.answer(
            "💰 <b>Баланс за сегодня</b>\n"
            f"Магазин: <code>{shop_id}</code>\n\n"
            f"К выплате всего: <b>{_format_money(float(stats['payout_total']))}</b>\n"
            f"Уже выведено: <b>{_format_money(float(stats['withdrawn']))}</b>\n"
            f"Остаток к выплате: <b>{_format_money(float(stats['left_to_withdraw']))}</b>\n\n"
            f"Выручка: {_format_money(float(stats['revenue']))}\n"
            f"Комиссия Uzum: {_format_money(float(stats['commission']))}\n"
            f"Логистика: {_format_money(float(stats['logistics']))}\n\n"
            "<i>Пока это расчёт по сегодняшним строкам Finance API.</i>",
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
    return [dict(row) for row in rows]


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


def format_sale_line(item: dict[str, Any]) -> str:
    title = escape(_finance_title(item))
    qty = _finance_qty(item)
    amount = _finance_revenue(item)
    status = escape(_finance_status(item))
    extra = []
    for k in ("orderId", "order_id", "skuId", "sku_id", "barcode"):
        v = item.get(k)
        if v not in (None, ""):
            extra.append(f"{k}: <code>{escape(str(v))}</code>")
            break
    extra_text = "\n" + " | ".join(extra) if extra else ""
    return (
        f"• <b>{title}</b>\n"
        f"  Штук: <b>{qty:g}</b> | Сумма: <b>{_format_money(amount)}</b> | Статус: <code>{status}</code>"
        f"{extra_text}"
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

        stats = _build_sales_stats(new_rows)
        lines = [format_sale_line(item) for item in new_rows[:8]]
        more = "" if len(new_rows) <= 8 else f"\n\nЕщё новых строк продаж: {len(new_rows) - 8}"
        text = (
            f"💸 <b>Новая продажа</b>\n"
            f"Магазин: <code>{shop_id}</code>\n"
            f"Новых строк: <b>{len(new_rows)}</b>\n"
            f"Сумма: <b>{_format_money(float(stats['revenue']))}</b>\n"
            f"Штук: <b>{float(stats['units']):g}</b>\n\n"
            + "\n\n".join(lines)
            + more
            + "\n\nПодробно: <code>/sales_today</code>"
        )

        try:
            await bot.send_message(telegram_id, text, reply_markup=MAIN_MENU)
        except Exception:
            logging.exception("Sales watcher: failed to send notification to %s", telegram_id)


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


@dp.message(F.text == "📄 Excel-отчёт")
async def button_export_products(message: Message) -> None:
    await export_products(message)


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
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())




