from __future__ import annotations

import asyncio
import logging
import os
import tempfile
from html import escape
from pathlib import Path
from typing import Any

from aiogram import Bot, Dispatcher, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import FSInputFile, Message, ReplyKeyboardMarkup, KeyboardButton
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

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
UZUM_API_BASE_URL = os.getenv("UZUM_API_BASE_URL", "https://api-seller.uzum.uz/api/seller-openapi").strip()
DB_PATH = os.getenv("DB_PATH", "bot.db").strip()
ENCRYPTION_KEY = os.getenv("ENCRYPTION_KEY", "").strip()

if not TELEGRAM_BOT_TOKEN:
    raise RuntimeError("TELEGRAM_BOT_TOKEN is empty. Set it in BotHost environment variables.")

db = Database(DB_PATH)
cipher = TokenCipher(ENCRYPTION_KEY)

bot = Bot(
    TELEGRAM_BOT_TOKEN,
    default=DefaultBotProperties(parse_mode=ParseMode.HTML),
)
dp = Dispatcher()

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
            "Команда: <code>/connect</code>"
        )
        return None

    if shop_id is None:
        await message.answer(
            "Токен подключён, но основной магазин не выбран.\n"
            "Напишите <code>/shops</code>, потом <code>/setshop SHOP_ID</code>."
        )
        return None

    return telegram_id, client, int(shop_id)


async def send_api_error(message: Message, error: Exception) -> None:
    text = escape(str(error))
    if len(text) > 3500:
        text = text[:3500] + "\n..."
    await message.answer(f"⚠️ Ошибка API:\n<code>{text}</code>")


def parse_args(text: str) -> str:
    parts = (text or "").split(maxsplit=1)
    return parts[1].strip() if len(parts) > 1 else ""


async def load_products(client: UzumClient, shop_id: int, *, search_query: str = "", max_pages: int = 20, page_size: int = 100) -> list[Any]:
    all_products: list[Any] = []
    for page in range(max_pages):
        data = await client.get_products(shop_id, search_query=search_query, page=page, size=page_size)
        items = extract_items(data)
        if not items:
            break
        all_products.extend(items)
        if len(items) < page_size:
            break
    return all_products


async def load_sku_rows(client: UzumClient, shop_id: int, *, search_query: str = "", max_pages: int = 20) -> list[dict[str, Any]]:
    products = await load_products(client, shop_id, search_query=search_query, max_pages=max_pages, page_size=100)
    return flatten_sku_rows(products)


async def connect_token(message: Message, token: str, state: FSMContext | None = None) -> None:
    telegram_id = upsert_from_message(message)
    token = token.strip()

    if not token or len(token) < 20:
        await message.answer("Похоже, это не Uzum API-токен. Отправьте полный токен или нажмите /cancel.")
        return

    try:
        client = UzumClient(token, UZUM_API_BASE_URL)
        data = await client.get_shops()
        shops = extract_items(data)

        if not shops:
            await message.answer(
                "Токен сработал, но список магазинов не найден.\n"
                "Ответ API:\n<pre>" + escape(compact_json_preview(data)) + "</pre>"
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
            + "\n\nПроверка остатков: <code>/stock</code> или <code>/lowstock</code>"
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
        "• <code>/connect</code> — подключить Uzum API-токен\n"
        "• <code>/disconnect</code> — удалить подключение\n"
        "• <code>/shops</code> — мои магазины\n"
        "• <code>/setshop SHOP_ID</code> — выбрать основной магазин\n"
        "• <code>/products [поиск]</code> — товары, цены и общий остаток\n"
        "• <code>/stock [поиск]</code> — FBO + FBS/DBS + итого по SKU\n"
        "• <code>/fbo [поиск]</code> — остатки FBO\n"
        "• <code>/fbs [поиск]</code> — остатки FBS/DBS\n"
        "• <code>/lowstock [порог]</code> — товары, которые заканчиваются по общему остатку\n"
        "• <code>/orders [CREATED]</code> — FBS/DBS заказы\n"
        "• <code>/export_products</code> — Excel: FBO, FBS/DBS, итого\n"
        "• <code>/debug_product</code> — сырой JSON первого товара\n"
        "• <code>/status</code> — статус подключения\n\n"
        "Для начала нажмите: <code>/connect</code>"
    )


@dp.message(Command("cancel"))
async def cancel(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer("Действие отменено.")


@dp.message(Command("status"))
async def status(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    user = db.get_user(telegram_id)
    shops = db.list_shops(telegram_id)
    connected = bool(user and user["uzum_token_encrypted"])
    default_shop_id = user["default_shop_id"] if user else None

    await message.answer(
        "📌 <b>Статус</b>\n\n"
        f"Uzum API: {'✅ подключён' if connected else '❌ не подключён'}\n"
        f"Магазинов: {len(shops)}\n"
        f"Основной магазин: {f'<code>{default_shop_id}</code>' if default_shop_id else 'не выбран'}\n\n"
        "Подписки и trial добавим следующим шагом."
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
        "Отправьте ваш <b>Uzum Seller OpenAPI token</b> следующим сообщением.\n\n"
        "Важно:\n"
        "• токен будет сохранён в зашифрованном виде;\n"
        "• после проверки я постараюсь удалить сообщение с токеном;\n"
        "• отменить: <code>/cancel</code>."
    )


@dp.message(ConnectStates.waiting_for_token, F.text)
async def connect_waiting_token(message: Message, state: FSMContext) -> None:
    await connect_token(message, message.text or "", state)


@dp.message(Command("disconnect"))
async def disconnect(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    db.disconnect_uzum(telegram_id)
    await message.answer("✅ Подключение к Uzum API удалено. Можно подключить заново через <code>/connect</code>.")


@dp.message(Command("pinguzum"))
async def ping_uzum(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    client = get_uzum_for_user(telegram_id)
    if client is None:
        await message.answer("Сначала подключите Uzum API-токен: <code>/connect</code>")
        return

    try:
        data = await client.get_shops()
        shops = extract_items(data)
        await message.answer(f"✅ Uzum API отвечает. Найдено магазинов: {len(shops)}")
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("shops"))
async def shops(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    client = get_uzum_for_user(telegram_id)
    if client is None:
        await message.answer("Сначала подключите Uzum API-токен: <code>/connect</code>")
        return

    try:
        data = await client.get_shops()
        items = extract_items(data)
        if not items:
            await message.answer("Ответ получен, но список магазинов не найден:\n<pre>" + escape(compact_json_preview(data)) + "</pre>")
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
            + "\n\nЧтобы выбрать: <code>/setshop SHOP_ID</code>"
        )
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("setshop"))
async def setshop(message: Message) -> None:
    telegram_id = upsert_from_message(message)
    arg = parse_args(message.text or "")

    if not arg.isdigit():
        await message.answer("Напишите так: <code>/setshop SHOP_ID</code>\nНапример: <code>/setshop 12345</code>")
        return

    shop_id = int(arg)
    ok = db.set_default_shop_id(telegram_id, shop_id)
    if not ok:
        await message.answer("Этот магазин не найден среди подключённых. Сначала обновите список: <code>/shops</code>")
        return

    await message.answer(f"✅ Основной магазин выбран: <code>{shop_id}</code>")


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
            await message.answer("Товары не найдены. Ответ API:\n<pre>" + escape(compact_json_preview(data)) + "</pre>")
            return

        title = f"📦 <b>Товары магазина <code>{shop_id}</code></b>"
        if search_query:
            title += f" по запросу “{escape(search_query)}”"
        lines = [format_product_line(item) for item in items[:10]]
        await message.answer(title + ":\n\n" + "\n\n".join(lines))
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
            await message.answer("SKU-остатки не найдены.")
            return

        if mode == "fbo":
            rows = [r for r in rows if (r.get("fbo") or 0) > 0]
            title = "📦 <b>Остатки FBO / склад Uzum</b>"
        elif mode == "fbs":
            rows = [r for r in rows if (r.get("fbs") or 0) > 0]
            title = "📦 <b>Остатки FBS/DBS / склад продавца</b>"
        else:
            title = "📦 <b>Остатки по SKU: FBO + FBS/DBS + итого</b>"

        if search_query:
            title += f"\nПоиск: <b>{escape(search_query)}</b>"

        if not rows:
            await message.answer(title + "\n\nНичего не найдено.")
            return

        lines = [format_sku_stock_line(row, mode=mode) for row in rows[:25]]
        await message.answer(title + f"\nПоказано: {min(len(rows), 25)} из {len(rows)}\n\n" + "\n\n".join(lines))
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
            await message.answer("SKU-остатки не найдены.")
            return

        low = [r for r in rows if r.get("total") is not None and r["total"] <= threshold]

        if not low:
            await message.answer(f"✅ В первых {len(rows)} SKU нет общего остатка ≤ {threshold}.")
            return

        lines = [format_sku_stock_line(row, mode="all") for row in low[:30]]
        await message.answer(
            f"⚠️ <b>Низкие остатки по общему количеству ≤ {threshold}:</b>\n"
            f"Показано: {min(len(low), 30)} из {len(low)}\n\n" + "\n\n".join(lines)
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
            await message.answer(f"Заказы со статусом <code>{escape(status)}</code> не найдены.")
            return

        lines = [format_order_line(item) for item in items[:10]]
        await message.answer(f"🧾 <b>Заказы {escape(status)} для магазина <code>{shop_id}</code>:</b>\n\n" + "\n".join(lines))
    except Exception as e:
        await send_api_error(message, e)


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
            await message.answer("Товар для debug не найден. Ответ API:\n<pre>" + escape(compact_json_preview(data, limit=3000)) + "</pre>")
            return

        await message.answer(
            "🔎 <b>Первый товар — сырой JSON</b>\n\n"
            "<pre>" + escape(compact_json_preview(items[0], limit=3200)) + "</pre>"
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
            await message.answer("SKU-остатки для экспорта не найдены.")
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

        await message.answer(f"✅ Экспортировано SKU-остатков: {len(rows)}")
        await message.answer_document(FSInputFile(filename))
    except Exception as e:
        await send_api_error(message, e)


async def main() -> None:
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
