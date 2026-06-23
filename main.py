from __future__ import annotations

import asyncio
import logging
import os
import tempfile
from pathlib import Path
from typing import Any

from aiogram import Bot, Dispatcher
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command
from aiogram.types import FSInputFile, Message
from dotenv import load_dotenv
from openpyxl import Workbook

from formatters import (
    compact_json_preview,
    extract_items,
    format_order_line,
    format_product_line,
    format_shop_line,
    get_stock_number,
    pick,
)
from uzum_client import UzumApiError, UzumClient


load_dotenv()
logging.basicConfig(level=logging.INFO)

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
UZUM_API_TOKEN = os.getenv("UZUM_API_TOKEN", "").strip()
UZUM_API_BASE_URL = os.getenv("UZUM_API_BASE_URL", "https://api-seller.uzum.uz/api/seller-openapi").strip()
OWNER_TELEGRAM_ID = os.getenv("OWNER_TELEGRAM_ID", "").strip()
DEFAULT_SHOP_ID = os.getenv("DEFAULT_SHOP_ID", "").strip()

if not TELEGRAM_BOT_TOKEN:
    raise RuntimeError("TELEGRAM_BOT_TOKEN is empty. Set it in BotHost environment variables or .env file.")

bot = Bot(
    TELEGRAM_BOT_TOKEN,
    default=DefaultBotProperties(parse_mode=ParseMode.MARKDOWN),
)
dp = Dispatcher()


def is_allowed(message: Message) -> bool:
    if not OWNER_TELEGRAM_ID:
        return True
    return str(message.from_user.id) == OWNER_TELEGRAM_ID if message.from_user else False


async def guard(message: Message) -> bool:
    if is_allowed(message):
        return True
    await message.answer("⛔ У вас нет доступа к этому боту.")
    return False


def uzum() -> UzumClient:
    return UzumClient(UZUM_API_TOKEN, UZUM_API_BASE_URL)


def parse_shop_id_and_rest(text: str, default_shop_id: str | None = None) -> tuple[int | None, str]:
    parts = (text or "").split(maxsplit=2)
    args = parts[1:] if len(parts) > 1 else []

    if args and args[0].isdigit():
        shop_id = int(args[0])
        rest = args[1] if len(args) > 1 else ""
        return shop_id, rest

    if default_shop_id and default_shop_id.isdigit():
        return int(default_shop_id), " ".join(args)

    return None, " ".join(args)


async def send_api_error(message: Message, error: Exception) -> None:
    text = str(error)
    if len(text) > 3500:
        text = text[:3500] + "\n..."
    await message.answer(f"⚠️ Ошибка API:\n`{text}`")


@dp.message(Command("start", "help"))
async def start(message: Message) -> None:
    if not await guard(message):
        return

    default = f"`{DEFAULT_SHOP_ID}`" if DEFAULT_SHOP_ID else "не задан"
    await message.answer(
        "👋 Бот для Uzum Seller MVP.\n\n"
        "Команды:\n"
        "• `/pinguzum` — проверить API-токен\n"
        "• `/shops` — список магазинов\n"
        "• `/products <shop_id> [поиск]` — товары и остатки\n"
        "• `/lowstock <shop_id> [порог]` — товары с низким остатком\n"
        "• `/orders <shop_id> [CREATED|PACKING|COMPLETED|CANCELED]` — FBS/DBS заказы\n"
        "• `/export_products <shop_id>` — выгрузить товары в Excel\n\n"
        f"Магазин по умолчанию: {default}\n"
        "Если `DEFAULT_SHOP_ID` задан, `shop_id` можно не писать."
    )


@dp.message(Command("pinguzum"))
async def ping_uzum(message: Message) -> None:
    if not await guard(message):
        return
    if not UZUM_API_TOKEN:
        await message.answer("⚠️ `UZUM_API_TOKEN` не задан в переменных окружения.")
        return

    try:
        data = await uzum().get_shops()
        shops = extract_items(data)
        await message.answer(f"✅ Uzum API отвечает. Найдено магазинов/организаций: {len(shops)}")
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("shops"))
async def shops(message: Message) -> None:
    if not await guard(message):
        return

    try:
        data = await uzum().get_shops()
        items = extract_items(data)
        if not items:
            await message.answer("Ответ получен, но список магазинов не найден:\n```json\n" + compact_json_preview(data) + "\n```")
            return

        lines = [format_shop_line(item) for item in items[:30]]
        await message.answer("🏪 Ваши магазины:\n\n" + "\n".join(lines))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("products"))
async def products(message: Message) -> None:
    if not await guard(message):
        return

    shop_id, search_query = parse_shop_id_and_rest(message.text or "", DEFAULT_SHOP_ID)
    if shop_id is None:
        await message.answer("Напишите так: `/products <shop_id> [поиск]`\nНапример: `/products 12345 наушники`")
        return

    try:
        data = await uzum().get_products(shop_id, search_query=search_query, page=0, size=10)
        items = extract_items(data)
        if not items:
            await message.answer("Товары не найдены. Ответ API:\n```json\n" + compact_json_preview(data) + "\n```")
            return

        title = f"📦 Товары магазина `{shop_id}`"
        if search_query:
            title += f" по запросу “{search_query}”"
        lines = [format_product_line(item) for item in items[:10]]
        await message.answer(title + ":\n\n" + "\n\n".join(lines))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("lowstock"))
async def lowstock(message: Message) -> None:
    if not await guard(message):
        return

    shop_id, rest = parse_shop_id_and_rest(message.text or "", DEFAULT_SHOP_ID)
    if shop_id is None:
        await message.answer("Напишите так: `/lowstock <shop_id> [порог]`\nНапример: `/lowstock 12345 5`")
        return

    threshold = 5
    if rest.strip().isdigit():
        threshold = int(rest.strip())

    try:
        # First 100 products are enough for MVP. Later we can paginate all products.
        data = await uzum().get_products(shop_id, page=0, size=100)
        items = extract_items(data)
        low = [item for item in items if (get_stock_number(item) is not None and get_stock_number(item) <= threshold)]

        if not low:
            await message.answer(f"✅ В первых {len(items)} товарах нет остатков ≤ {threshold}.")
            return

        lines = [format_product_line(item) for item in low[:20]]
        await message.answer(f"⚠️ Низкие остатки ≤ {threshold}:\n\n" + "\n\n".join(lines))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("orders"))
async def orders(message: Message) -> None:
    if not await guard(message):
        return

    shop_id, rest = parse_shop_id_and_rest(message.text or "", DEFAULT_SHOP_ID)
    if shop_id is None:
        await message.answer("Напишите так: `/orders <shop_id> [status]`\nНапример: `/orders 12345 CREATED`")
        return

    status = rest.strip().upper() if rest.strip() else "CREATED"

    try:
        data = await uzum().get_fbs_orders(shop_id, status=status, page=0, size=10)
        items = extract_items(data)
        if not items:
            await message.answer(f"Заказы со статусом `{status}` не найдены. Ответ API:\n```json\n{compact_json_preview(data)}\n```")
            return

        lines = [format_order_line(item) for item in items[:10]]
        await message.answer(f"🧾 Заказы `{status}` для магазина `{shop_id}`:\n\n" + "\n".join(lines))
    except Exception as e:
        await send_api_error(message, e)


@dp.message(Command("export_products"))
async def export_products(message: Message) -> None:
    if not await guard(message):
        return

    shop_id, search_query = parse_shop_id_and_rest(message.text or "", DEFAULT_SHOP_ID)
    if shop_id is None:
        await message.answer("Напишите так: `/export_products <shop_id> [поиск]`\nНапример: `/export_products 12345`")
        return

    try:
        all_items: list[Any] = []
        # MVP: up to 5 pages x 100 = 500 products.
        for page in range(5):
            data = await uzum().get_products(shop_id, search_query=search_query, page=page, size=100)
            items = extract_items(data)
            if not items:
                break
            all_items.extend(items)
            if len(items) < 100:
                break

        if not all_items:
            await message.answer("Товары для экспорта не найдены.")
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Products"
        ws.append(["SKU/ID", "Название", "Цена", "Остаток", "Raw status"])

        for item in all_items:
            ws.append([
                pick(item, "skuId", "sku", "id", "productId"),
                pick(item, "title", "name", "skuTitle", "productTitle", "skuFullName"),
                pick(item, "price", "sellPrice", "fullPrice", "currentPrice", "skuPrice"),
                pick(item, "leftover", "leftovers", "quantity", "amount", "availableAmount", "stock", "stockAmount", "fbsAmount"),
                pick(item, "status", "state", "productStatus"),
            ])

        for column in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = min(max(max_len + 2, 12), 60)

        tmp_dir = Path(tempfile.gettempdir())
        filename = tmp_dir / f"uzum_products_{shop_id}.xlsx"
        wb.save(filename)

        await message.answer(f"✅ Экспортировано товаров: {len(all_items)}")
        await message.answer_document(FSInputFile(filename))
    except Exception as e:
        await send_api_error(message, e)


async def main() -> None:
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
