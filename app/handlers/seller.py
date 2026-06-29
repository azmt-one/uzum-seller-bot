from aiogram import Router, F
from aiogram.filters import Command
from aiogram.types import Message
from app.db import log_command
from app.security import deny_if_not_allowed
from app.services.uzum_client import (
    create_invoice,
    get_balance,
    get_orders,
    get_reviews,
    get_sales_summary,
    get_stock,
)

router = Router()


@router.message(Command("sales"))
@router.message(F.text == "📊 Продажи")
async def sales(message: Message):
    if await deny_if_not_allowed(message):
        return
    await log_command(message.from_user.id, "sales")
    data = await get_sales_summary("today")
    await message.answer(
        "📊 <b>Продажи сегодня</b>\n\n"
        f"Заказов: <b>{data.orders}</b>\n"
        f"Выручка: <b>{data.revenue:,}</b> сум\n"
        f"Возвраты: <b>{data.returns}</b>\n"
        f"Чистая сумма: <b>{data.net_revenue:,}</b> сум"
        .replace(",", " ")
    )


@router.message(Command("balance"))
@router.message(F.text == "💰 Баланс")
async def balance(message: Message):
    if await deny_if_not_allowed(message):
        return
    await log_command(message.from_user.id, "balance")
    data = await get_balance()
    await message.answer(
        "💰 <b>Баланс</b>\n\n"
        f"Доступно: <b>{data.available:,}</b> {data.currency}\n"
        f"В ожидании: <b>{data.pending:,}</b> {data.currency}"
        .replace(",", " ")
    )


@router.message(Command("orders"))
@router.message(F.text == "📦 Заказы")
async def orders(message: Message):
    if await deny_if_not_allowed(message):
        return
    await log_command(message.from_user.id, "orders")
    orders_list = await get_orders()
    text = "📦 <b>Заказы</b>\n\n"
    for order in orders_list:
        text += f"#{order.order_id} — {order.status} — {order.amount:,} сум\n".replace(",", " ")
    await message.answer(text)


@router.message(Command("stock"))
@router.message(F.text == "🏬 Остатки")
async def stock(message: Message):
    if await deny_if_not_allowed(message):
        return
    await log_command(message.from_user.id, "stock")
    stock_list = await get_stock()
    text = "🏬 <b>Остатки</b>\n\n"
    for item in stock_list:
        text += f"{item.sku} — {item.name}: <b>{item.quantity}</b> шт.\n"
    await message.answer(text)


@router.message(Command("reviews"))
@router.message(F.text == "⭐ Отзывы")
async def reviews(message: Message):
    if await deny_if_not_allowed(message):
        return
    await log_command(message.from_user.id, "reviews")
    reviews_list = await get_reviews()
    text = "⭐ <b>Отзывы</b>\n\n" + "\n\n".join(reviews_list)
    await message.answer(text)


@router.message(Command("invoice"))
async def invoice_command(message: Message):
    if await deny_if_not_allowed(message):
        return
    await log_command(message.from_user.id, "invoice")
    parts = (message.text or "").split(maxsplit=1)
    order_id = parts[1].strip() if len(parts) > 1 else "DEMO-001"
    invoice_text = await create_invoice(order_id)
    await message.answer(f"📄 <b>Накладная</b>\n\n<pre>{invoice_text}</pre>")


@router.message(F.text == "📄 Накладная")
async def invoice_button(message: Message):
    if await deny_if_not_allowed(message):
        return
    await message.answer(
        "Чтобы создать накладную, напишите команду так:\n\n"
        "<code>/invoice DEMO-001</code>\n\n"
        "Позже сделаем выбор заказа кнопками."
    )
