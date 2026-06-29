from aiogram import Router, F
from aiogram.filters import Command
from aiogram.types import Message
from app.db import log_command
from app.keyboards import main_menu
from app.security import deny_if_not_allowed

router = Router()


@router.message(Command("start"))
async def start(message: Message):
    if await deny_if_not_allowed(message):
        return
    await log_command(message.from_user.id, "/start")
    await message.answer(
        "👋 Бот для UZUM Seller запущен.\n\n"
        "Сейчас работает первая версия: кнопки, доступ по ID, база SQLite и демо-разделы.\n"
        "Следующий шаг — подключить реальные данные UZUM.",
        reply_markup=main_menu(),
    )


@router.message(Command("id"))
async def my_id(message: Message):
    await message.answer(f"Ваш Telegram ID: <code>{message.from_user.id}</code>")


@router.message(Command("help"))
@router.message(F.text == "🆘 Помощь")
async def help_message(message: Message):
    if await deny_if_not_allowed(message):
        return
    await log_command(message.from_user.id, "help")
    await message.answer(
        "Команды:\n"
        "/start — открыть меню\n"
        "/id — узнать свой Telegram ID\n"
        "/sales — продажи\n"
        "/balance — баланс\n"
        "/orders — заказы\n"
        "/stock — остатки\n"
        "/reviews — отзывы\n"
        "/invoice DEMO-001 — создать демо-накладную\n\n"
        "Важно: реальные данные UZUM появятся после подключения доступа к кабинету/API."
    )


@router.message(F.text == "⚙️ Настройки")
async def settings_page(message: Message):
    if await deny_if_not_allowed(message):
        return
    await log_command(message.from_user.id, "settings")
    await message.answer(
        "⚙️ Настройки первой версии:\n\n"
        "• доступ задается через ADMIN_IDS\n"
        "• база хранится в BOT_DB_PATH\n"
        "• токен Telegram хранится в TELEGRAM_TOKEN\n"
        "• UZUM_TOKEN пока пустой — подключим позже"
    )
