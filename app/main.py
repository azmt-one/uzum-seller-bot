import asyncio
import logging
from aiogram import Bot, Dispatcher
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties
from app.config import settings
from app.db import init_db
from app.handlers import common, seller


async def main() -> None:
    logging.basicConfig(level=logging.INFO)

    if not settings.telegram_token or settings.telegram_token == "PASTE_TELEGRAM_BOT_TOKEN_HERE":
        raise RuntimeError("Не указан TELEGRAM_TOKEN. Добавьте токен в .env или переменные окружения bothost.ru")

    await init_db()

    bot = Bot(
        token=settings.telegram_token,
        default=DefaultBotProperties(parse_mode=ParseMode.HTML),
    )
    dp = Dispatcher()
    dp.include_router(common.router)
    dp.include_router(seller.router)

    logging.info("UZUM Seller Bot started")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
