from aiogram.types import Message
from app.config import settings
from app.db import save_user


async def is_allowed(message: Message) -> bool:
    user = message.from_user
    if not user:
        return False

    allowed_ids = settings.allowed_user_ids
    allowed = user.id in allowed_ids if allowed_ids else True

    await save_user(
        telegram_id=user.id,
        first_name=user.first_name,
        username=user.username,
        is_allowed=allowed,
    )
    return allowed


async def deny_if_not_allowed(message: Message) -> bool:
    if await is_allowed(message):
        return False

    await message.answer(
        "⛔️ У вас нет доступа к этому боту.\n\n"
        f"Ваш Telegram ID: <code>{message.from_user.id}</code>\n"
        "Добавьте этот ID в ADMIN_IDS на сервере."
    )
    return True
