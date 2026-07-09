"""
Интеграция AI-ответов на отзывы в существующий main.py.

Этот файл НЕ запускает отдельного бота. Он только добавляет handlers в ваш текущий dp.

В main.py нужно добавить:

    from review_ai_integration import setup_review_ai_handlers

    setup_review_ai_handlers(
        dp,
        menu_for_message=menu_for_message,
        get_user_language=get_user_language,
        require_active_subscription=require_active_subscription,
        upsert_from_message=upsert_from_message,
    )

Авто-патчер apply_ai_review_patch.py сделает это сам.
"""

from __future__ import annotations

from collections.abc import Awaitable, Callable
from html import escape
from typing import Any

from aiogram import F
from aiogram.filters import Command
from aiogram.types import Message

from ai_reviews import generate_review_reply_simple


def _arg_after_command(text: str | None) -> str:
    if not text:
        return ""
    parts = text.split(maxsplit=1)
    return parts[1].strip() if len(parts) > 1 else ""


def _menu(message: Message, menu_for_message: Callable[[Message], Any] | None):
    if not menu_for_message:
        return None
    try:
        return menu_for_message(message)
    except Exception:
        return None


async def _maybe_require_subscription(
    message: Message,
    upsert_from_message: Callable[[Message], int] | None,
    require_active_subscription: Callable[..., Awaitable[bool]] | None,
) -> bool:
    if not require_active_subscription:
        return True

    telegram_id = message.from_user.id if message.from_user else None
    try:
        if upsert_from_message:
            telegram_id = upsert_from_message(message)
    except Exception:
        pass

    try:
        if telegram_id is not None:
            return bool(await require_active_subscription(message, int(telegram_id)))
        return bool(await require_active_subscription(message))
    except TypeError:
        return bool(await require_active_subscription(message))


def _preferred_lang(
    message: Message,
    get_user_language: Callable[[int], str] | None,
    fallback: str = "ru",
) -> str:
    if fallback in {"ru", "uz"}:
        return fallback
    if not get_user_language or not message.from_user:
        return "ru"
    try:
        lang = get_user_language(int(message.from_user.id))
        return "uz" if str(lang).lower().startswith("uz") else "ru"
    except Exception:
        return "ru"


async def _send_help(
    message: Message,
    *,
    menu_for_message: Callable[[Message], Any] | None,
) -> None:
    await message.answer(
        "🤖 <b>AI-ответ на отзыв</b>\n\n"
        "Отправьте команду и текст отзыва:\n\n"
        "<code>/ai_review Товар пришёл без коробки, думала подделка</code>\n\n"
        "Команды:\n"
        "• <code>/ai_review текст</code> — ответ на русском\n"
        "• <code>/ai_review_uz текст</code> — ответ на узбекском\n"
        "• <code>/ai_review_short текст</code> — короткий ответ\n"
        "• <code>/ai_review_original текст</code> — если клиент пишет «не оригинал»\n\n"
        "Пока бот генерирует текст, а продавец копирует его в кабинет Uzum.",
        reply_markup=_menu(message, menu_for_message),
    )


async def _generate_and_send(
    message: Message,
    *,
    language: str,
    tone: str,
    menu_for_message: Callable[[Message], Any] | None,
    get_user_language: Callable[[int], str] | None,
    require_active_subscription: Callable[..., Awaitable[bool]] | None,
    upsert_from_message: Callable[[Message], int] | None,
) -> None:
    allowed = await _maybe_require_subscription(
        message,
        upsert_from_message=upsert_from_message,
        require_active_subscription=require_active_subscription,
    )
    if not allowed:
        return

    review_text = _arg_after_command(message.text)

    if not review_text:
        await _send_help(message, menu_for_message=menu_for_message)
        return

    lang = _preferred_lang(message, get_user_language, language)
    wait_text = "🤖 Javob tayyorlayapman..." if lang == "uz" else "🤖 Генерирую ответ..."
    await message.answer(wait_text)

    try:
        answer = await generate_review_reply_simple(
            review_text,
            language=lang,       # type: ignore[arg-type]
            tone=tone,           # type: ignore[arg-type]
        )

        if lang == "uz":
            result = (
                "🤖 <b>Tayyor javob:</b>\n\n"
                f"{escape(answer)}\n\n"
                "Nusxa ko‘chirib, Uzum kabinetiga joylashtiring."
            )
        else:
            result = (
                "🤖 <b>Готовый ответ:</b>\n\n"
                f"{escape(answer)}\n\n"
                "Скопируйте и вставьте в кабинет Uzum."
            )

        await message.answer(result, reply_markup=_menu(message, menu_for_message))

    except Exception as e:
        await message.answer(
            "⚠️ Не получилось сгенерировать ответ.\n\n"
            "Проверьте переменные окружения BotHost:\n"
            "<code>OPENAI_API_KEY</code>\n"
            "<code>OPENAI_MODEL</code>\n\n"
            f"Ошибка: <code>{escape(str(e)[:700])}</code>",
            reply_markup=_menu(message, menu_for_message),
        )


def setup_review_ai_handlers(
    dp,
    *,
    menu_for_message: Callable[[Message], Any] | None = None,
    get_user_language: Callable[[int], str] | None = None,
    require_active_subscription: Callable[..., Awaitable[bool]] | None = None,
    upsert_from_message: Callable[[Message], int] | None = None,
) -> None:
    """
    Регистрирует команды AI-ответов в существующем Dispatcher.
    """

    @dp.message(Command("ai_review"))
    async def ai_review_ru(message: Message) -> None:
        await _generate_and_send(
            message,
            language="ru",
            tone="polite",
            menu_for_message=menu_for_message,
            get_user_language=get_user_language,
            require_active_subscription=require_active_subscription,
            upsert_from_message=upsert_from_message,
        )

    @dp.message(Command("ai_review_uz"))
    async def ai_review_uz(message: Message) -> None:
        await _generate_and_send(
            message,
            language="uz",
            tone="polite",
            menu_for_message=menu_for_message,
            get_user_language=get_user_language,
            require_active_subscription=require_active_subscription,
            upsert_from_message=upsert_from_message,
        )

    @dp.message(Command("ai_review_short"))
    async def ai_review_short(message: Message) -> None:
        await _generate_and_send(
            message,
            language="ru",
            tone="short",
            menu_for_message=menu_for_message,
            get_user_language=get_user_language,
            require_active_subscription=require_active_subscription,
            upsert_from_message=upsert_from_message,
        )

    @dp.message(Command("ai_review_original"))
    async def ai_review_original(message: Message) -> None:
        await _generate_and_send(
            message,
            language="ru",
            tone="protect_original",
            menu_for_message=menu_for_message,
            get_user_language=get_user_language,
            require_active_subscription=require_active_subscription,
            upsert_from_message=upsert_from_message,
        )

    @dp.message(F.text == "🤖 AI-ответ на отзыв")
    async def ai_review_button(message: Message) -> None:
        await _send_help(message, menu_for_message=menu_for_message)
