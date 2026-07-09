"""
AI-ответы на отзывы покупателей для Uzum Seller Assistant.

Положите файл рядом с main.py.

Переменные окружения:
    OPENAI_API_KEY=sk-...
    OPENAI_MODEL=gpt-5.4-mini   # необязательно
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Literal

from openai import AsyncOpenAI


ReviewLanguage = Literal["ru", "uz"]
ReviewTone = Literal["polite", "short", "warm", "protect_original", "apology"]

DEFAULT_MODEL = os.getenv("OPENAI_MODEL", "gpt-5.4-mini").strip() or "gpt-5.4-mini"


@dataclass(slots=True)
class ReviewReplyInput:
    review_text: str
    language: ReviewLanguage = "ru"
    tone: ReviewTone = "polite"
    product_name: str | None = None
    rating: int | None = None


def _get_client() -> AsyncOpenAI:
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError(
            "OPENAI_API_KEY не задан. Добавьте OPENAI_API_KEY в переменные окружения BotHost."
        )
    return AsyncOpenAI(api_key=api_key)


def _language_name(language: ReviewLanguage) -> str:
    if language == "uz":
        return "узбекском языке, латиницей, без кириллицы"
    return "русском языке"


def _tone_rules(tone: ReviewTone) -> str:
    if tone == "short":
        return "Ответ должен быть очень коротким: 1-2 предложения."
    if tone == "warm":
        return "Ответ должен быть тёплым, дружелюбным и заботливым."
    if tone == "protect_original":
        return (
            "Если клиент сомневается в оригинальности, аккуратно объясни, что продавец "
            "работает с оригинальным товаром или проверенными поставщиками. "
            "Не спорь и не обвиняй клиента."
        )
    if tone == "apology":
        return "Сделай акцент на извинении и готовности проверить ситуацию."
    return "Ответ должен быть вежливым, спокойным и нейтральным."


def _clean_review_text(text: str) -> str:
    text = (text or "").strip()
    if len(text) > 1500:
        text = text[:1500] + "..."
    return text


async def generate_review_reply(data: ReviewReplyInput) -> str:
    review_text = _clean_review_text(data.review_text)
    if not review_text:
        raise ValueError("Пустой текст отзыва.")

    product_part = f"\nТовар: {data.product_name}" if data.product_name else ""
    rating_part = f"\nОценка клиента: {data.rating}/5" if data.rating else ""

    instructions = (
        "Ты помощник продавца на маркетплейсе Uzum. "
        "Пиши готовые ответы на отзывы покупателей. "
        "Не пиши, что ты искусственный интеллект. "
        "Не обещай возврат денег, скидку, замену, компенсацию или действия, которые продавец не подтвердил. "
        "Не обвиняй клиента и не спорь с ним. "
        "Если отзыв негативный, извинись и скажи, что продавец проверит ситуацию. "
        "Если отзыв положительный, поблагодари за покупку. "
        "Ответ должен быть 2-4 предложения, без заголовка. "
        "Можно добавить один уместный эмодзи в конце."
    )

    prompt = f"""
Сгенерируй ответ на {_language_name(data.language)}.

Стиль:
{_tone_rules(data.tone)}

Данные отзыва:
{rating_part}{product_part}

Отзыв клиента:
{review_text}
""".strip()

    client = _get_client()
    response = await client.responses.create(
        model=DEFAULT_MODEL,
        instructions=instructions,
        input=prompt,
        max_output_tokens=350,
    )

    text = getattr(response, "output_text", "") or ""
    text = text.strip()

    if not text:
        raise RuntimeError("OpenAI вернул пустой ответ.")

    if len(text) > 1000:
        text = text[:1000].rsplit(" ", 1)[0].strip() + "..."

    return text


async def generate_review_reply_simple(
    review_text: str,
    *,
    language: ReviewLanguage = "ru",
    tone: ReviewTone = "polite",
    product_name: str | None = None,
    rating: int | None = None,
) -> str:
    return await generate_review_reply(
        ReviewReplyInput(
            review_text=review_text,
            language=language,
            tone=tone,
            product_name=product_name,
            rating=rating,
        )
    )
