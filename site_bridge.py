from __future__ import annotations

"""Telegram ↔ Seller Assistant Web bridge.

Works with the user's existing aiogram bot without replacing its business logic.

Install once in the existing bot:
    from site_bridge import install_site_bridge
    install_site_bridge(globals())

For the loader/wrapper version of main.py use:
    namespace = _exec_original_main()
    install_site_bridge(namespace)

Required bot environment variables:
    WEB_APP_URL=https://your-site.example
    WEB_SYNC_SECRET=<same random secret as BOT_SYNC_SECRET on the website>
"""

import asyncio
import base64
import hashlib
import hmac
import json
import os
import time
import urllib.request
from datetime import datetime, timezone
from typing import Any

from aiogram import F
from aiogram.filters import Command
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup, Message, WebAppInfo

WEB_APP_URL = os.getenv("WEB_APP_URL", "").strip().rstrip("/")
WEB_SYNC_SECRET = os.getenv("WEB_SYNC_SECRET", "").strip()


def _row_value(row: Any, name: str, default: Any = None) -> Any:
    if row is None:
        return default
    if isinstance(row, dict):
        return row.get(name, default)
    try:
        return row[name]
    except Exception:
        return getattr(row, name, default)


def _iso(value: Any) -> str | None:
    if not value:
        return None
    if isinstance(value, datetime):
        dt = value
    else:
        try:
            dt = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        except ValueError:
            return str(value)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc).isoformat()


def _subscription_payload(telegram_id: int, namespace: dict[str, Any]) -> tuple[str, str | None]:
    get_subscription = namespace.get("get_subscription_row")
    subscription = None
    if callable(get_subscription):
        try:
            subscription = get_subscription(telegram_id)
        except Exception:
            subscription = None
    if not subscription:
        return "unknown", None

    blocked = bool(int(_row_value(subscription, "blocked", 0) or 0))
    dates = [_row_value(subscription, "subscription_until"), _row_value(subscription, "trial_until")]
    parsed: list[datetime] = []
    for value in dates:
        if not value:
            continue
        try:
            dt = value if isinstance(value, datetime) else datetime.fromisoformat(str(value).replace("Z", "+00:00"))
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            parsed.append(dt.astimezone(timezone.utc))
        except Exception:
            continue
    until = max(parsed) if parsed else None
    if blocked:
        state = "blocked"
    elif until and until > datetime.now(timezone.utc):
        state = "active"
    else:
        state = "expired"
    return state, _iso(until)


def _cost_rows(telegram_id: int, shop_id: int | None, namespace: dict[str, Any]) -> list[dict[str, Any]]:
    if not shop_id:
        return []
    getter = namespace.get("get_unit_cost_map")
    if not callable(getter):
        return []
    try:
        values = getter(telegram_id, int(shop_id)) or {}
    except Exception:
        return []
    rows: list[dict[str, Any]] = []
    iterable = values.items() if isinstance(values, dict) else []
    for sku, row in iterable:
        rows.append({
            "sku_id": str(sku),
            "product_name": str(_row_value(row, "title", "") or ""),
            "unit_cost": float(_row_value(row, "cost", 0) or 0),
        })
        if len(rows) >= 3000:
            break
    return rows


def _payload_for(message: Message, namespace: dict[str, Any], *, sensitive: bool) -> dict[str, Any]:
    telegram_id = int(message.from_user.id)
    db = namespace.get("db")
    user = None
    if db is not None and hasattr(db, "get_user"):
        try:
            user = db.get_user(telegram_id)
        except Exception:
            user = None

    get_language = namespace.get("get_user_language")
    locale = "ru"
    if callable(get_language):
        try:
            locale = "uz" if str(get_language(telegram_id)).lower().startswith("uz") else "ru"
        except Exception:
            pass

    subscription_state, subscription_until = _subscription_payload(telegram_id, namespace)
    default_shop_id = _row_value(user, "default_shop_id")
    encrypted_token = _row_value(user, "uzum_token_encrypted", "") or _row_value(user, "encrypted_token", "")
    payload: dict[str, Any] = {
        "telegram_id": telegram_id,
        "first_name": message.from_user.first_name or "",
        "last_name": message.from_user.last_name or "",
        "username": message.from_user.username or "",
        "locale": locale,
        "subscription_until": subscription_until,
        "subscription_state": subscription_state,
        "default_shop_id": int(default_shop_id) if str(default_shop_id or "").isdigit() else None,
        "iat": int(time.time()),
    }
    if sensitive:
        payload["encrypted_token"] = str(encrypted_token or "")
        payload["costs"] = _cost_rows(telegram_id, payload["default_shop_id"], namespace)
    return payload


def _validate_settings() -> None:
    if not WEB_APP_URL:
        raise RuntimeError("WEB_APP_URL не задан в переменных бота")
    if not WEB_APP_URL.startswith("https://"):
        raise RuntimeError("WEB_APP_URL должен начинаться с https://")
    if len(WEB_SYNC_SECRET) < 32:
        raise RuntimeError("WEB_SYNC_SECRET должен содержать не менее 32 символов")


def signed_web_url(message: Message, namespace: dict[str, Any]) -> str:
    _validate_settings()
    # Sensitive fields are intentionally never placed in the URL/history/logs.
    raw = json.dumps(_payload_for(message, namespace, sensitive=False), ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    payload = base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")
    signature = hmac.new(WEB_SYNC_SECRET.encode("utf-8"), payload.encode("ascii"), hashlib.sha256).hexdigest()
    return f"{WEB_APP_URL}/auth/telegram/bridge?payload={payload}&sig={signature}"


def _sync_request(payload: dict[str, Any]) -> None:
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    signature = hmac.new(WEB_SYNC_SECRET.encode("utf-8"), raw, hashlib.sha256).hexdigest()
    request = urllib.request.Request(
        f"{WEB_APP_URL}/api/integrations/telegram/sync",
        data=raw,
        method="POST",
        headers={
            "Content-Type": "application/json",
            "Accept": "application/json",
            "X-Seller-Signature": signature,
            "User-Agent": "uzum-seller-assistant-bot/site-bridge",
        },
    )
    with urllib.request.urlopen(request, timeout=25) as response:
        if response.status >= 300:
            raise RuntimeError(f"Сайт вернул HTTP {response.status}")
        response.read()


async def sync_to_site(message: Message, namespace: dict[str, Any]) -> None:
    _validate_settings()
    payload = _payload_for(message, namespace, sensitive=True)
    await asyncio.to_thread(_sync_request, payload)


def _button(message: Message, namespace: dict[str, Any]) -> InlineKeyboardMarkup:
    url = signed_web_url(message, namespace)
    locale = _payload_for(message, namespace, sensitive=False).get("locale")
    label = "🌐 Veb-kabinetni ochish" if locale == "uz" else "🌐 Открыть веб-кабинет"
    return InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=label, web_app=WebAppInfo(url=url))]])


def install_site_bridge(namespace: dict[str, Any]) -> None:
    """Register web-cabinet handlers on the bot's existing Dispatcher."""
    dp = namespace.get("dp")
    if dp is None:
        raise RuntimeError("site_bridge: в main.py не найден dp = Dispatcher()")

    @dp.message(Command("site", "web", "cabinet"))
    async def open_site_command(message: Message) -> None:
        try:
            # Server-to-server transfer keeps the token/costs out of URLs.
            await sync_to_site(message, namespace)
            markup = _button(message, namespace)
        except Exception as exc:
            await message.answer(f"⚠️ Веб-кабинет пока не настроен: {exc}")
            return
        locale = _payload_for(message, namespace, sensitive=False).get("locale")
        text = (
            "🌐 <b>Seller Assistant veb-kabineti</b>\n\nBot va sayt bitta akkauntda ishlaydi. "
            "Do‘kon, til, obuna, Uzum ulanishi va tannarxlar sinxronlanadi."
            if locale == "uz"
            else "🌐 <b>Веб-кабинет Seller Assistant</b>\n\nБот и сайт работают под одним аккаунтом. "
                 "Магазин, язык, подписка, подключение Uzum и себестоимость синхронизируются."
        )
        await message.answer(text, reply_markup=markup)

    @dp.message(F.text.in_({"🌐 Веб-кабинет", "🌐 Открыть сайт", "🌐 Veb-kabinet"}))
    async def open_site_button(message: Message) -> None:
        await open_site_command(message)
