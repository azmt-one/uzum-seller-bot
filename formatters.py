from __future__ import annotations

from html import escape
from typing import Any


def extract_items(data: Any) -> list[Any]:
    """Best-effort extractor for arrays in different Uzum response wrappers."""
    if isinstance(data, list):
        return data

    if not isinstance(data, dict):
        return []

    for key in (
        "payload",
        "content",
        "items",
        "products",
        "productList",
        "shopProducts",
        "skuList",
        "data",
        "result",
        "rows",
        "orders",
        "orderItems",
        "elements",
        "values",
    ):
        value = data.get(key)
        if isinstance(value, list):
            return value
        if isinstance(value, dict):
            nested = extract_items(value)
            if nested:
                return nested

    return []


def pick(obj: Any, *keys: str, default: Any = "—") -> Any:
    if not isinstance(obj, dict):
        return default

    lower_map = {str(k).lower(): k for k in obj.keys()}
    for key in keys:
        actual = lower_map.get(key.lower())
        if actual is not None:
            value = obj.get(actual)
            if value not in (None, ""):
                return value

    for value in obj.values():
        if isinstance(value, dict):
            found = pick(value, *keys, default=None)
            if found not in (None, ""):
                return found

    return default


def compact_json_preview(obj: Any, limit: int = 700) -> str:
    import json

    text = json.dumps(obj, ensure_ascii=False, indent=2, default=str)
    return text if len(text) <= limit else text[:limit] + "\n..."


def format_money(value: Any) -> str:
    if value in (None, "—", ""):
        return "—"
    try:
        return f"{float(value):,.0f}".replace(",", " ")
    except Exception:
        return escape(str(value))


def safe(value: Any) -> str:
    return escape(str(value))


def format_shop_line(shop: Any) -> str:
    shop_id = pick(shop, "id", "shopId", "sellerId", "organizationId")
    title = pick(shop, "title", "name", "shopTitle", "organizationName", "legalName")
    return f"• ID: <code>{safe(shop_id)}</code> — {safe(title)}"


def format_product_line(product: Any) -> str:
    sku_id = pick(product, "skuId", "sku", "id", "productId")
    title = pick(product, "title", "name", "skuTitle", "productTitle", "skuFullName")
    price = pick(product, "price", "sellPrice", "fullPrice", "currentPrice", "skuPrice", default="—")
    stock = pick(product, "leftover", "leftovers", "quantity", "amount", "availableAmount", "stock", "stockAmount", "fbsAmount", default="—")
    return f"• <code>{safe(sku_id)}</code> — {safe(title)}\n  Цена: {format_money(price)} сум | Остаток: {safe(stock)}"


def get_stock_number(product: Any) -> float | None:
    value = pick(product, "leftover", "leftovers", "quantity", "amount", "availableAmount", "stock", "stockAmount", "fbsAmount", default=None)
    try:
        return float(value)
    except Exception:
        return None


def format_order_line(order: Any) -> str:
    order_id = pick(order, "id", "orderId", "customerOrderId", "number")
    status = pick(order, "status", "orderStatus")
    created_at = pick(order, "createdAt", "creationDate", "date", "dateCreated")
    total = pick(order, "total", "totalPrice", "price", "amount", default="—")
    return f"• Заказ <code>{safe(order_id)}</code> | {safe(status)} | {safe(created_at)} | {format_money(total)} сум"
