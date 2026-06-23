from __future__ import annotations

import json
from html import escape
from typing import Any


STOCK_KEYS = (
    "leftover",
    "leftovers",
    "quantity",
    "availableQuantity",
    "available_quantity",
    "availableAmount",
    "available_amount",
    "amount",
    "stock",
    "stocks",
    "stockAmount",
    "stock_amount",
    "fbsAmount",
    "fbs_amount",
    "qty",
    "count",
)

STATUS_KEYS = ("status", "state", "productStatus", "skuStatus", "availability")


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

    # One level down: many Uzum objects nest sku/product/shop data.
    for value in obj.values():
        if isinstance(value, dict):
            found = pick(value, *keys, default=None)
            if found not in (None, ""):
                return found

    return default


def compact_json_preview(obj: Any, limit: int = 700) -> str:
    text = json.dumps(obj, ensure_ascii=False, indent=2, default=str)
    return text if len(text) <= limit else text[:limit] + "\n..."


def normalize_value(value: Any) -> Any:
    """Convert Uzum wrapper dicts like {'title':'Sotuvda','value':'IN_STOCK'} to a simple value."""
    if value is None:
        return "—"
    if isinstance(value, dict):
        for key in ("value", "text", "name", "title", "label"):
            v = value.get(key)
            if v not in (None, ""):
                return v
        return json.dumps(value, ensure_ascii=False, default=str)
    if isinstance(value, list):
        simple = [normalize_value(v) for v in value]
        return ", ".join(str(v) for v in simple if v not in (None, "", "—")) or "—"
    return value


def excel_value(value: Any) -> str | int | float:
    """openpyxl cannot write dict/list values directly."""
    value = normalize_value(value)
    if value is None or value == "—":
        return ""
    if isinstance(value, (int, float, str)):
        return value
    return json.dumps(value, ensure_ascii=False, default=str)


def format_money(value: Any) -> str:
    value = normalize_value(value)
    if value in (None, "—", ""):
        return "—"
    try:
        return f"{float(value):,.0f}".replace(",", " ")
    except Exception:
        return escape(str(value))


def safe(value: Any) -> str:
    return escape(str(normalize_value(value)))


def find_numeric_by_keys(obj: Any, keys: tuple[str, ...]) -> float | None:
    if isinstance(obj, dict):
        lower_map = {str(k).lower(): k for k in obj.keys()}
        for key in keys:
            actual = lower_map.get(key.lower())
            if actual is None:
                continue
            value = normalize_value(obj.get(actual))
            # Do not treat status strings as stock.
            if isinstance(value, str) and value.upper() in {"IN_STOCK", "OUT_OF_STOCK", "ACTIVE", "INACTIVE"}:
                continue
            try:
                return float(value)
            except Exception:
                pass
        for value in obj.values():
            found = find_numeric_by_keys(value, keys)
            if found is not None:
                return found
    elif isinstance(obj, list):
        for value in obj:
            found = find_numeric_by_keys(value, keys)
            if found is not None:
                return found
    return None


def find_status(obj: Any) -> str | None:
    value = normalize_value(pick(obj, *STATUS_KEYS, default=None))
    if value not in (None, "", "—"):
        return str(value)

    # Uzum sometimes returns status in additional/colorized object or attributes.
    if isinstance(obj, dict):
        for k, v in obj.items():
            if str(k).lower() in {"value", "text"} and isinstance(v, str) and ("STOCK" in v.upper() or v.upper() in {"ACTIVE", "INACTIVE"}):
                return v
            if isinstance(v, (dict, list)):
                nested = find_status(v)
                if nested:
                    return nested
    elif isinstance(obj, list):
        for v in obj:
            nested = find_status(v)
            if nested:
                return nested
    return None


def stock_display(product: Any) -> str:
    number = get_stock_number(product)
    if number is not None:
        # Keep integer-looking values clean.
        return str(int(number)) if number.is_integer() else str(number)

    status = find_status(product)
    if status:
        mapping = {
            "IN_STOCK": "в продаже",
            "OUT_OF_STOCK": "нет в продаже",
            "ACTIVE": "активен",
            "INACTIVE": "неактивен",
        }
        return mapping.get(status.upper(), status)

    return "не передан API"


def has_any_numeric_stock(products: list[Any]) -> bool:
    return any(get_stock_number(p) is not None for p in products)


def format_shop_line(shop: Any) -> str:
    shop_id = pick(shop, "id", "shopId", "sellerId", "organizationId")
    title = pick(shop, "title", "name", "shopTitle", "organizationName", "legalName")
    return f"• ID: <code>{safe(shop_id)}</code> — {safe(title)}"


def format_product_line(product: Any) -> str:
    sku_id = pick(product, "skuId", "sku", "id", "productId")
    title = pick(product, "title", "name", "skuTitle", "productTitle", "skuFullName")
    price = pick(product, "price", "sellPrice", "fullPrice", "currentPrice", "skuPrice", default="—")
    status = find_status(product)
    line = f"• <code>{safe(sku_id)}</code> — {safe(title)}\n  Цена: {format_money(price)} сум | Остаток: {safe(stock_display(product))}"
    if status and get_stock_number(product) is None:
        line += f" | Статус: {safe(status)}"
    return line


def get_stock_number(product: Any) -> float | None:
    return find_numeric_by_keys(product, STOCK_KEYS)


def format_order_line(order: Any) -> str:
    order_id = pick(order, "id", "orderId", "customerOrderId", "number")
    status = pick(order, "status", "orderStatus")
    created_at = pick(order, "createdAt", "creationDate", "date", "dateCreated")
    total = pick(order, "total", "totalPrice", "price", "amount", default="—")
    return f"• Заказ <code>{safe(order_id)}</code> | {safe(status)} | {safe(created_at)} | {format_money(total)} сум"
