from __future__ import annotations

import json
from html import escape
from typing import Any

STATUS_KEYS = ("status", "state", "productStatus", "skuStatus", "availability")


def extract_items(data: Any) -> list[Any]:
    if isinstance(data, list):
        return data
    if not isinstance(data, dict):
        return []
    for key in (
        "payload", "content", "items", "products", "productList", "shopProducts", "data",
        "result", "rows", "orders", "orderItems", "elements", "values",
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
    text = json.dumps(obj, ensure_ascii=False, indent=2, default=str)
    return text if len(text) <= limit else text[:limit] + "\n..."


def normalize_value(value: Any) -> Any:
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
    value = normalize_value(value)
    if value is None or value == "—":
        return ""
    if isinstance(value, (int, float, str)):
        return value
    return json.dumps(value, ensure_ascii=False, default=str)


def to_number(value: Any) -> float | None:
    value = normalize_value(value)
    if value in (None, "", "—"):
        return None
    try:
        return float(value)
    except Exception:
        return None


def clean_num(value: float | int | None) -> str:
    if value is None:
        return "—"
    try:
        value = float(value)
        return str(int(value)) if value.is_integer() else str(value)
    except Exception:
        return str(value)


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


def find_status(obj: Any) -> str | None:
    value = normalize_value(pick(obj, *STATUS_KEYS, default=None))
    if value not in (None, "", "—"):
        return str(value)
    if isinstance(obj, dict):
        for k, v in obj.items():
            if str(k).lower() in {"value", "text"} and isinstance(v, str) and (
                "STOCK" in v.upper() or v.upper() in {"ACTIVE", "INACTIVE"}
            ):
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


def status_display(value: Any) -> str:
    value = normalize_value(value)
    mapping = {
        "IN_STOCK": "в продаже",
        "OUT_OF_STOCK": "нет в продаже",
        "ACTIVE": "активен",
        "INACTIVE": "неактивен",
        "ON_MODERATION": "на модерации",
    }
    return mapping.get(str(value).upper(), str(value))


def first_number(obj: dict[str, Any], *keys: str) -> float | None:
    for key in keys:
        if key in obj:
            n = to_number(obj.get(key))
            if n is not None:
                return n
    lower_map = {str(k).lower(): k for k in obj.keys()}
    for key in keys:
        actual = lower_map.get(key.lower())
        if actual is not None:
            n = to_number(obj.get(actual))
            if n is not None:
                return n
    return None


def format_shop_line(shop: Any) -> str:
    shop_id = pick(shop, "id", "shopId", "sellerId", "organizationId")
    title = pick(shop, "title", "name", "shopTitle", "organizationName", "legalName")
    return f"• ID: `{safe(shop_id)}` — {safe(title)}"


def flatten_sku_rows(products: list[Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for product in products:
        if not isinstance(product, dict):
            continue
        product_id = product.get("productId") or product.get("id")
        product_title = pick(product, "productTitle", "title", "name", default="—")
        product_status = find_status(product)
        category = pick(product, "category", default="—")
        sku_list = product.get("skuList")
        if not isinstance(sku_list, list):
            continue
        for sku in sku_list:
            if not isinstance(sku, dict):
                continue
            total = first_number(sku, "quantityAvailable", "quantityActive", "availableQuantity", "availableAmount", "quantity")
            active = first_number(sku, "quantityActive", "quantityAvailable")
            fbs = first_number(sku, "quantityFbs", "fbsQuantity", "quantityDbs", "dbsQuantity")
            additional = first_number(sku, "quantityAdditional")
            explicit_fbo = first_number(sku, "quantityFbo", "quantityFBO", "fboQuantity", "quantityWarehouse", "warehouseQuantity")
            total_n = total if total is not None else 0.0
            fbs_n = fbs if fbs is not None else 0.0
            if explicit_fbo is not None:
                fbo = explicit_fbo
            elif total is not None:
                fbo = max(total_n - fbs_n, 0.0)
            else:
                fbo = None
            sku_status = find_status(sku) or product_status
            rows.append({
                "product_id": product_id,
                "sku_id": sku.get("skuId") or sku.get("id"),
                "barcode": sku.get("barcode"),
                "seller_item_code": sku.get("sellerItemCode") or sku.get("article"),
                "product_title": sku.get("productTitle") or product_title,
                "sku_title": sku.get("skuTitle") or sku.get("skuFullTitle") or sku.get("title") or "—",
                "sku_full_title": sku.get("skuFullTitle") or sku.get("skuTitle") or "—",
                "category": category,
                "price": sku.get("price") or sku.get("marketPrice"),
                "market_price": sku.get("marketPrice"),
                "total": total,
                "active": active,
                "fbo": fbo,
                "fbs": fbs,
                "additional": additional,
                "sold": first_number(sku, "quantitySold"),
                "returned": first_number(sku, "quantityReturned"),
                "missing": first_number(sku, "quantityMissing"),
                "defected": first_number(sku, "quantityDefected"),
                "pending": first_number(sku, "quantityPending"),
                "status": sku_status,
                "raw": sku,
            })
    return rows


def get_stock_number(product: Any) -> float | None:
    if isinstance(product, dict) and isinstance(product.get("skuList"), list):
        total = 0.0
        found = False
        for sku in product["skuList"]:
            if isinstance(sku, dict):
                n = first_number(sku, "quantityAvailable", "quantityActive")
                if n is not None:
                    total += n
                    found = True
        return total if found else None
    return None


def format_product_line(product: Any) -> str:
    product_id = pick(product, "productId", "id")
    title = pick(product, "title", "name", "productTitle", default="—")
    status = find_status(product)
    sku_rows = flatten_sku_rows([product]) if isinstance(product, dict) else []
    if sku_rows:
        total = sum((r["total"] or 0) for r in sku_rows if r["total"] is not None)
        fbo = sum((r["fbo"] or 0) for r in sku_rows if r["fbo"] is not None)
        fbs = sum((r["fbs"] or 0) for r in sku_rows if r["fbs"] is not None)
        price = sku_rows[0].get("price")
        return (
            f"• `{safe(product_id)}` — {safe(title)}\n"
            f"Цена: {format_money(price)} сум | FBO: {safe(clean_num(fbo))} | "
            f"FBS/DBS: {safe(clean_num(fbs))} | Итого: {safe(clean_num(total))}"
            + (f" | Статус: {safe(status_display(status))}" if status else "")
        )
    price = pick(product, "price", "sellPrice", "fullPrice", "currentPrice", "skuPrice", default="—")
    return f"• `{safe(product_id)}` — {safe(title)}\nЦена: {format_money(price)} сум" + (
        f" | Статус: {safe(status_display(status))}" if status else ""
    )


def format_sku_stock_line(row: dict[str, Any], mode: str = "all") -> str:
    sku_id = row.get("sku_id")
    title = row.get("sku_full_title") or row.get("sku_title") or row.get("product_title")
    price = row.get("price")
    fbo = row.get("fbo")
    fbs = row.get("fbs")
    total = row.get("total")
    active = row.get("active")
    status = row.get("status")
    if mode == "fbo":
        stock_part = f"FBO: {safe(clean_num(fbo))}"
    elif mode == "fbs":
        stock_part = f"FBS/DBS: {safe(clean_num(fbs))}"
    else:
        stock_part = f"FBO: {safe(clean_num(fbo))} | FBS/DBS: {safe(clean_num(fbs))} | Итого: {safe(clean_num(total))}"
    extra = f" | Активно: {safe(clean_num(active))}" if active is not None and active != total else ""
    return (
        f"• `{safe(sku_id)}` — {safe(title)}\n"
        f"{stock_part}{extra} | Цена: {format_money(price)} сум"
        + (f" | Статус: {safe(status_display(status))}" if status else "")
    )


def format_order_line(order: Any) -> str:
    order_id = pick(order, "id", "orderId", "customerOrderId", "number")
    status = pick(order, "status", "orderStatus")
    created_at = pick(order, "createdAt", "creationDate", "date", "dateCreated")
    total = pick(order, "total", "totalPrice", "price", "amount", default="—")
    return f"• Заказ `{safe(order_id)}` | {safe(status)} | {safe(created_at)} | {format_money(total)} сум"
