from __future__ import annotations

import hashlib
import json
import math
from datetime import datetime, timezone
from typing import Any, Iterable


def number(value: Any) -> float | None:
    """Return a finite number without treating booleans as money."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        result = float(value)
    elif isinstance(value, str):
        cleaned = (
            value.replace("\u00a0", "")
            .replace(" ", "")
            .replace("UZS", "")
            .replace("сум", "")
            .replace("so‘m", "")
            .replace(",", ".")
            .strip()
        )
        try:
            result = float(cleaned)
        except (TypeError, ValueError):
            return None
    else:
        return None
    return result if math.isfinite(result) else None


def text(value: Any) -> str:
    return str(value or "").strip()


def enum_text(value: Any) -> str:
    if isinstance(value, dict):
        for key in ("value", "code", "status", "name", "text", "title", "type"):
            candidate = value.get(key)
            if candidate not in (None, ""):
                return text(candidate)
        return ""
    return text(value)


def normalize_key(value: Any) -> str:
    return text(value).lower()


def parse_api_datetime(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        parsed = value
    else:
        raw = text(value)
        if not raw:
            return None
        try:
            parsed = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        except ValueError:
            return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


def stock_record(row: dict[str, Any]) -> dict[str, Any]:
    """Normalize one SKU while keeping Uzum as the only cost source."""
    raw = row.get("raw") if isinstance(row.get("raw"), dict) else {}
    purchase_price = number(row.get("purchase_price"))
    if purchase_price is None:
        purchase_price = number(raw.get("purchasePrice"))
    if purchase_price is not None and purchase_price <= 0:
        purchase_price = None

    identifiers = [
        row.get("sku_id") or raw.get("skuId") or raw.get("id"),
        row.get("seller_item_code") or raw.get("sellerItemCode") or raw.get("article"),
        row.get("barcode") or raw.get("barcode"),
        row.get("sku_full_title") or raw.get("skuFullTitle"),
        row.get("sku_title") or raw.get("skuTitle"),
    ]
    aliases: list[str] = []
    for value in identifiers:
        key = normalize_key(value)
        if key and key not in {"-", "—"} and key not in aliases:
            aliases.append(key)

    fallback = "|".join(
        normalize_key(row.get(field))
        for field in ("product_id", "sku_full_title", "sku_title", "product_title")
    )
    primary_key = aliases[0] if aliases else hashlib.sha256(fallback.encode("utf-8")).hexdigest()[:24]
    return {
        "sku_key": primary_key,
        "aliases": aliases,
        "sku_id": text(row.get("sku_id") or raw.get("skuId") or raw.get("id")),
        "barcode": text(row.get("barcode") or raw.get("barcode")),
        "seller_item_code": text(
            row.get("seller_item_code") or raw.get("sellerItemCode") or raw.get("article")
        ),
        "sku_title": text(row.get("sku_full_title") or row.get("sku_title") or raw.get("skuTitle")),
        "product_title": text(row.get("product_title")),
        "purchase_price": purchase_price,
        "ikpu": text(row.get("ikpu") or raw.get("ikpu")),
        "paid_storage_price_item": number(
            row.get("paid_storage_price_item")
            if row.get("paid_storage_price_item") is not None
            else raw.get("paidStoragePriceItem")
        ),
        "paid_storage_amount": number(
            row.get("paid_storage_amount")
            if row.get("paid_storage_amount") is not None
            else raw.get("paidStorageAmount")
        ),
        "paid_storage": bool(row.get("paid_storage") or raw.get("pstorage")),
    }


def build_stock_records(rows: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    records: dict[str, dict[str, Any]] = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        record = stock_record(row)
        current = records.get(record["sku_key"])
        if current is None:
            records[record["sku_key"]] = record
            continue
        # Duplicate API rows describe the same SKU. Prefer the row that has a
        # documented purchase price and merge identifiers, never add prices.
        if current.get("purchase_price") is None and record.get("purchase_price") is not None:
            current["purchase_price"] = record["purchase_price"]
        current["aliases"] = list(dict.fromkeys([*current.get("aliases", []), *record.get("aliases", [])]))
        for field in (
            "sku_id",
            "barcode",
            "seller_item_code",
            "sku_title",
            "product_title",
            "ikpu",
            "paid_storage_price_item",
            "paid_storage_amount",
        ):
            if current.get(field) in (None, "") and record.get(field) not in (None, ""):
                current[field] = record[field]
        current["paid_storage"] = bool(current.get("paid_storage") or record.get("paid_storage"))
    return list(records.values())


def expense_items(data: Any) -> list[dict[str, Any]]:
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]
    if not isinstance(data, dict):
        return []
    for key in ("payments", "content", "items", "data", "rows"):
        value = data.get(key)
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
    for key in ("payload", "result"):
        nested = data.get(key)
        rows = expense_items(nested)
        if rows:
            return rows
    return []


def expense_category(row: dict[str, Any]) -> str:
    haystack = " ".join(
        enum_text(row.get(field)).lower()
        for field in ("name", "source", "code")
    )
    # Uzum may return operation names in Russian, English or Uzbek regardless
    # of the language selected in the bot.  Fold Uzbek apostrophe variants so
    # strings such as ``pulli targ‘ibot`` are classified consistently.
    folded = (
        haystack.replace("‘", "'")
        .replace("’", "'")
        .replace("ʻ", "'")
        .replace("ʼ", "'")
    )
    compact = folded.replace("'", "")
    # Order commission and delivery are already reflected in Finance payout.
    # Keep such ledger rows auditable but never deduct them for a second time.
    if any(word in folded for word in ("commission", "комисс", "komiss")):
        return "order_charge"
    if any(word in folded for word in ("logistic", "delivery", "логист", "достав", "yetkaz")):
        return "order_charge"
    if any(word in folded for word in ("storage", "хран", "saql", "ombor")):
        return "storage"
    if any(
        word in folded
        for word in ("advert", "реклам", "reklama", "promotion", "продвиж", "marketing")
    ) or "targibot" in compact:
        return "advertising"
    if any(word in folded for word in ("penalty", "fine", "штраф", "jarima")):
        return "penalty"
    return "other"


def expense_display_name(item: dict[str, Any], lang: str = "ru") -> str:
    """Return a user-language label instead of leaking Uzum's raw locale."""
    category = text(item.get("category") or "other")
    raw_name = text(item.get("name"))
    uz = str(lang or "ru").lower() == "uz"
    labels = {
        "storage": ("Хранение Uzum", "Uzum saqlash xizmati"),
        "advertising": ("Платное продвижение на Uzum", "Uzum'da pulli targ‘ibot"),
        "penalty": ("Штраф Uzum", "Uzum jarimasi"),
        "order_charge": ("Комиссия или логистика по заказу", "Buyurtma komissiyasi yoki logistikasi"),
        "other": ("Прочая операция Uzum", "Uzum boshqa operatsiyasi"),
    }
    ru_label, uz_label = labels.get(category, labels["other"])
    if category == "other":
        # Preserve an already localized description when it is useful.  Raw
        # Latin Uzbek/English text is replaced for Russian users.
        if uz and raw_name:
            return raw_name
        if not uz and any("а" <= char.lower() <= "я" or char.lower() == "ё" for char in raw_name):
            return raw_name
    return uz_label if uz else ru_label


def expense_display_source(item: dict[str, Any], lang: str = "ru") -> str:
    """Localize the compact source label shown in Telegram expense rows."""
    category = text(item.get("category") or "other")
    uz = str(lang or "ru").lower() == "uz"
    labels = {
        "storage": ("Хранение", "Saqlash"),
        "advertising": ("Продвижение", "Marketing"),
        "penalty": ("Штрафы", "Jarimalar"),
        "order_charge": ("Заказ", "Buyurtma"),
        "other": ("Uzum", "Uzum"),
    }
    ru_label, uz_label = labels.get(category, labels["other"])
    return uz_label if uz else ru_label


def normalize_expense(row: dict[str, Any]) -> dict[str, Any] | None:
    status = enum_text(row.get("status")).upper()
    payment_type = enum_text(row.get("type")).upper()
    amount = number(row.get("paymentPrice"))
    if amount is None:
        return None
    amount = abs(amount)
    booked_statuses = {
        "COMPLETED",
        "CONFIRMED",
        "SUCCESS",
        "SUCCEEDED",
        "PAID",
        "PROCESSED",
        "DONE",
    }
    if status in booked_statuses:
        booked = True
        signed_amount = -amount if payment_type == "INCOME" else amount
    else:
        # Unknown and intermediate statuses are deliberately excluded.  A
        # missing/renamed API status must not inflate confirmed expenses.
        booked = False
        signed_amount = 0.0
    identity = text(row.get("id") or row.get("externalId"))
    if not identity:
        identity = hashlib.sha256(
            json.dumps(row, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")
        ).hexdigest()[:24]
    return {
        "identity": identity,
        "name": text(row.get("name") or row.get("source") or row.get("code") or "Uzum expense"),
        "source": enum_text(row.get("source")),
        "status": status,
        "type": payment_type,
        "date": parse_api_datetime(row.get("dateService") or row.get("dateCreated")),
        "amount": amount,
        "signed_amount": signed_amount,
        "booked": booked,
        "category": expense_category(row),
        "raw": row,
    }


def summarize_expenses(rows: Iterable[dict[str, Any]]) -> dict[str, Any]:
    normalized: dict[str, dict[str, Any]] = {}
    for row in rows:
        item = normalize_expense(row)
        if item is not None:
            normalized[item["identity"]] = item
    booked = [item for item in normalized.values() if item["booked"]]
    pending = [item for item in normalized.values() if not item["booked"]]
    categories = {
        "storage": 0.0,
        "advertising": 0.0,
        "penalty": 0.0,
        "other": 0.0,
        "order_charge": 0.0,
    }
    for item in booked:
        categories[item["category"]] += float(item["signed_amount"])
        item["included_in_profit"] = item["category"] != "order_charge"
    total = sum(
        float(item["signed_amount"])
        for item in booked
        if item["category"] != "order_charge"
    )
    deductions = sum(
        float(item["signed_amount"])
        for item in booked
        if item["category"] != "order_charge"
        and float(item["signed_amount"]) > 0
    )
    refunds = sum(
        abs(float(item["signed_amount"]))
        for item in booked
        if item["category"] != "order_charge"
        and float(item["signed_amount"]) < 0
    )
    ordered = sorted(
        booked,
        key=lambda item: (item.get("date") or datetime.min.replace(tzinfo=timezone.utc), abs(float(item["signed_amount"]))),
        reverse=True,
    )
    return {
        "total": total,
        # Keep deductions and refunds separate for a readable profit bridge.
        # ``total`` remains the signed net value for backwards compatibility.
        "deductions": deductions,
        "refunds": refunds,
        "storage": categories["storage"],
        "advertising": categories["advertising"],
        "penalty": categories["penalty"],
        "other": categories["other"],
        "order_charge": categories["order_charge"],
        "booked_count": len(booked),
        "pending_count": len(pending),
        "rows": ordered,
        "pending_rows": pending,
    }


def supply_reminder_bucket(start_at: datetime | None, now: datetime) -> str | None:
    if start_at is None:
        return None
    now_utc = now.astimezone(timezone.utc)
    hours = (start_at.astimezone(timezone.utc) - now_utc).total_seconds() / 3600.0
    if 0 <= hours <= 3:
        return "3h"
    if 3 < hours <= 24:
        return "24h"
    return None


def return_reminder_bucket(
    start_at: datetime | None,
    now: datetime,
    *,
    storage_status: str = "",
) -> str | None:
    status = text(storage_status).upper()
    if status in {"COMPLETED"}:
        return None
    if start_at is None:
        return None
    now_utc = now.astimezone(timezone.utc)
    hours = (start_at.astimezone(timezone.utc) - now_utc).total_seconds() / 3600.0
    if hours <= 0:
        return "active" if status in {"ACTIVE", "EXPIRED"} else None
    if hours <= 24:
        return "1d"
    if hours <= 48:
        return "2d"
    if hours <= 72:
        return "3d"
    return None
