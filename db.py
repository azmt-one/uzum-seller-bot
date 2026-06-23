from __future__ import annotations

import json
import sqlite3
from datetime import datetime, timezone
from typing import Any

from cryptography.fernet import Fernet


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class Database:
    def __init__(self, path: str = "bot.db") -> None:
        self.path = path
        self._init()

    def connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init(self) -> None:
        with self.connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS users (
                    telegram_id INTEGER PRIMARY KEY,
                    username TEXT,
                    first_name TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    uzum_token_encrypted TEXT,
                    token_connected_at TEXT,
                    default_shop_id INTEGER,
                    trial_started_at TEXT,
                    trial_until TEXT,
                    trial_used INTEGER DEFAULT 0,
                    plan TEXT,
                    subscription_until TEXT
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS shops (
                    telegram_id INTEGER NOT NULL,
                    shop_id INTEGER NOT NULL,
                    title TEXT,
                    raw_json TEXT,
                    updated_at TEXT NOT NULL,
                    PRIMARY KEY (telegram_id, shop_id)
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS payments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    telegram_id INTEGER NOT NULL,
                    plan TEXT NOT NULL,
                    amount INTEGER NOT NULL,
                    currency TEXT NOT NULL,
                    provider_payment_id TEXT,
                    paid_at TEXT NOT NULL,
                    subscription_until TEXT
                )
                """
            )
            conn.commit()

    def upsert_user(self, telegram_id: int, username: str | None = None, first_name: str | None = None) -> None:
        ts = now_iso()
        with self.connect() as conn:
            conn.execute(
                """
                INSERT INTO users (telegram_id, username, first_name, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(telegram_id) DO UPDATE SET
                    username=excluded.username,
                    first_name=excluded.first_name,
                    updated_at=excluded.updated_at
                """,
                (telegram_id, username, first_name, ts, ts),
            )
            conn.commit()

    def get_user(self, telegram_id: int) -> sqlite3.Row | None:
        with self.connect() as conn:
            return conn.execute("SELECT * FROM users WHERE telegram_id = ?", (telegram_id,)).fetchone()

    def save_connection(self, telegram_id: int, encrypted_token: str, shops: list[dict[str, Any]]) -> int | None:
        ts = now_iso()
        default_shop_id: int | None = None

        normalized_shops: list[tuple[int, str, str]] = []
        for shop in shops:
            shop_id = extract_shop_id(shop)
            if shop_id is None:
                continue
            if default_shop_id is None:
                default_shop_id = shop_id
            title = extract_shop_title(shop)
            normalized_shops.append((shop_id, title, json.dumps(shop, ensure_ascii=False, default=str)))

        with self.connect() as conn:
            conn.execute(
                """
                UPDATE users
                SET uzum_token_encrypted = ?,
                    token_connected_at = ?,
                    default_shop_id = ?,
                    updated_at = ?
                WHERE telegram_id = ?
                """,
                (encrypted_token, ts, default_shop_id, ts, telegram_id),
            )
            conn.execute("DELETE FROM shops WHERE telegram_id = ?", (telegram_id,))
            conn.executemany(
                """
                INSERT INTO shops (telegram_id, shop_id, title, raw_json, updated_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                [(telegram_id, shop_id, title, raw, ts) for shop_id, title, raw in normalized_shops],
            )
            conn.commit()

        return default_shop_id

    def get_encrypted_token(self, telegram_id: int) -> str | None:
        row = self.get_user(telegram_id)
        if not row:
            return None
        return row["uzum_token_encrypted"]

    def has_uzum_connection(self, telegram_id: int) -> bool:
        return bool(self.get_encrypted_token(telegram_id))

    def get_default_shop_id(self, telegram_id: int) -> int | None:
        row = self.get_user(telegram_id)
        if not row:
            return None
        return row["default_shop_id"]

    def set_default_shop_id(self, telegram_id: int, shop_id: int) -> bool:
        with self.connect() as conn:
            exists = conn.execute(
                "SELECT 1 FROM shops WHERE telegram_id = ? AND shop_id = ?",
                (telegram_id, shop_id),
            ).fetchone()
            if not exists:
                return False
            conn.execute(
                "UPDATE users SET default_shop_id = ?, updated_at = ? WHERE telegram_id = ?",
                (shop_id, now_iso(), telegram_id),
            )
            conn.commit()
            return True

    def list_shops(self, telegram_id: int) -> list[sqlite3.Row]:
        with self.connect() as conn:
            return conn.execute(
                "SELECT * FROM shops WHERE telegram_id = ? ORDER BY shop_id",
                (telegram_id,),
            ).fetchall()

    def disconnect_uzum(self, telegram_id: int) -> None:
        with self.connect() as conn:
            conn.execute(
                """
                UPDATE users
                SET uzum_token_encrypted = NULL,
                    token_connected_at = NULL,
                    default_shop_id = NULL,
                    updated_at = ?
                WHERE telegram_id = ?
                """,
                (now_iso(), telegram_id),
            )
            conn.execute("DELETE FROM shops WHERE telegram_id = ?", (telegram_id,))
            conn.commit()


class TokenCipher:
    def __init__(self, key: str) -> None:
        if not key:
            raise RuntimeError(
                "ENCRYPTION_KEY is empty. Generate it with: "
                "python -c \"from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())\""
            )
        self.fernet = Fernet(key.encode())

    def encrypt(self, token: str) -> str:
        return self.fernet.encrypt(token.encode()).decode()

    def decrypt(self, encrypted_token: str) -> str:
        return self.fernet.decrypt(encrypted_token.encode()).decode()


def extract_shop_id(shop: dict[str, Any]) -> int | None:
    for key in ("id", "shopId", "sellerId", "organizationId"):
        value = shop.get(key)
        if value is not None:
            try:
                return int(value)
            except Exception:
                pass

    for value in shop.values():
        if isinstance(value, dict):
            nested = extract_shop_id(value)
            if nested is not None:
                return nested

    return None


def extract_shop_title(shop: dict[str, Any]) -> str:
    for key in ("title", "name", "shopTitle", "organizationName", "legalName"):
        value = shop.get(key)
        if value:
            return str(value)

    for value in shop.values():
        if isinstance(value, dict):
            nested = extract_shop_title(value)
            if nested and nested != "—":
                return nested

    return "—"
