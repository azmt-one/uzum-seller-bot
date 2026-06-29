import os
import aiosqlite
from app.config import settings


async def get_db() -> aiosqlite.Connection:
    db_path = settings.bot_db_path
    db_dir = os.path.dirname(db_path)
    if db_dir:
        os.makedirs(db_dir, exist_ok=True)

    db = await aiosqlite.connect(db_path)
    db.row_factory = aiosqlite.Row
    return db


async def init_db() -> None:
    async with await get_db() as db:
        await db.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                telegram_id INTEGER PRIMARY KEY,
                first_name TEXT,
                username TEXT,
                is_allowed INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        await db.execute(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        await db.execute(
            """
            CREATE TABLE IF NOT EXISTS command_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                telegram_id INTEGER,
                command TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        await db.commit()


async def save_user(telegram_id: int, first_name: str | None, username: str | None, is_allowed: bool) -> None:
    async with await get_db() as db:
        await db.execute(
            """
            INSERT INTO users (telegram_id, first_name, username, is_allowed)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(telegram_id) DO UPDATE SET
                first_name = excluded.first_name,
                username = excluded.username,
                is_allowed = excluded.is_allowed
            """,
            (telegram_id, first_name, username, int(is_allowed)),
        )
        await db.commit()


async def log_command(telegram_id: int, command: str) -> None:
    async with await get_db() as db:
        await db.execute(
            "INSERT INTO command_logs (telegram_id, command) VALUES (?, ?)",
            (telegram_id, command),
        )
        await db.commit()
