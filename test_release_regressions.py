from __future__ import annotations

import ast
import asyncio
import atexit
import os
import shutil
import sqlite3
import tempfile
import unittest
from datetime import timedelta
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch

from cryptography.fernet import Fernet
from openpyxl import Workbook


TEST_ROOT = Path(tempfile.mkdtemp(prefix="sellerpro-r11-tests-"))
atexit.register(shutil.rmtree, TEST_ROOT, True)
os.environ["TELEGRAM_BOT_TOKEN"] = f"{123456}:{'A' * 35}"
os.environ["ENCRYPTION_KEY"] = Fernet.generate_key().decode()
os.environ["DB_PATH"] = str(TEST_ROOT / "main.db")
os.environ["PDF_FONT_PATH"] = str(Path(__file__).with_name("DejaVuSans.ttf"))

import main  # noqa: E402

main.init_language_tables()


class ReleaseRegressionTests(unittest.TestCase):
    def test_release_marker_and_stock_function_are_current(self) -> None:
        self.assertEqual(
            main.PREMIUM_RELEASE_VERSION,
            "2026.07.19-premium-r11-stability-security",
        )
        self.assertTrue(callable(main.send_stock_list))

    def test_every_static_reply_button_has_one_handler(self) -> None:
        source = Path(main.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        buttons: set[str] = set()
        routes: dict[str, set[str]] = {}

        for node in ast.walk(tree):
            if not (
                isinstance(node, ast.Call)
                and isinstance(node.func, ast.Name)
                and node.func.id == "KeyboardButton"
            ):
                continue
            for keyword in node.keywords:
                if (
                    keyword.arg == "text"
                    and isinstance(keyword.value, ast.Constant)
                    and isinstance(keyword.value.value, str)
                ):
                    buttons.add(keyword.value.value)

        for node in tree.body:
            if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                continue
            for decorator in node.decorator_list:
                if not ast.unparse(decorator).startswith("dp.message("):
                    continue
                values = {
                    child.value
                    for child in ast.walk(decorator)
                    if isinstance(child, ast.Constant)
                    and isinstance(child.value, str)
                }
                for value in values:
                    if value in buttons:
                        routes.setdefault(value, set()).add(node.name)

        missing = sorted(buttons - routes.keys())
        duplicates = {
            button: sorted(handlers)
            for button, handlers in routes.items()
            if len(handlers) > 1
        }
        self.assertEqual(missing, [])
        self.assertEqual(duplicates, {})

    def test_active_stock_watchers_respect_personal_settings(self) -> None:
        low_source = __import__("inspect").getsource(main.check_low_stock_once)
        zero_source = __import__("inspect").getsource(main.check_out_of_stock_once)
        change_source = __import__("inspect").getsource(main.check_stock_change_once)
        self.assertIn('connected_watch_groups("notify_low_stock")', low_source)
        self.assertIn("low_stock_threshold", low_source)
        self.assertIn('connected_watch_groups("notify_out_of_stock")', zero_source)
        self.assertIn('connected_watch_groups("notify_stock_change")', change_source)

    def test_trial_command_adds_days_instead_of_shortening(self) -> None:
        telegram_id = 9_876_543_210
        main.ensure_subscription(telegram_id)
        before_row = main.get_subscription_row(telegram_id) or {}
        before = main._dt_from_db(before_row.get("trial_until"))
        self.assertIsNotNone(before)
        after = main.set_trial_days(telegram_id, 2)
        assert before is not None
        self.assertGreaterEqual(after - before, timedelta(days=2))

    def test_sqlite_backup_contains_uncheckpointed_wal_data(self) -> None:
        source = TEST_ROOT / "wal-source.db"
        destination = TEST_ROOT / "wal-backup.db"
        conn = sqlite3.connect(source)
        try:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA wal_autocheckpoint=0")
            conn.execute("CREATE TABLE important_data (value TEXT NOT NULL)")
            conn.execute("INSERT INTO important_data VALUES ('present')")
            conn.commit()

            main.create_sqlite_backup(source, destination)
            with sqlite3.connect(destination) as backup:
                value = backup.execute("SELECT value FROM important_data").fetchone()
                integrity = backup.execute("PRAGMA integrity_check").fetchone()
            self.assertEqual(value, ("present",))
            self.assertEqual(integrity, ("ok",))
        finally:
            conn.close()

    def test_cost_import_is_bounded_by_row_limit(self) -> None:
        workbook_path = TEST_ROOT / "costs.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["SKU", "Себестоимость"])
        for index in range(5):
            sheet.append([f"sku-{index}", 1000 + index])
        workbook.save(workbook_path)
        workbook.close()

        previous_limit = main.COST_IMPORT_MAX_ROWS
        main.COST_IMPORT_MAX_ROWS = 3
        try:
            items, errors = main._parse_costs_workbook(str(workbook_path))
        finally:
            main.COST_IMPORT_MAX_ROWS = previous_limit
        self.assertEqual(len(items), 3)
        self.assertTrue(any("лимит" in error.lower() for error in errors))


class StockRouteAsyncTests(unittest.IsolatedAsyncioTestCase):
    async def test_stock_route_loads_and_paginates_rows(self) -> None:
        message = SimpleNamespace(from_user=SimpleNamespace(id=765_432_101))
        rows = [
            {
                "sku_id": 101,
                "sku_full_title": "Тестовый товар",
                "fbo": 4,
                "fbs": 2,
                "total": 6,
                "price": 10_000,
            }
        ]
        with (
            patch.object(
                main,
                "require_connection",
                new=AsyncMock(return_value=(765_432_101, object(), 42)),
            ),
            patch.object(main, "load_sku_rows", new=AsyncMock(return_value=rows)),
            patch.object(main, "send_paginated_list", new=AsyncMock()) as sender,
        ):
            await main.send_stock_list(message, mode="all")

        sender.assert_awaited_once()
        kwargs = sender.await_args.kwargs
        self.assertEqual(kwargs["kind"], "stock_all")
        self.assertEqual(len(kwargs["items"]), 1)
        self.assertIn("Тестовый товар", kwargs["items"][0])

    async def test_duplicate_in_flight_request_is_rejected(self) -> None:
        middleware = main.UserConcurrencyMiddleware()
        event = SimpleNamespace(from_user=SimpleNamespace(id=765_432_102))
        started = asyncio.Event()
        release = asyncio.Event()

        async def slow_handler(_event, _data):
            started.set()
            await release.wait()
            return "done"

        first = asyncio.create_task(middleware(slow_handler, event, {}))
        await started.wait()
        duplicate = AsyncMock(return_value="unexpected")
        try:
            result = await middleware(duplicate, event, {})
            self.assertIsNone(result)
            duplicate.assert_not_awaited()
        finally:
            release.set()
        self.assertEqual(await first, "done")
        self.assertNotIn(event.from_user.id, main._ACTIVE_USER_HANDLERS)


if __name__ == "__main__":
    unittest.main()
