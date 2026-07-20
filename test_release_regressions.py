from __future__ import annotations

import ast
import asyncio
import atexit
import os
import shutil
import sqlite3
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch

from cryptography.fernet import Fernet
from openpyxl import Workbook


TEST_ROOT = Path(tempfile.mkdtemp(prefix="sellerpro-r12-tests-"))
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
            "2026.07.20-premium-r12-uzum-finance-reminders",
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

    def test_manual_cost_is_preserved_but_never_used_for_profit(self) -> None:
        telegram_id = 9_876_543_211
        shop_id = 501
        main.save_unit_cost(telegram_id, shop_id, "SKU-1", 123_456, "Legacy")
        self.assertEqual(main.get_unit_cost_map(telegram_id, shop_id), {})

        main._replace_uzum_sku_financials(
            telegram_id,
            shop_id,
            [{"sku_id": "SKU-1", "purchase_price": 42_000}],
        )
        costs = main.get_unit_cost_map(telegram_id, shop_id)
        self.assertEqual(costs["sku-1"]["cost"], 42_000)
        self.assertEqual(costs["sku-1"]["source"], "uzum")

    def test_logistics_notification_settings_are_user_controllable(self) -> None:
        self.assertIn("notify_supply_reminders", main.AUTOMATION_BOOL_FIELDS)
        self.assertIn("notify_return_pickup", main.AUTOMATION_BOOL_FIELDS)
        self.assertIn(
            "notify_supply_reminders",
            main.NOTIFICATION_SECTION_FIELDS["stock"],
        )
        self.assertIn(
            "notify_return_pickup",
            main.NOTIFICATION_SECTION_FIELDS["stock"],
        )

    def test_logistics_outbox_key_is_idempotent(self) -> None:
        telegram_id = 9_876_543_212
        shop_id = 502
        event_key = "invoice:77:slot:2026-07-20T12:00:00+00:00:24h"
        first = main._enqueue_notification(
            "supply_reminder",
            telegram_id,
            shop_id,
            event_key,
            {"text": "test"},
        )
        second = main._enqueue_notification(
            "supply_reminder",
            telegram_id,
            shop_id,
            event_key,
            {"text": "test"},
        )
        self.assertTrue(first)
        self.assertFalse(second)

    def test_logistics_messages_are_bilingual_and_escaped(self) -> None:
        invoice = {
            "id": 77,
            "invoiceNumber": "<INV-77>",
            "status": "CREATED",
            "warehouseName": "<Main>",
            "timeSlotReservation": {
                "timeFrom": "2026-07-20T17:00:00+05:00",
                "timeTo": "2026-07-20T18:00:00+05:00",
            },
        }
        ru = main.build_supply_reminder_message(
            invoice,
            shop_id=42,
            bucket="24h",
            lang="ru",
        )
        uz = main.build_supply_reminder_message(
            invoice,
            shop_id=42,
            bucket="3h",
            lang="uz",
        )
        self.assertIn("Срок поставки", ru)
        self.assertIn("Yetkazish vaqti", uz)
        self.assertNotIn("<INV-77>", ru)
        self.assertIn("&lt;INV-77&gt;", ru)

    def test_business_profit_deducts_each_expense_once(self) -> None:
        result = main.calculate_business_profit(
            {
                "profit": 700_000,
                "cost_total": 300_000,
                "coverage": 1.0,
                "missing_count": 0,
            },
            {
                "revenue": 1_200_000,
                "commission": 120_000,
                "logistics": 80_000,
                "payout_total": 1_000_000,
            },
            {
                "tax_percent": 5,
                "advertising_monthly": 30_000,
                "storage_monthly": 20_000,
                "other_monthly": 10_000,
            },
            days=30,
            uzum_expenses={
                "available": True,
                "total": 40_000,
                "order_charge": 200_000,
            },
        )
        self.assertEqual(result["tax_expense"], 60_000)
        self.assertEqual(result["operating_expenses"], 160_000)
        self.assertEqual(result["net_profit"], 540_000)
        self.assertEqual(result["uzum_order_charge_already_in_payout"], 200_000)
        self.assertTrue(result["complete"])

    def test_no_sales_period_can_be_complete_without_fake_cost_coverage(self) -> None:
        result = main.calculate_business_profit(
            {"profit": 0, "cost_total": 0, "coverage": 0, "missing_count": 0},
            {"revenue": 0},
            {"tax_percent": 0},
            uzum_expenses={"available": True, "total": 0},
        )
        self.assertTrue(result["complete"])
        self.assertEqual(result["coverage"], 0)

    def test_hourly_digest_queue_is_persistent_and_idempotent(self) -> None:
        telegram_id = 9_876_543_213
        shop_id = 503
        now = datetime(2026, 7, 20, 12, 0, tzinfo=timezone.utc)
        item = {
            "orderId": "ORDER-77",
            "id": "SALE-77",
            "skuId": 1001,
            "productTitle": "Test product",
            "skuTitle": "SKU-1001",
            "status": "DELIVERED",
            "amount": 2,
            "sellerPrice": 100_000,
            "commission": 20_000,
            "logisticDeliveryFee": 10_000,
            "sellerProfit": 170_000,
        }
        main.set_sales_notification_mode(telegram_id, "hourly")
        main.reset_sales_digest_schedule(
            telegram_id,
            shop_id,
            now=now - timedelta(hours=2),
            clear_queue=True,
        )
        first = main.enqueue_sales_digest_events(
            telegram_id,
            shop_id,
            [item],
            detected_at=now - timedelta(minutes=10),
        )
        second = main.enqueue_sales_digest_events(
            telegram_id,
            shop_id,
            [item],
            detected_at=now - timedelta(minutes=5),
        )
        summary = main.load_sales_digest_summary(
            telegram_id,
            shop_id,
            period_start=now - timedelta(hours=2),
            period_end=now,
        )
        self.assertEqual(main.get_sales_notification_mode(telegram_id), "hourly")
        self.assertEqual(first, 1)
        self.assertEqual(second, 0)
        self.assertEqual(summary["positions"], 1)
        self.assertEqual(summary["orders"], 1)
        self.assertEqual(summary["units"], 2)
        self.assertEqual(summary["revenue"], 200_000)
        self.assertIn("Продажи за час", main.build_sales_digest_message(summary, lang="ru"))
        self.assertIn("Soatlik savdo hisoboti", main.build_sales_digest_message(summary, lang="uz"))


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

    async def test_return_pagination_deduplicates_and_stops_on_short_page(self) -> None:
        client = SimpleNamespace(
            get_returns=AsyncMock(
                side_effect=[
                    {"payload": {"returnList": [{"id": 1}, {"id": 2}]}},
                    {"payload": {"returnList": [{"id": 2}]}},
                ]
            )
        )
        rows = await main._load_return_invoices(
            client,
            42,
            max_pages=5,
            page_size=2,
        )
        self.assertEqual([row["id"] for row in rows], [1, 2])
        self.assertEqual(client.get_returns.await_count, 2)

    async def test_due_hourly_digest_is_sent_once_and_queue_is_cleared(self) -> None:
        telegram_id = 9_876_543_214
        shop_id = 504
        now = datetime(2026, 7, 20, 12, 0, tzinfo=timezone.utc)
        item = {
            "orderId": "ORDER-88",
            "id": "SALE-88",
            "skuId": 1002,
            "productTitle": "Hourly product",
            "skuTitle": "SKU-1002",
            "status": "DELIVERED",
            "amount": 1,
            "sellerPrice": 150_000,
            "sellerProfit": 120_000,
        }
        main.reset_sales_digest_schedule(
            telegram_id,
            shop_id,
            now=now - timedelta(hours=2),
            clear_queue=True,
        )
        main.enqueue_sales_digest_events(
            telegram_id,
            shop_id,
            [item],
            detected_at=now - timedelta(minutes=15),
        )
        with (
            patch.object(main.bot, "send_message", new=AsyncMock()) as sender,
            patch.object(main, "get_user_language", return_value="ru"),
        ):
            sent = await main.maybe_send_hourly_sales_digest(
                telegram_id,
                shop_id,
                now=now,
            )
            sent_again = await main.maybe_send_hourly_sales_digest(
                telegram_id,
                shop_id,
                now=now,
            )

        self.assertTrue(sent)
        self.assertFalse(sent_again)
        sender.assert_awaited_once()
        summary = main.load_sales_digest_summary(
            telegram_id,
            shop_id,
            period_start=now - timedelta(hours=1),
            period_end=now,
        )
        self.assertEqual(summary["positions"], 0)


if __name__ == "__main__":
    unittest.main()
