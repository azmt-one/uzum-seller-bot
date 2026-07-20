from __future__ import annotations

import unittest
from datetime import datetime, timedelta, timezone

from uzum_finance import (
    build_stock_records,
    expense_items,
    parse_api_datetime,
    return_reminder_bucket,
    stock_record,
    summarize_expenses,
    supply_reminder_bucket,
)


class UzumCostSourceTests(unittest.TestCase):
    def test_sale_price_is_never_used_as_cost(self) -> None:
        record = stock_record(
            {
                "sku_id": 101,
                "sku_title": "Test SKU",
                "price": 99_000,
                "raw": {"price": 99_000, "sellPrice": 98_000},
            }
        )
        self.assertIsNone(record["purchase_price"])

    def test_purchase_price_and_catalog_fields_come_from_uzum(self) -> None:
        record = stock_record(
            {
                "sku_id": 101,
                "seller_item_code": "ART-1",
                "barcode": "478000000001",
                "raw": {
                    "purchasePrice": 42_500,
                    "ikpu": "12345678901234567",
                    "paidStorageAmount": 1_250,
                    "paidStoragePriceItem": 250,
                    "pstorage": True,
                },
            }
        )
        self.assertEqual(record["purchase_price"], 42_500)
        self.assertEqual(record["ikpu"], "12345678901234567")
        self.assertEqual(record["paid_storage_amount"], 1_250)
        self.assertTrue(record["paid_storage"])
        self.assertIn("101", record["aliases"])
        self.assertIn("art-1", record["aliases"])

    def test_duplicate_sku_does_not_add_purchase_prices(self) -> None:
        rows = build_stock_records(
            [
                {"sku_id": 101, "purchase_price": 10_000},
                {"sku_id": 101, "purchase_price": 10_000, "barcode": "ABC"},
            ]
        )
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["purchase_price"], 10_000)
        self.assertIn("abc", rows[0]["aliases"])

    def test_zero_purchase_price_is_missing(self) -> None:
        self.assertIsNone(stock_record({"sku_id": 101, "purchase_price": 0})["purchase_price"])


class UzumExpenseLedgerTests(unittest.TestCase):
    def test_nested_payload_is_extracted(self) -> None:
        rows = expense_items({"payload": {"items": [{"id": 1}, {"id": 2}]}})
        self.assertEqual([row["id"] for row in rows], [1, 2])

    def test_only_confirmed_additional_expenses_reduce_profit(self) -> None:
        summary = summarize_expenses(
            [
                {
                    "id": 1,
                    "status": "COMPLETED",
                    "type": "EXPENSE",
                    "paymentPrice": 100,
                    "name": "Paid storage",
                },
                {
                    "id": 2,
                    "status": {"value": "SUCCESS"},
                    "type": "EXPENSE",
                    "paymentPrice": 200,
                    "name": "Advertising campaign",
                },
                {
                    "id": 3,
                    "status": "COMPLETED",
                    "type": "INCOME",
                    "paymentPrice": 50,
                    "name": "Advertising refund",
                },
                {
                    "id": 4,
                    "status": "COMPLETED",
                    "type": "EXPENSE",
                    "paymentPrice": 30,
                    "source": "ORDER_COMMISSION",
                    "name": "Order commission",
                },
                {
                    "id": 5,
                    "status": "PENDING",
                    "type": "EXPENSE",
                    "paymentPrice": 999,
                    "name": "Penalty",
                },
                {
                    "id": 6,
                    "status": "UNKNOWN_NEW_STATUS",
                    "type": "EXPENSE",
                    "paymentPrice": 777,
                    "name": "Other",
                },
            ]
        )
        self.assertEqual(summary["storage"], 100)
        self.assertEqual(summary["advertising"], 150)
        self.assertEqual(summary["order_charge"], 30)
        self.assertEqual(summary["total"], 250)
        self.assertEqual(summary["booked_count"], 4)
        self.assertEqual(summary["pending_count"], 2)
        commission = next(row for row in summary["rows"] if row["identity"] == "4")
        self.assertFalse(commission["included_in_profit"])


class LogisticsReminderTests(unittest.TestCase):
    def setUp(self) -> None:
        self.now = datetime(2026, 7, 20, 12, 0, tzinfo=timezone.utc)

    def test_supply_buckets(self) -> None:
        self.assertEqual(supply_reminder_bucket(self.now + timedelta(hours=24), self.now), "24h")
        self.assertEqual(supply_reminder_bucket(self.now + timedelta(hours=3), self.now), "3h")
        self.assertIsNone(supply_reminder_bucket(self.now + timedelta(hours=25), self.now))
        self.assertIsNone(supply_reminder_bucket(self.now - timedelta(seconds=1), self.now))

    def test_return_buckets(self) -> None:
        self.assertEqual(return_reminder_bucket(self.now + timedelta(hours=72), self.now), "3d")
        self.assertEqual(return_reminder_bucket(self.now + timedelta(hours=48), self.now), "2d")
        self.assertEqual(return_reminder_bucket(self.now + timedelta(hours=24), self.now), "1d")
        self.assertEqual(
            return_reminder_bucket(
                self.now - timedelta(minutes=1),
                self.now,
                storage_status="ACTIVE",
            ),
            "active",
        )
        self.assertIsNone(
            return_reminder_bucket(
                self.now - timedelta(minutes=1),
                self.now,
                storage_status="COMPLETED",
            )
        )

    def test_api_datetime_is_normalized_to_utc(self) -> None:
        parsed = parse_api_datetime("2026-07-20T17:00:00+05:00")
        self.assertEqual(parsed, self.now)


if __name__ == "__main__":
    unittest.main()
