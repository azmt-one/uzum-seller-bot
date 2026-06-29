from dataclasses import dataclass
from datetime import datetime
from app.config import settings


@dataclass
class SalesSummary:
    orders: int
    revenue: int
    returns: int
    net_revenue: int


@dataclass
class BalanceSummary:
    available: int
    pending: int
    currency: str = "сум"


@dataclass
class OrderInfo:
    order_id: str
    status: str
    amount: int


@dataclass
class StockItem:
    sku: str
    name: str
    quantity: int


async def get_sales_summary(period: str = "today") -> SalesSummary:
    """
    Сейчас это демо-данные.
    Когда появится рабочий доступ к UZUM, здесь будут реальные запросы к кабинету/API.
    """
    if settings.uzum_token:
        # TODO: подключить реальный UZUM API/запросы кабинета.
        pass
    return SalesSummary(orders=0, revenue=0, returns=0, net_revenue=0)


async def get_balance() -> BalanceSummary:
    return BalanceSummary(available=0, pending=0)


async def get_orders() -> list[OrderInfo]:
    return [
        OrderInfo(order_id="DEMO-001", status="Ожидает подключения UZUM", amount=0),
    ]


async def get_stock() -> list[StockItem]:
    return [
        StockItem(sku="DEMO-SKU", name="Товар появится после подключения UZUM", quantity=0),
    ]


async def get_reviews() -> list[str]:
    return ["Пока демо-режим. После подключения UZUM здесь будут отзывы покупателей."]


async def create_invoice(order_id: str) -> str:
    now = datetime.now().strftime("%d.%m.%Y %H:%M")
    return f"Демо-накладная для заказа {order_id}\nСоздано: {now}\nПосле подключения UZUM бот сможет создавать настоящую накладную."
