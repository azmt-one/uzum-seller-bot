from __future__ import annotations

from typing import Any, Iterable

import httpx


class UzumApiError(Exception):
    """Raised when Uzum Seller API returns an error response."""


class UzumClient:
    def __init__(self, token: str, base_url: str = "https://api-seller.uzum.uz/api/seller-openapi") -> None:
        if not token:
            raise ValueError("Uzum API token is empty")
        self.base_url = base_url.rstrip("/")
        self.headers = {
            "Authorization": token,
            "Accept": "application/json",
            "Content-Type": "application/json",
        }

    async def _request(self, method: str, path: str, **kwargs: Any) -> Any:
        url = f"{self.base_url}{path}"
        timeout = httpx.Timeout(30.0, connect=10.0)
        async with httpx.AsyncClient(timeout=timeout) as client:
            response = await client.request(method, url, headers=self.headers, **kwargs)
        if response.status_code >= 400:
            try:
                body = response.json()
            except Exception:
                body = response.text[:1000]
            raise UzumApiError(f"Uzum API error {response.status_code}: {body}")
        content_type = response.headers.get("content-type", "")
        if "application/json" in content_type or response.text.startswith(("{", "[")):
            return response.json()
        return response.content

    async def get_shops(self) -> Any:
        return await self._request("GET", "/v1/shops")

    async def get_products(
        self,
        shop_id: int,
        *,
        search_query: str = "",
        page: int = 0,
        size: int = 20,
        sort_by: str = "DEFAULT",
        order: str = "ASC",
        product_filter: str = "ALL",
    ) -> Any:
        params = {
            "searchQuery": search_query,
            "sortBy": sort_by,
            "order": order,
            "page": page,
            "size": size,
            "filter": product_filter,
        }
        return await self._request("GET", f"/v1/product/shop/{shop_id}", params=params)

    async def get_fbs_orders(
        self,
        shop_id: int,
        *,
        status: str = "CREATED",
        page: int = 0,
        size: int = 20,
        scheme: str | None = None,
    ) -> Any:
        params: dict[str, Any] = {
            "shopIds": shop_id,
            "status": status,
            "page": page,
            "size": size,
        }
        if scheme:
            params["scheme"] = scheme
        return await self._request("GET", "/v2/fbs/orders", params=params)

    async def count_fbs_orders(self, shop_id: int, *, status: str = "CREATED") -> Any:
        params = {"shopIds": shop_id, "status": status}
        return await self._request("GET", "/v2/fbs/orders/count", params=params)

    async def get_fbs_sku_stocks(self, *, page: int = 0, size: int = 50) -> Any:
        params = {"page": page, "size": size}
        return await self._request("GET", "/v3/fbs/sku/stocks", params=params)

    async def get_expenses(
        self,
        *,
        shop_id: int | None = None,
        shop_ids: Iterable[int] | None = None,
        date_from: int | None = None,
        date_to: int | None = None,
        sources: Iterable[str] | None = None,
        page: int = 0,
        size: int = 20,
    ) -> Any:
        params: dict[str, Any] = {"page": page, "size": size}
        resolved_shop_ids = [int(value) for value in (shop_ids or [])]
        if shop_id is not None and int(shop_id) not in resolved_shop_ids:
            resolved_shop_ids.append(int(shop_id))
        if resolved_shop_ids:
            params["shopIds"] = resolved_shop_ids
        if date_from is not None:
            params["dateFrom"] = int(date_from)
        if date_to is not None:
            params["dateTo"] = int(date_to)
        source_values = [str(value) for value in (sources or []) if str(value).strip()]
        if source_values:
            params["sources"] = source_values
        return await self._request("GET", "/v1/finance/expenses", params=params)

    async def get_fbo_invoices(
        self,
        shop_id: int,
        *,
        page: int = 0,
        size: int = 20,
    ) -> Any:
        params = {"page": int(page), "size": int(size)}
        return await self._request("GET", f"/v1/shop/{int(shop_id)}/invoice", params=params)

    async def get_fbo_invoice_products(self, shop_id: int, invoice_id: int) -> Any:
        params = {"invoiceId": int(invoice_id)}
        return await self._request(
            "GET",
            f"/v1/shop/{int(shop_id)}/invoice/products",
            params=params,
        )

    async def get_returns(
        self,
        shop_id: int,
        *,
        page: int = 0,
        size: int = 20,
    ) -> Any:
        params = {"page": int(page), "size": int(size)}
        return await self._request("GET", f"/v1/shop/{int(shop_id)}/return", params=params)

    async def get_return(self, shop_id: int, return_id: int) -> Any:
        return await self._request(
            "GET",
            f"/v1/shop/{int(shop_id)}/return/{int(return_id)}",
        )
