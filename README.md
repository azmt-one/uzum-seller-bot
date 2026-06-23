# Uzum Seller Telegram Bot — Multi-seller MVP fix3

Версия с точными остатками FBO + FBS/DBS по SKU из ответа `/v1/product/shop/{shopId}`.

## Как считаются остатки

По реальному ответу Uzum Seller API внутри `skuList` есть поля:

- `quantityAvailable` — общий доступный остаток;
- `quantityActive` — активный остаток в продаже;
- `quantityFbs` — остаток FBS/DBS на складе продавца;
- `quantityAdditional` — дополнительный остаток, если передан;
- `quantitySold`, `quantityReturned`, `quantityMissing`, `quantityDefected`, `quantityPending` — служебные количества.

Если отдельного поля FBO нет, бот считает:

```text
FBO = quantityAvailable - quantityFbs
Итого = quantityAvailable
FBS/DBS = quantityFbs
```

Например, если `quantityAvailable=5`, `quantityFbs=0`, значит:
- FBO: 5
- FBS/DBS: 0
- Итого: 5

## Команды

- `/connect` — подключить личный Uzum API-токен селлера
- `/disconnect` — удалить подключение
- `/status` — статус подключения
- `/pinguzum` — проверить Uzum API
- `/shops` — список магазинов селлера
- `/setshop SHOP_ID` — выбрать основной магазин
- `/products [поиск]` — товары, цены и агрегированный остаток
- `/stock [поиск]` — FBO + FBS/DBS + итого по SKU
- `/fbo [поиск]` — только FBO-остатки
- `/fbs [поиск]` — только FBS/DBS-остатки
- `/lowstock [порог]` — низкие остатки по общему количеству
- `/orders [status]` — FBS/DBS заказы
- `/export_products` — Excel с колонками FBO, FBS/DBS, итого
- `/debug_product` — сырой JSON первого товара

## Переменные окружения

```env
TELEGRAM_BOT_TOKEN=...
ENCRYPTION_KEY=...
OWNER_TELEGRAM_ID=
DB_PATH=bot.db
UZUM_API_BASE_URL=https://api-seller.uzum.uz/api/seller-openapi
```

Сгенерировать `ENCRYPTION_KEY`:

```python
import base64, os; print(base64.urlsafe_b64encode(os.urandom(32)).decode())
```

`ENCRYPTION_KEY` нельзя терять. Если он изменится, старые Uzum API-токены в базе нельзя будет расшифровать.
