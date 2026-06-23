# Uzum Seller Telegram Bot — Multi-seller MVP

Версия с нормальной схемой для многих селлеров: каждый пользователь сам подключает свой Uzum Seller OpenAPI token через `/connect`.

## Что умеет

- `/connect` — подключить личный Uzum API-токен селлера
- `/disconnect` — удалить подключение
- `/status` — статус подключения
- `/pinguzum` — проверить Uzum API
- `/shops` — список магазинов селлера
- `/setshop SHOP_ID` — выбрать основной магазин
- `/products [поиск]` — товары и остатки
- `/stock [поиск]` — короткий alias для товаров/остатков
- `/lowstock [порог]` — товары с низким остатком
- `/orders [status]` — FBS/DBS заказы
- `/export_products` — выгрузка товаров в Excel

## Переменные окружения

```env
TELEGRAM_BOT_TOKEN=...
ENCRYPTION_KEY=...
OWNER_TELEGRAM_ID=
DB_PATH=bot.db
UZUM_API_BASE_URL=https://api-seller.uzum.uz/api/seller-openapi
```

Сгенерировать `ENCRYPTION_KEY`:

```bash
python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
```

`ENCRYPTION_KEY` нельзя терять. Если он изменится, старые Uzum API-токены в базе нельзя будет расшифровать.

## Локальный запуск

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
python main.py
```

## Запуск на BotHost

1. Обновите файлы в GitHub.
2. В BotHost добавьте переменные:
   - `TELEGRAM_BOT_TOKEN`
   - `ENCRYPTION_KEY`
   - `OWNER_TELEGRAM_ID`
   - `DB_PATH=bot.db`
3. Пересоберите и перезапустите бота.
4. В Telegram проверьте:
   - `/start`
   - `/connect`

## Важно для продакшена

SQLite подходит только для MVP/теста. Для коммерческого запуска лучше перейти на PostgreSQL, чтобы данные не потерялись при пересборке/миграции сервера.
