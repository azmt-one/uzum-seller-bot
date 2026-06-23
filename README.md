# Uzum Seller Telegram Bot MVP

Минимальный Telegram-бот для Uzum Seller API: магазины, товары/остатки, низкие остатки, FBS/DBS-заказы и экспорт товаров в Excel.

## 1. Что уже есть

- `/pinguzum` — проверка подключения к Uzum API
- `/shops` — список ваших магазинов
- `/products <shop_id> [поиск]` — список товаров
- `/lowstock <shop_id> [порог]` — товары с низким остатком
- `/orders <shop_id> [status]` — заказы FBS/DBS по статусу
- `/export_products <shop_id>` — Excel-выгрузка товаров

## 2. Переменные окружения

Скопируйте `.env.example` в `.env` для локального запуска или задайте эти переменные в панели BotHost:

```env
TELEGRAM_BOT_TOKEN=...
UZUM_API_TOKEN=...
OWNER_TELEGRAM_ID=
DEFAULT_SHOP_ID=
UZUM_API_BASE_URL=https://api-seller.uzum.uz/api/seller-openapi
```

Важно: `UZUM_API_TOKEN` храните только на сервере. Не публикуйте его в GitHub.

## 3. Локальный запуск

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
python main.py
```

## 4. Запуск на BotHost

1. Создайте бота в Telegram через `@BotFather`.
2. Создайте проект/бота в BotHost.
3. Загрузите файлы из этого проекта или подключите GitHub.
4. В переменных окружения укажите `TELEGRAM_BOT_TOKEN`, `UZUM_API_TOKEN`, желательно `OWNER_TELEGRAM_ID`.
5. Команда запуска: `python main.py`.
6. Для теста напишите боту `/pinguzum`, потом `/shops`.

## 5. Рекомендации для продакшена

Для одного собственного магазина можно хранить `UZUM_API_TOKEN` в переменных окружения BotHost.  
Если хотите делать сервис для многих селлеров, нужно добавить:

- регистрацию продавцов;
- шифрование API-токенов;
- PostgreSQL;
- тарифы/подписку;
- админ-панель;
- фоновые уведомления по расписанию;
- логи и мониторинг ошибок.

## 6. Следующий этап

После проверки этого MVP обычно добавляют:

- ежедневный отчет утром;
- уведомления о новых заказах;
- авто-выгрузку Excel;
- уведомления “товар заканчивается”;
- AI-анализ продаж и рекомендаций.
