# Обновление Seller.pro.uz до r11 на BotHost

## Перед загрузкой

1. Не удаляйте постоянный каталог `/app/data` и текущую базу `bot.db`.
2. Не меняйте существующие `TELEGRAM_BOT_TOKEN`, `ENCRYPTION_KEY`, `ADMIN_IDS` и `DB_PATH`.
3. Если BotHost позволяет скачать постоянный каталог, сначала остановите бот и сохраните весь `/app/data` целиком (включая возможные файлы `bot.db-wal` и `bot.db-shm`).

## Загрузка

1. Распакуйте релизный ZIP или загрузите его содержимое целиком.
2. Убедитесь, что вместе с кодом загружены:
   - `seller_pdf_report.py`;
   - `DejaVuSans.ttf`;
   - `DejaVuSans-Bold.ttf`;
   - `DejaVu-fonts-LICENSE.txt`.
3. Сохраните переменные окружения и перезапустите контейнер.

Рекомендуемые значения:

```env
DB_PATH=/app/data/bot.db
LOSS_REPORT_FILTERS=ALL,ARCHIVE,DEFECTED
LOSS_REPORT_MAX_REQUESTS=60
NOTIFICATION_MAX_ATTEMPTS=10
MAX_CONCURRENT_USER_HANDLERS=40
DROP_PENDING_UPDATES=0
```

## Проверка после запуска

В журнале должны появиться строки:

```text
PREMIUM_RELEASE_LOADED version=2026.07.19-premium-r11-stability-security
STABILITY_SECURITY_LOADED: stock routes + watcher settings + safe backup + bounded Excel import
PDF_FONT_READY: regular=/app/DejaVuSans.ttf bold=/app/DejaVuSans-Bold.ttf
```

Затем проверьте в Telegram:

1. **📦 Склад → 📦 Все остатки** — должен открыться список SKU со страницами.
2. **⚙️ Настройки → 💰 Себестоимость и расходы**.
3. **⚙️ Настройки → 🔐 API и подключение**.
4. `/stock`, `/fbo`, `/fbs` — без ошибки `send_stock_list is not defined`.
5. `/backup_db` от администратора — бот должен прислать целостный файл базы.
6. `/trial TELEGRAM_ID 2` — к текущему сроку пользователя должны добавиться два дня.

На узбекском аналогично проверьте **📦 Barcha qoldiq**, **💰 Tannarx va xarajat** и **🔐 API va ulanish**.
