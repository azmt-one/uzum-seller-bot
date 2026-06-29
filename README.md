# UZUM Seller Telegram Bot — первая рабочая версия

Это стартовый проект бота для Telegram. Он уже умеет:

- запускаться 24/7 на сервере;
- показывать меню с кнопками;
- пускать только разрешенных пользователей по Telegram ID;
- хранить данные в постоянной SQLite-базе;
- переживать перезагрузку, если база лежит в `/data`;
- иметь готовые разделы: продажи, баланс, заказы, остатки, отзывы, накладная.

Пока данные UZUM стоят в демо-режиме. Следующий шаг — подключить реальные запросы к кабинету/API UZUM.

---

## 1. Что нужно подготовить

1. Токен бота от `@BotFather`.
2. Твой Telegram ID. Узнать можно командой `/id` в боте после запуска.
3. Telegram ID жены, если ей тоже нужен доступ. Уже можно указать `938965878`.
4. Сервер bothost.ru.

---

## 2. Настройка `.env`

Скопируй файл:

```bash
cp .env.example .env
```

Открой `.env` и заполни:

```env
TELEGRAM_TOKEN=токен_от_BotFather
ADMIN_IDS=твой_id,938965878
BOT_DB_PATH=/data/uzum_bot.db
UZUM_TOKEN=
```

Важно: на bothost.ru лучше добавлять эти значения в **Environment / Переменные окружения**, а не писать токен прямо в коде.

---

## 3. Запуск на компьютере через Docker

```bash
docker compose up -d --build
```

Посмотреть логи:

```bash
docker compose logs -f bot
```

Остановить:

```bash
docker compose down
```

---

## 4. Запуск без Docker

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m app.main
```

На Linux/Mac:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m app.main
```

---

## 5. Как загрузить на bothost.ru

1. Загрузи папку проекта или подключи GitHub-репозиторий.
2. Укажи запуск через Dockerfile, если bothost это поддерживает.
3. В переменные окружения добавь:

```env
TELEGRAM_TOKEN=токен_от_BotFather
ADMIN_IDS=твой_id,938965878
BOT_DB_PATH=/data/uzum_bot.db
UZUM_TOKEN=
```

4. Обязательно сделай постоянную папку/volume `/data`, чтобы база не исчезала после обновления.
5. Запусти проект.
6. Напиши боту `/start`.

---

## 6. Команды бота

```text
/start — открыть меню
/id — узнать свой Telegram ID
/sales — продажи
/balance — баланс
/orders — заказы
/stock — остатки
/reviews — отзывы
/invoice DEMO-001 — демо-накладная
```

---

## 7. Что делать дальше

Когда бот запустится и будет отвечать, нужно подключить реальные данные UZUM.

Для этого нужно понять, какой доступ есть к UZUM seller-кабинету:

- официальный API;
- токен/ключ API;
- или запросы, которые делает личный кабинет.

Подключение делается в файле:

```text
app/services/uzum_client.py
```

Там сейчас демо-данные. Мы будем постепенно заменять их на настоящие запросы.
