from aiogram.types import KeyboardButton, ReplyKeyboardMarkup


def main_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📊 Продажи"), KeyboardButton(text="💰 Баланс")],
            [KeyboardButton(text="📦 Заказы"), KeyboardButton(text="🏬 Остатки")],
            [KeyboardButton(text="⭐ Отзывы"), KeyboardButton(text="📄 Накладная")],
            [KeyboardButton(text="⚙️ Настройки"), KeyboardButton(text="🆘 Помощь")],
        ],
        resize_keyboard=True,
        input_field_placeholder="Выберите раздел",
    )
