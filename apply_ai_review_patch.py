"""
Авто-патчер main.py для добавления AI-ответов на отзывы.

Что делает:
1. Проверяет наличие main.py в текущей папке.
2. Делает backup: main_before_ai_review.py
3. Добавляет импорт:
       from review_ai_integration import setup_review_ai_handlers
4. Добавляет регистрацию handlers:
       setup_review_ai_handlers(...)
5. Не удаляет старые функции.

Запуск:
    python apply_ai_review_patch.py
"""

from __future__ import annotations

from pathlib import Path
import re
import sys


MAIN_FILE = Path("main.py")
BACKUP_FILE = Path("main_before_ai_review.py")

IMPORT_LINE = "from review_ai_integration import setup_review_ai_handlers\n"

SETUP_BLOCK = """
# --- AI-ответы на отзывы покупателей ---
setup_review_ai_handlers(
    dp,
    menu_for_message=menu_for_message if "menu_for_message" in globals() else None,
    get_user_language=get_user_language if "get_user_language" in globals() else None,
    require_active_subscription=require_active_subscription if "require_active_subscription" in globals() else None,
    upsert_from_message=upsert_from_message if "upsert_from_message" in globals() else None,
)
"""


def fail(text: str) -> None:
    print(f"❌ {text}")
    sys.exit(1)


def main() -> None:
    if not MAIN_FILE.exists():
        fail("Файл main.py не найден. Положите этот патчер в ту же папку, где находится main.py.")

    text = MAIN_FILE.read_text(encoding="utf-8")

    if "setup_review_ai_handlers(" in text:
        print("✅ AI-ответы уже подключены в main.py. Повторно менять файл не нужно.")
        return

    if not BACKUP_FILE.exists():
        BACKUP_FILE.write_text(text, encoding="utf-8")
        print(f"✅ Создан backup: {BACKUP_FILE}")
    else:
        print(f"ℹ️ Backup уже есть: {BACKUP_FILE}")

    # 1) Добавляем импорт после project imports.
    if IMPORT_LINE not in text:
        if "from uzum_client import UzumClient\n" in text:
            text = text.replace(
                "from uzum_client import UzumClient\n",
                "from uzum_client import UzumClient\n" + IMPORT_LINE,
                1,
            )
        else:
            lines = text.splitlines(keepends=True)
            insert_at = 0
            for i, line in enumerate(lines[:100]):
                if line.startswith("import ") or line.startswith("from "):
                    insert_at = i + 1
            lines.insert(insert_at, IMPORT_LINE)
            text = "".join(lines)

    # 2) Вставляем setup перед async def main(), если оно есть.
    if re.search(r"\nasync\s+def\s+main\s*\(", text):
        text = re.sub(
            r"\nasync\s+def\s+main\s*\(",
            "\n" + SETUP_BLOCK + "\nasync def main(",
            text,
            count=1,
        )
    elif '\nif __name__ == "__main__":' in text:
        text = text.replace(
            '\nif __name__ == "__main__":',
            "\n" + SETUP_BLOCK + '\nif __name__ == "__main__":',
            1,
        )
    else:
        text = text.rstrip() + "\n\n" + SETUP_BLOCK + "\n"

    MAIN_FILE.write_text(text, encoding="utf-8")
    print("✅ Готово: AI-ответы добавлены в main.py")
    print("Теперь добавьте OPENAI_API_KEY в BotHost и openai>=1.93.0 в requirements.txt")


if __name__ == "__main__":
    main()
