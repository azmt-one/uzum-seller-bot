from __future__ import annotations

"""Safely installs site_bridge.py into the current Uzum Seller Assistant main.py.

Run in the bot repository root:
    python apply_site_bridge.py

The script creates main.py.before_site_bridge.py and supports both:
1) the current large single-file bot (`dp = Dispatcher()`), and
2) the loader/wrapper main.py (`namespace = _exec_original_main()`).
"""

from pathlib import Path
import shutil

MAIN = Path("main.py")
BACKUP = Path("main.py.before_site_bridge.py")
IMPORT = "from site_bridge import install_site_bridge"


def main() -> None:
    if not MAIN.exists():
        raise SystemExit("main.py не найден. Запустите скрипт в корне репозитория бота.")
    text = MAIN.read_text(encoding="utf-8")
    if "install_site_bridge(" in text:
        print("Интеграция уже установлена — изменений нет.")
        return
    if not Path("site_bridge.py").exists():
        raise SystemExit("site_bridge.py не найден рядом со скриптом.")

    shutil.copy2(MAIN, BACKUP)

    if "namespace = _exec_original_main()" in text:
        marker = "namespace = _exec_original_main()"
        replacement = marker + f"\n    {IMPORT}\n    install_site_bridge(namespace)"
        text = text.replace(marker, replacement, 1)
    elif "dp = Dispatcher()" in text:
        # `globals()` is a live dictionary. The bridge sees functions defined later,
        # including language, subscription and cost helpers.
        marker = "dp = Dispatcher()"
        replacement = marker + f"\n\n{IMPORT}\ninstall_site_bridge(globals())"
        text = text.replace(marker, replacement, 1)
    else:
        raise SystemExit("Не найден dp = Dispatcher() или loader-структура. main.py не изменён.")

    MAIN.write_text(text, encoding="utf-8")
    print("Готово: site_bridge установлен в main.py")
    print(f"Резервная копия: {BACKUP}")


if __name__ == "__main__":
    main()
