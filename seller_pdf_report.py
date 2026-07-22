from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from reportlab.lib.colors import Color, HexColor, white
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph


W, H = A4
MARGIN = 38
TOTAL_PAGES = 4

PURPLE = HexColor("#6C3CEB")
PURPLE_DARK = HexColor("#4721A8")
PURPLE_SOFT = HexColor("#F1EDFF")
INK = HexColor("#1F2937")
MUTED = HexColor("#667085")
LINE = HexColor("#E7E8EE")
BG = HexColor("#F7F8FC")
GREEN = HexColor("#16A36A")
GREEN_SOFT = HexColor("#EAF8F2")
RED = HexColor("#E55252")
RED_SOFT = HexColor("#FDEEEE")
ORANGE = HexColor("#F59E0B")
ORANGE_SOFT = HexColor("#FFF6E2")
BLUE = HexColor("#2878E8")
BLUE_SOFT = HexColor("#EAF2FE")

FONT_REGULAR = "SellerPDFRegular"
FONT_BOLD = "SellerPDFBold"


def _t(lang: str, ru: str, uz: str) -> str:
    return uz if str(lang).lower().startswith("uz") else ru


def _num(value: Any) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _money(value: Any, lang: str) -> str:
    suffix = " so‘m" if str(lang).lower().startswith("uz") else " сум"
    return f"{int(round(_num(value))):,}".replace(",", " ") + suffix


def _qty(value: Any) -> str:
    number = _num(value)
    return f"{number:.0f}" if abs(number - round(number)) < 0.001 else f"{number:.1f}"


def _percent(value: Any) -> str:
    return f"{_num(value) * 100:.1f}%"


def _change(current: Any, previous: Any) -> float | None:
    current_value = _num(current)
    previous_value = _num(previous)
    if abs(previous_value) < 0.000001:
        return None if abs(current_value) < 0.000001 else 1.0
    return (current_value - previous_value) / abs(previous_value)


def _change_text(value: float | None, lang: str) -> str:
    if value is None:
        return _t(lang, "без сравнения", "taqqoslash yo‘q")
    prefix = "+" if value >= 0 else ""
    return f"{prefix}{value * 100:.1f}%"


def _short(value: Any, limit: int) -> str:
    text = " ".join(str(value or "-").split())
    if len(text) <= limit:
        return text
    return text[: max(1, limit - 3)].rstrip() + "..."


def _register_fonts(regular_path: str | Path, bold_path: str | Path) -> None:
    if FONT_REGULAR not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(FONT_REGULAR, str(regular_path)))
    if FONT_BOLD not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(FONT_BOLD, str(bold_path)))


def _text_width(text: str, font: str = FONT_REGULAR, size: float = 9) -> float:
    return pdfmetrics.stringWidth(str(text), font, size)


def _rounded(c: canvas.Canvas, x: float, y: float, w: float, h: float, fill, radius: float = 12) -> None:
    c.setFillColor(fill)
    c.setStrokeColor(fill)
    c.roundRect(x, y, w, h, radius, fill=1, stroke=0)


def _paragraph(
    c: canvas.Canvas,
    text: str,
    x: float,
    y_top: float,
    width: float,
    *,
    size: float = 8,
    leading: float | None = None,
    color=INK,
    bold: bool = False,
    max_height: float = 200,
) -> float:
    style = ParagraphStyle(
        "seller_pdf",
        fontName=FONT_BOLD if bold else FONT_REGULAR,
        fontSize=size,
        leading=leading or size * 1.35,
        textColor=color,
        alignment=TA_LEFT,
        spaceAfter=0,
        spaceBefore=0,
    )
    paragraph = Paragraph(str(text), style)
    _, height = paragraph.wrap(width, max_height)
    paragraph.drawOn(c, x, y_top - height)
    return height


def _pill(c: canvas.Canvas, text: str, x: float, y: float, fill, color=white, size: float = 7.5) -> float:
    width = _text_width(text, FONT_BOLD, size) + 18
    _rounded(c, x, y, width, 21, fill, 10.5)
    c.setFillColor(color)
    c.setFont(FONT_BOLD, size)
    c.drawCentredString(x + width / 2, y + 7, text)
    return width


def _header(c: canvas.Canvas, page: int, title: str, subtitle: str, lang: str) -> None:
    c.setFillColor(PURPLE_DARK)
    c.rect(0, H - 104, W, 104, fill=1, stroke=0)
    c.setFillColor(Color(1, 1, 1, alpha=0.08))
    c.circle(W - 42, H - 18, 78, fill=1, stroke=0)
    c.circle(W - 5, H - 85, 52, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(FONT_BOLD, 16)
    c.drawString(MARGIN, H - 34, "Seller.pro.uz")
    c.setFont(FONT_BOLD, 8.5)
    c.drawRightString(
        W - MARGIN,
        H - 36,
        _t(lang, f"СТРАНИЦА {page}/{TOTAL_PAGES}", f"SAHIFA {page}/{TOTAL_PAGES}"),
    )
    c.setFont(FONT_BOLD, 17)
    c.drawString(MARGIN, H - 70, title)
    c.setFillColor(HexColor("#DED5FF"))
    c.setFont(FONT_REGULAR, 8.2)
    c.drawString(MARGIN, H - 89, _short(subtitle, 105))


def _footer(c: canvas.Canvas, page: int, lang: str) -> None:
    c.setStrokeColor(LINE)
    c.setLineWidth(0.7)
    c.line(MARGIN, 30, W - MARGIN, 30)
    c.setFillColor(MUTED)
    c.setFont(FONT_REGULAR, 7)
    c.drawString(
        MARGIN,
        17,
        _t(lang, "Отчёт сформирован по данным Uzum API и настройкам пользователя.", "Hisobot Uzum API va foydalanuvchi sozlamalari asosida tuzildi."),
    )
    c.drawRightString(W - MARGIN, 17, f"Seller.pro.uz  |  {page}")


def _section(c: canvas.Canvas, title: str, subtitle: str, x: float, y: float) -> None:
    c.setFillColor(INK)
    c.setFont(FONT_BOLD, 13)
    c.drawString(x, y, title)
    c.setFillColor(MUTED)
    c.setFont(FONT_REGULAR, 7.8)
    c.drawString(x, y - 15, _short(subtitle, 112))


def _kpi(
    c: canvas.Canvas,
    x: float,
    y: float,
    w: float,
    h: float,
    label: str,
    value: str,
    note: str,
    accent,
    soft,
) -> None:
    _rounded(c, x, y, w, h, white, 12)
    c.setStrokeColor(LINE)
    c.setLineWidth(0.7)
    c.roundRect(x, y, w, h, 12, fill=0, stroke=1)
    _rounded(c, x + 12, y + h - 25, 7, 7, accent, 3.5)
    c.setFillColor(MUTED)
    c.setFont(FONT_REGULAR, 7.7)
    c.drawString(x + 26, y + h - 25, _short(label, 27))
    c.setFillColor(INK)
    c.setFont(FONT_BOLD, 13.2)
    c.drawString(x + 12, y + 28, _short(value, 24))
    tag_width = min(w - 24, _text_width(note, FONT_BOLD, 6.7) + 14)
    _rounded(c, x + 12, y + 9, tag_width, 15, soft, 7)
    c.setFillColor(accent)
    c.setFont(FONT_BOLD, 6.7)
    c.drawCentredString(x + 12 + tag_width / 2, y + 14, _short(note, 30))


def _line_chart(c: canvas.Canvas, x: float, y: float, w: float, h: float, rows: list[dict[str, Any]]) -> None:
    values = [_num(row.get("revenue")) for row in rows]
    labels = [str(row.get("label") or "") for row in rows]
    if not values:
        values = [0.0, 0.0]
        labels = ["", ""]
    if len(values) == 1:
        values.append(values[0])
        labels.append(labels[0])
    max_value = max(values) * 1.12 or 1.0
    plot_x, plot_y = x + 36, y + 25
    plot_w, plot_h = w - 50, h - 42
    c.setStrokeColor(LINE)
    c.setLineWidth(0.5)
    for index in range(4):
        yy = plot_y + plot_h * index / 3
        c.line(plot_x, yy, plot_x + plot_w, yy)
        c.setFillColor(MUTED)
        c.setFont(FONT_REGULAR, 6)
        c.drawRightString(plot_x - 6, yy - 2, f"{max_value * index / 3 / 1000:.0f}k")
    points: list[tuple[float, float]] = []
    for index, value in enumerate(values):
        px = plot_x + plot_w * index / max(1, len(values) - 1)
        py = plot_y + value / max_value * plot_h
        points.append((px, py))
    area = c.beginPath()
    area.moveTo(points[0][0], plot_y)
    for px, py in points:
        area.lineTo(px, py)
    area.lineTo(points[-1][0], plot_y)
    area.close()
    c.saveState()
    c.setFillColor(Color(108 / 255, 60 / 255, 235 / 255, alpha=0.11))
    c.drawPath(area, fill=1, stroke=0)
    c.restoreState()
    path = c.beginPath()
    path.moveTo(*points[0])
    for point in points[1:]:
        path.lineTo(*point)
    c.setStrokeColor(PURPLE)
    c.setLineWidth(2)
    c.drawPath(path, fill=0, stroke=1)
    for index, (px, py) in enumerate(points):
        c.setFillColor(white)
        c.setStrokeColor(PURPLE)
        c.circle(px, py, 2.7, fill=1, stroke=1)
        step = max(1, len(points) // 7)
        if index % step == 0 or index == len(points) - 1:
            c.setFillColor(MUTED)
            c.setFont(FONT_REGULAR, 5.8)
            c.drawCentredString(px, plot_y - 13, labels[index])


def _bar(c: canvas.Canvas, x: float, y: float, w: float, value: float, maximum: float, color, label: str, amount: str) -> None:
    c.setFillColor(INK)
    c.setFont(FONT_REGULAR, 7.7)
    c.drawString(x, y + 13, _short(label, 36))
    c.setFont(FONT_BOLD, 7.7)
    c.drawRightString(x + w, y + 13, amount)
    _rounded(c, x, y, w, 7, HexColor("#EEF0F5"), 3.5)
    fill_width = max(7, w * max(0.0, value) / max(1.0, maximum))
    _rounded(c, x, y, min(w, fill_width), 7, color, 3.5)


def _table(
    c: canvas.Canvas,
    x: float,
    y_top: float,
    widths: list[float],
    headers: list[str],
    rows: list[list[str]],
    *,
    row_h: float = 27,
    right: set[int] | None = None,
    last_colors: list[Any] | None = None,
) -> float:
    right = right or set()
    total_width = sum(widths)
    _rounded(c, x, y_top - row_h, total_width, row_h, PURPLE_SOFT, 8)
    cursor = x
    for index, (header, width) in enumerate(zip(headers, widths)):
        c.setFillColor(PURPLE_DARK)
        c.setFont(FONT_BOLD, 7)
        if index in right:
            c.drawRightString(cursor + width - 8, y_top - 17, header)
        else:
            c.drawString(cursor + 8, y_top - 17, header)
        cursor += width
    y = y_top - row_h
    for row_index, row in enumerate(rows):
        y -= row_h
        if row_index % 2:
            c.setFillColor(BG)
            c.rect(x, y, total_width, row_h, fill=1, stroke=0)
        c.setStrokeColor(LINE)
        c.setLineWidth(0.35)
        c.line(x, y, x + total_width, y)
        cursor = x
        for index, (value, width) in enumerate(zip(row, widths)):
            color = last_colors[row_index] if last_colors and index == len(row) - 1 and row_index < len(last_colors) else INK
            c.setFillColor(color)
            font = FONT_BOLD if index == 0 else FONT_REGULAR
            c.setFont(font, 6.9)
            shown = _short(value, 70)
            while _text_width(shown, font, 6.9) > width - 16 and len(shown) > 5:
                shown = shown[:-4].rstrip() + "..."
            if index in right:
                c.drawRightString(cursor + width - 8, y + 9.5, shown)
            else:
                c.drawString(cursor + 8, y + 9.5, shown)
            cursor += width
    return y


def _priority_card(c: canvas.Canvas, x: float, y: float, w: float, h: float, number: int, title: str, body: str, accent, soft) -> None:
    _rounded(c, x, y, w, h, white, 12)
    c.setStrokeColor(LINE)
    c.roundRect(x, y, w, h, 12, fill=0, stroke=1)
    _rounded(c, x + 12, y + h - 34, 24, 24, soft, 8)
    c.setFillColor(accent)
    c.setFont(FONT_BOLD, 10)
    c.drawCentredString(x + 24, y + h - 26, str(number))
    c.setFillColor(INK)
    c.setFont(FONT_BOLD, 8.5)
    c.drawString(x + 45, y + h - 25, _short(title, 34))
    _paragraph(c, _short(body, 155), x + 12, y + h - 43, w - 24, size=7, leading=9.5, color=MUTED, max_height=h - 49)


def _period_subtitle(payload: dict[str, Any], lang: str) -> str:
    generated = payload.get("generated_at")
    if isinstance(generated, datetime):
        generated_text = generated.strftime("%d.%m.%Y %H:%M")
    else:
        generated_text = str(generated or "-")
    return _t(
        lang,
        f"Период: {payload.get('period_label') or '-'}  |  Магазин: {payload.get('shop_id') or '-'}  |  Сформирован: {generated_text}",
        f"Davr: {payload.get('period_label') or '-'}  |  Do‘kon: {payload.get('shop_id') or '-'}  |  Yaratildi: {generated_text}",
    )


def _page_summary(c: canvas.Canvas, payload: dict[str, Any], lang: str) -> None:
    stats = dict(payload.get("stats") or {})
    previous = dict(payload.get("previous_stats") or {})
    profit = dict(payload.get("profit") or {})
    business = dict(payload.get("business_profit") or {})
    coverage = _num(business.get("coverage", profit.get("coverage")))
    business_complete = bool(business.get("complete"))
    comparison_available = bool(payload.get("comparison_available", True))
    revenue_change = _change(stats.get("revenue"), previous.get("revenue")) if comparison_available else None
    orders_change = _change(stats.get("orders"), previous.get("orders")) if comparison_available else None
    title = _t(lang, "Управленческий отчёт магазина", "Do‘kon boshqaruv hisoboti")
    _header(c, 1, title, _period_subtitle(payload, lang), lang)
    _section(c, _t(lang, "Главное за период", "Davrning asosiy ko‘rsatkichlari"), _t(lang, "Ключевые показатели и динамика относительно предыдущего такого же периода", "Asosiy ko‘rsatkichlar va oldingi shu davrga nisbatan dinamika"), MARGIN, H - 135)
    gap = 10
    card_w = (W - 2 * MARGIN - 2 * gap) / 3
    y1, y2 = H - 240, H - 320
    revenue_accent = RED if revenue_change is not None and revenue_change < 0 else GREEN
    revenue_soft = RED_SOFT if revenue_accent == RED else GREEN_SOFT
    orders_accent = RED if orders_change is not None and orders_change < 0 else PURPLE
    orders_soft = RED_SOFT if orders_accent == RED else PURPLE_SOFT
    _kpi(c, MARGIN, y1, card_w, 70, _t(lang, "Выручка", "Tushum"), _money(stats.get("revenue"), lang), _change_text(revenue_change, lang), revenue_accent, revenue_soft)
    _kpi(c, MARGIN + card_w + gap, y1, card_w, 70, _t(lang, "Заказы", "Buyurtmalar"), str(int(_num(stats.get("orders")))), _change_text(orders_change, lang), orders_accent, orders_soft)
    _kpi(c, MARGIN + 2 * (card_w + gap), y1, card_w, 70, _t(lang, "Продано", "Sotildi"), f"{_qty(stats.get('units'))} {_t(lang, 'шт.', 'dona')}", _t(lang, "чистое количество", "sof miqdor"), BLUE, BLUE_SOFT)
    _kpi(c, MARGIN, y2, card_w, 70, _t(lang, "К выплате", "To‘lovga"), _money(stats.get("payout_total"), lang), _percent(_num(stats.get("payout_total")) / max(1.0, _num(stats.get("revenue")))), GREEN, GREEN_SOFT)
    result_value = _num(business.get("net_profit", profit.get("profit")))
    profit_label = _t(
        lang,
        "Чистая прибыль" if business_complete else "Известный результат",
        "Sof foyda" if business_complete else "Ma’lum natija",
    )
    profit_accent = RED if result_value < 0 else PURPLE
    profit_soft = RED_SOFT if result_value < 0 else PURPLE_SOFT
    _kpi(c, MARGIN + card_w + gap, y2, card_w, 70, profit_label, _money(result_value, lang), _t(lang, f"себестоимость {coverage * 100:.1f}%", f"tannarx {coverage * 100:.1f}%"), profit_accent, profit_soft)
    _kpi(c, MARGIN + 2 * (card_w + gap), y2, card_w, 70, _t(lang, "Отмены", "Bekor qilish"), str(int(_num(stats.get("cancelled")))), _t(lang, f"доля {_num(stats.get('cancellation_rate')) * 100:.1f}%", f"ulushi {_num(stats.get('cancellation_rate')) * 100:.1f}%"), ORANGE, ORANGE_SOFT)

    chart_y, chart_h = 315, 172
    _rounded(c, MARGIN, chart_y, W - 2 * MARGIN, chart_h, white, 12)
    c.setStrokeColor(LINE)
    c.roundRect(MARGIN, chart_y, W - 2 * MARGIN, chart_h, 12, fill=0, stroke=1)
    c.setFillColor(INK)
    c.setFont(FONT_BOLD, 10)
    c.drawString(MARGIN + 14, chart_y + chart_h - 22, _t(lang, "Выручка по дням", "Kunlik tushum"))
    _line_chart(c, MARGIN + 14, chart_y + 8, W - 2 * MARGIN - 28, chart_h - 38, list(payload.get("daily") or []))

    _rounded(c, MARGIN, 66, W - 2 * MARGIN, 222, PURPLE_SOFT, 14)
    _pill(c, _t(lang, "КРАТКИЙ ВЫВОД", "QISQA XULOSA"), MARGIN + 14, 246, PURPLE)
    if not comparison_available:
        trend_title = _t(lang, "Сравнение недоступно", "Taqqoslash mavjud emas")
        trend_color = BLUE
    elif (revenue_change or 0) >= 0:
        trend_title = _t(lang, "Продажи растут", "Savdo o‘smoqda")
        trend_color = GREEN
    else:
        trend_title = _t(lang, "Продажи снизились", "Savdo kamaydi")
        trend_color = RED
    stock = list(payload.get("stock") or [])
    stock_available = bool(payload.get("stock_data_available", True))
    risk_count = sum(1 for row in stock if _num(row.get("total")) <= 0 or (row.get("days_left") is not None and _num(row.get("days_left")) <= 7))
    missing_count = int(_num(profit.get("missing_count")))
    if stock_available:
        stock_insight_title = _t(lang, "Риск дефицита", "Qoldiq xavfi")
        stock_insight_body = _t(lang, f"Товаров с нулевым или критическим прогнозом: {risk_count}.", f"Nol yoki xavfli prognozli tovarlar: {risk_count}.")
        stock_insight_color = ORANGE
    else:
        stock_insight_title = _t(lang, "Остатки недоступны", "Qoldiq mavjud emas")
        stock_insight_body = _t(lang, "Uzum API временно не вернул данные склада.", "Uzum API ombor ma’lumotini vaqtincha qaytarmadi.")
        stock_insight_color = BLUE
    insights = [
        (trend_color, trend_title, _t(lang, f"Выручка к предыдущему периоду: {_change_text(revenue_change, lang)}.", f"Tushum oldingi davrga nisbatan: {_change_text(revenue_change, lang)}.")),
        (stock_insight_color, stock_insight_title, stock_insight_body),
        (RED, _t(lang, "Отмены и возвраты", "Bekor va qaytarish"), _t(lang, f"Отменено {int(_num(stats.get('cancelled')))}, возвращено {_qty(stats.get('returns'))} шт.", f"Bekor {int(_num(stats.get('cancelled')))}, qaytarilgan {_qty(stats.get('returns'))} dona.")),
        (PURPLE, _t(lang, "Себестоимость Uzum", "Uzum tannarxi"), _t(lang, f"Без purchasePrice: {missing_count} SKU. Значения не подставляются.", f"purchasePricesiz: {missing_count} SKU. Qiymatlar taxmin qilinmaydi.")),
    ]
    c.setFillColor(INK)
    c.setFont(FONT_BOLD, 11)
    c.drawString(MARGIN + 14, 220, _t(lang, "Что важно владельцу магазина", "Do‘kon egasi uchun muhim"))
    yy = 190
    for color, item_title, body in insights:
        c.setFillColor(color)
        c.circle(MARGIN + 20, yy + 2, 4, fill=1, stroke=0)
        c.setFillColor(INK)
        c.setFont(FONT_BOLD, 8)
        c.drawString(MARGIN + 32, yy, _short(item_title, 31))
        c.setFillColor(MUTED)
        c.setFont(FONT_REGULAR, 7.4)
        c.drawString(MARGIN + 180, yy, _short(body, 72))
        yy -= 30
    _footer(c, 1, lang)
    c.showPage()


def _page_finance(c: canvas.Canvas, payload: dict[str, Any], lang: str) -> None:
    stats = dict(payload.get("stats") or {})
    profit = dict(payload.get("profit") or {})
    business = dict(payload.get("business_profit") or {})
    coverage = _num(business.get("coverage", profit.get("coverage")))
    business_complete = bool(business.get("complete"))
    _header(c, 2, _t(lang, "Продажи и прибыль", "Savdo va foyda"), _t(lang, "Расходы площадки, себестоимость и товары, формирующие результат", "Platforma xarajatlari, tannarx va natijani shakllantiruvchi tovarlar"), lang)
    _section(c, _t(lang, "Как получился результат", "Natija qanday hisoblandi"), _t(lang, "Каждая сумма показана в порядке расчёта; комиссия и логистика не вычитаются дважды", "Har bir summa hisob tartibida; komissiya va logistika ikki marta ayrilmaydi"), MARGIN, H - 135)
    calc_revenue = _num(business.get("calculation_revenue", stats.get("revenue")))
    calc_payout = _num(business.get("calculation_payout", stats.get("payout_total")))
    calc_commission = _num(business.get("calculation_commission", stats.get("commission")))
    calc_logistics = _num(business.get("calculation_logistics", stats.get("logistics")))
    residual = calc_revenue - calc_payout - calc_commission - calc_logistics
    other_payout = _num(business.get("other_payout_deductions", max(0.0, residual)))
    payout_adjustment = _num(business.get("payout_adjustment", max(0.0, -residual)))
    cost_total = _num(business.get("cost_total", profit.get("cost_total")))
    before_tax = _num(business.get("known_profit", calc_payout - cost_total))
    tax = _num(business.get("tax_expense"))
    signed_uzum = _num(business.get("uzum_expense_total"))
    uzum_deductions = _num(business.get("uzum_expense_deductions", max(0.0, signed_uzum)))
    uzum_refunds = _num(business.get("uzum_expense_refunds", max(0.0, -signed_uzum)))
    external = _num(
        business.get(
            "external_expense_total",
            _num(business.get("advertising_expense"))
            + _num(business.get("storage_expense"))
            + _num(business.get("other_expense")),
        )
    )
    profit_value = _num(business.get("net_profit", profit.get("profit")))
    result_label = _t(
        lang,
        "Чистая прибыль" if business_complete else "Результат по известным данным",
        "Sof foyda" if business_complete else "Ma’lum ma’lumotlar natijasi",
    )
    before_tax_color = RED if before_tax < 0 else GREEN
    result_color = RED if profit_value < 0 else GREEN
    bridge_rows = [
        [_t(lang, "Выручка для расчёта", "Hisob uchun tushum"), "+", _money(calc_revenue, lang)],
        [_t(lang, "Комиссия Uzum", "Uzum komissiyasi"), "−", _money(calc_commission, lang)],
        [_t(lang, "Логистика", "Logistika"), "−", _money(calc_logistics, lang)],
        [_t(lang, "Другие удержания внутри выплаты", "To‘lov ichidagi boshqa ushlanmalar"), "−", _money(other_payout, lang)],
        [_t(lang, "Корректировка выплаты", "To‘lov tuzatishi"), "+", _money(payout_adjustment, lang)],
        [_t(lang, "К выплате", "To‘lovga"), "=", _money(calc_payout, lang)],
        [_t(lang, "Себестоимость", "Tannarx"), "−", _money(cost_total, lang)],
        [_t(lang, "Прибыль до налога", "Soliqdan oldingi foyda"), "=", _money(before_tax, lang)],
        [_t(lang, "Налог", "Soliq"), "−", _money(tax, lang)],
        [_t(lang, "Доп. расходы Uzum", "Uzum qo‘shimcha xarajatlari"), "−", _money(uzum_deductions, lang)],
        [_t(lang, "Возвраты от Uzum", "Uzum qaytargan mablag‘"), "+", _money(uzum_refunds, lang)],
        [_t(lang, "Внешние расходы", "Tashqi xarajatlar"), "−", _money(external, lang)],
        [result_label, "=", _money(profit_value, lang)],
    ]
    bridge_colors = [INK, RED, RED, RED, GREEN, BLUE, RED, before_tax_color, RED, RED, GREEN, RED, result_color]
    _table(
        c,
        MARGIN,
        690,
        [302, 45, W - 2 * MARGIN - 347],
        [_t(lang, "ПОКАЗАТЕЛЬ", "KO‘RSATKICH"), _t(lang, "ЗНАК", "BELGI"), _t(lang, "СУММА", "SUMMA")],
        bridge_rows,
        row_h=13.5,
        right={1, 2},
        last_colors=bridge_colors,
    )

    _section(c, _t(lang, "Товары-лидеры", "Yetakchi tovarlar"), _t(lang, "Топ по выручке с оценкой прибыли и маржи", "Tushum bo‘yicha top, foyda va marja bahosi"), MARGIN, 494)
    product_rows: list[list[str]] = []
    margin_colors: list[Any] = []
    for item in list(payload.get("products") or [])[:6]:
        margin = item.get("margin")
        margin_text = "-" if margin is None else f"{_num(margin):.1f}%"
        product_rows.append([
            _short(item.get("title") or item.get("sku"), 50),
            _qty(item.get("qty")),
            _money(item.get("revenue"), lang),
            _money(item.get("profit"), lang) if item.get("profit") is not None else _t(lang, "нет данных", "ma’lumot yo‘q"),
            margin_text,
        ])
        margin_colors.append(MUTED if margin is None else RED if _num(margin) < 0 else ORANGE if _num(margin) < 10 else GREEN)
    if not product_rows:
        product_rows.append([_t(lang, "Продаж не найдено", "Savdo topilmadi"), "-", "-", "-", "-"])
        margin_colors.append(MUTED)
    _table(c, MARGIN, 466, [202, 50, 92, 92, 65], [_t(lang, "ТОВАР", "TOVAR"), _t(lang, "ШТ.", "DONA"), _t(lang, "ВЫРУЧКА", "TUSHUM"), _t(lang, "ПРИБЫЛЬ", "FOYDA"), _t(lang, "МАРЖА", "MARJA")], product_rows, row_h=26, right={1, 2, 3, 4}, last_colors=margin_colors)

    _section(c, _t(lang, "Качество расчёта", "Hisob sifati"), _t(lang, "Покрытие себестоимостью защищает отчёт от ложной общей прибыли", "Tannarx qamrovi noto‘g‘ri umumiy foydadan himoya qiladi"), MARGIN, 268)
    gap = 10
    card_w = (W - 2 * MARGIN - 2 * gap) / 3
    _kpi(c, MARGIN, 166, card_w, 76, _t(lang, "Покрытие", "Qamrov"), f"{coverage * 100:.1f}%", _t(lang, "по выручке", "tushum bo‘yicha"), GREEN if coverage >= 0.8 else ORANGE, GREEN_SOFT if coverage >= 0.8 else ORANGE_SOFT)
    _kpi(c, MARGIN + card_w + gap, 166, card_w, 76, _t(lang, "С себестоимостью", "Tannarxli"), str(int(_num(profit.get("known_count")))), "SKU", PURPLE, PURPLE_SOFT)
    _kpi(c, MARGIN + 2 * (card_w + gap), 166, card_w, 76, _t(lang, "Без себестоимости", "Tannarxsiz"), str(int(_num(profit.get("missing_count")))), "SKU", RED if coverage < 0.8 else ORANGE, RED_SOFT if coverage < 0.8 else ORANGE_SOFT)
    _rounded(c, MARGIN, 67, W - 2 * MARGIN, 76, BG, 12)
    c.setFillColor(PURPLE)
    c.setFont(FONT_BOLD, 8.5)
    c.drawString(MARGIN + 14, 121, _t(lang, "Как читать прибыль", "Foydani qanday o‘qish kerak"))
    expenses_available = bool(business.get("uzum_expenses_available", True))
    note = _t(
        lang,
        "При покрытии ниже 100% результат относится только к SKU с purchasePrice. Бот не подставляет среднюю себестоимость."
        + (" Расходы Uzum временно недоступны — итог неполный." if not expenses_available else ""),
        "Qamrov 100% dan past bo‘lsa, natija faqat purchasePrice bor SKUlarga tegishli. Bot o‘rtacha tannarxni taxmin qilmaydi."
        + (" Uzum xarajatlari vaqtincha olinmadi — natija to‘liq emas." if not expenses_available else ""),
    )
    _paragraph(c, note, MARGIN + 14, 107, W - 2 * MARGIN - 28, size=7.5, leading=10.5, color=MUTED)
    _footer(c, 2, lang)
    c.showPage()


def _page_stock(c: canvas.Canvas, payload: dict[str, Any], lang: str) -> None:
    stock = list(payload.get("stock") or [])
    stock_available = bool(payload.get("stock_data_available", True))
    total_sku = len([row for row in stock if not row.get("loss_only")]) if stock_available else 0
    low = [row for row in stock if _num(row.get("total")) > 0 and (_num(row.get("total")) <= _num(row.get("low_stock_threshold") or 5) or (row.get("days_left") is not None and _num(row.get("days_left")) <= 7))] if stock_available else []
    zero = [row for row in stock if _num(row.get("total")) <= 0 and not row.get("loss_only")] if stock_available else []
    no_sales = [row for row in stock if _num(row.get("total")) > 0 and _num(row.get("sold_7")) <= 0] if stock_available else []
    _header(c, 3, _t(lang, "Складские риски и план действий", "Ombor xavflari va harakat rejasi"), _t(lang, "Что требует внимания сегодня, чтобы не терять продажи", "Savdoni yo‘qotmaslik uchun bugun nimalarga e’tibor berish kerak"), lang)
    _section(c, _t(lang, "Состояние склада", "Ombor holati"), _t(lang, "Остаток и прогноз рассчитаны по текущему темпу продаж", "Qoldiq va prognoz joriy savdo tezligi asosida hisoblangan"), MARGIN, H - 135)
    gap = 10
    card_w = (W - 2 * MARGIN - 3 * gap) / 4
    cards = [
        (_t(lang, "SKU в продаже", "Savdodagi SKU"), str(total_sku) if stock_available else "-", BLUE, BLUE_SOFT),
        (_t(lang, "Низкий остаток", "Kam qoldiq"), str(len(low)) if stock_available else "-", ORANGE, ORANGE_SOFT),
        (_t(lang, "Нет в наличии", "Qoldiq yo‘q"), str(len(zero)) if stock_available else "-", RED, RED_SOFT),
        (_t(lang, "Без продаж", "Savdosiz"), str(len(no_sales)) if stock_available else "-", PURPLE, PURPLE_SOFT),
    ]
    for index, (label, value, color, soft) in enumerate(cards):
        x = MARGIN + index * (card_w + gap)
        _rounded(c, x, 609, card_w, 70, soft, 12)
        c.setFillColor(color)
        c.setFont(FONT_BOLD, 16)
        c.drawString(x + 12, 644, value)
        c.setFillColor(INK)
        c.setFont(FONT_REGULAR, 7)
        c.drawString(x + 12, 624, _short(label, 22))

    _section(c, _t(lang, "Товары, которые скоро закончатся", "Tez tugaydigan tovarlar"), _t(lang, "Приоритет рассчитан по остатку и скорости продаж", "Ustuvorlik qoldiq va savdo tezligi bo‘yicha hisoblangan"), MARGIN, 596)
    low_sorted = sorted(low + zero, key=lambda row: (0 if _num(row.get("total")) <= 0 else 1, _num(row.get("days_left")) if row.get("days_left") is not None else 9999, _num(row.get("total"))))
    stock_rows: list[list[str]] = []
    priority_colors: list[Any] = []
    for row in low_sorted[:5]:
        total = _num(row.get("total"))
        days_left = row.get("days_left")
        if total <= 0:
            priority = _t(lang, "Срочно", "Shoshilinch")
            color = RED
        elif days_left is not None and _num(days_left) <= 3:
            priority = _t(lang, "Высокий", "Yuqori")
            color = ORANGE
        else:
            priority = _t(lang, "Средний", "O‘rta")
            color = PURPLE
        stock_rows.append([
            _short(row.get("title") or row.get("sku"), 48),
            _qty(total),
            f"{_num(row.get('sold_7')) / 7:.1f}/{_t(lang, 'день', 'kun')}",
            "-" if days_left is None else f"{_num(days_left):.0f} {_t(lang, 'дн.', 'kun')}",
            priority,
        ])
        priority_colors.append(color)
    if not stock_rows:
        if stock_available:
            stock_rows.append([_t(lang, "Критических остатков нет", "Xavfli qoldiq yo‘q"), "-", "-", "-", _t(lang, "Норма", "Yaxshi")])
            priority_colors.append(GREEN)
        else:
            stock_rows.append([_t(lang, "Данные остатков временно недоступны", "Qoldiq ma’lumoti vaqtincha mavjud emas"), "-", "-", "-", _t(lang, "Повторить", "Qayta urinish")])
            priority_colors.append(ORANGE)
    _table(c, MARGIN, 568, [205, 58, 82, 70, 85], [_t(lang, "ТОВАР", "TOVAR"), _t(lang, "ОСТАТОК", "QOLDIQ"), _t(lang, "СКОРОСТЬ", "TEZLIK"), _t(lang, "ХВАТИТ", "YETADI"), _t(lang, "ПРИОРИТЕТ", "USTUVOR")], stock_rows, row_h=27, right={1, 2, 3, 4}, last_colors=priority_colors)

    _section(c, _t(lang, "Рекомендуемый план", "Tavsiya etilgan reja"), _t(lang, "Задачи расположены по влиянию на продажи и прибыль", "Vazifalar savdo va foydaga ta’siri bo‘yicha joylashtirilgan"), MARGIN, 388)
    actions = list(payload.get("actions") or [])[:4]
    if not actions:
        actions = [{"title_ru": "Критических действий нет", "title_uz": "Muhim vazifa yo‘q", "body_ru": "Продолжайте контролировать продажи и остатки.", "body_uz": "Savdo va qoldiqlarni nazorat qilishda davom eting.", "priority": "info"}]
    positions = [(MARGIN, 289), (MARGIN + 255 + 10, 289), (MARGIN, 199), (MARGIN + 255 + 10, 199)]
    colors = {"critical": (RED, RED_SOFT), "warning": (ORANGE, ORANGE_SOFT), "info": (BLUE, BLUE_SOFT)}
    for index, (action, (x, y)) in enumerate(zip(actions, positions), start=1):
        accent, soft = colors.get(str(action.get("priority") or "info"), (PURPLE, PURPLE_SOFT))
        if lang.startswith("uz"):
            title = str(action.get("title_uz") or action.get("title_ru") or action.get("title") or "-")
            body = str(action.get("body_uz") or action.get("body_ru") or action.get("body") or "-")
        else:
            title = str(action.get("title_ru") or action.get("title_uz") or action.get("title") or "-")
            body = str(action.get("body_ru") or action.get("body_uz") or action.get("body") or "-")
        _priority_card(c, x, y, 255, 78, index, title, body, accent, soft)
    _rounded(c, MARGIN, 67, W - 2 * MARGIN, 105, PURPLE_DARK, 14)
    c.setFillColor(white)
    c.setFont(FONT_BOLD, 10.5)
    c.drawString(MARGIN + 15, 145, _t(lang, "Фокус на следующий период", "Keyingi davr uchun fokus"))
    c.setFillColor(HexColor("#DED5FF"))
    c.setFont(FONT_REGULAR, 7.5)
    focus_text = (
        _t(lang, "Сначала устраните дефицит и отмены, затем работайте с маржой и товарами без продаж.", "Avval qoldiq va bekorlarni hal qiling, keyin marja va savdosiz tovarlar bilan ishlang.")
        if stock_available
        else _t(lang, "Повторите отчёт позже; остальные доступные рекомендации сохранены выше.", "Hisobotni keyinroq qaytaring; boshqa mavjud tavsiyalar yuqorida saqlandi.")
    )
    c.drawString(MARGIN + 15, 125, focus_text)
    c.setFillColor(white)
    c.setFont(FONT_BOLD, 13)
    if stock_available:
        c.drawString(MARGIN + 15, 92, _t(lang, f"Критических остатков: {len(zero) + len(low)}", f"Xavfli qoldiqlar: {len(zero) + len(low)}"))
        c.drawString(MARGIN + 290, 92, _t(lang, f"Без продаж: {len(no_sales)}", f"Savdosiz: {len(no_sales)}"))
    else:
        c.drawString(MARGIN + 15, 92, _t(lang, "Данные склада: временно недоступны", "Ombor ma’lumoti: vaqtincha mavjud emas"))
    _footer(c, 3, lang)
    c.showPage()


def _page_problems(c: canvas.Canvas, payload: dict[str, Any], lang: str) -> None:
    problems = list(payload.get("problems") or [])
    cancellations = [row for row in problems if row.get("kind") == "cancel"]
    returns = [row for row in problems if row.get("kind") == "return"]
    defect_events = list(payload.get("defect_events") or [])
    cumulative = list(payload.get("cumulative_defects") or [])
    loss_data_available = bool(payload.get("loss_data_available", True))
    cancelled_units = sum(_num(row.get("qty")) for row in cancellations)
    cancelled_value = sum(_num(row.get("revenue")) for row in cancellations)
    returned_units = sum(_num(row.get("qty")) for row in returns)
    defect_units = sum(_num(row.get("defected_delta")) for row in defect_events)
    defect_value = sum(_num(row.get("estimated_value")) for row in defect_events)
    _header(c, 4, _t(lang, "Отмены, возвраты и брак", "Bekor, qaytarish va yaroqsiz"), _t(lang, "Финансовое влияние проблемных операций и товары, требующие проверки", "Muammoli operatsiyalarning moliyaviy ta’siri va tekshiriladigan tovarlar"), lang)
    _section(c, _t(lang, "Итог за период", "Davr yakuni"), _t(lang, "Отмены и возвраты - Finance API; новый брак - история изменений FBO", "Bekor va qaytarish - Finance API; yangi yaroqsiz - FBO o‘zgarishlar tarixi"), MARGIN, H - 135)
    gap = 10
    card_w = (W - 2 * MARGIN - 2 * gap) / 3
    y1, y2 = 610, 532
    _kpi(c, MARGIN, y1, card_w, 68, _t(lang, "Отменено заказов", "Bekor buyurtma"), str(len({str(row.get('order_id')) for row in cancellations})), _t(lang, f"{_qty(cancelled_units)} шт.", f"{_qty(cancelled_units)} dona"), RED, RED_SOFT)
    _kpi(c, MARGIN + card_w + gap, y1, card_w, 68, _t(lang, "Сумма отмен", "Bekor summasi"), _money(cancelled_value, lang), _t(lang, "по цене продажи", "sotuv narxida"), RED, RED_SOFT)
    _kpi(c, MARGIN + 2 * (card_w + gap), y1, card_w, 68, _t(lang, "Возвраты", "Qaytarish"), f"{_qty(returned_units)} {_t(lang, 'шт.', 'dona')}", str(len(returns)), ORANGE, ORANGE_SOFT)
    _kpi(c, MARGIN, y2, card_w, 68, _t(lang, "Новый брак", "Yangi yaroqsiz"), f"+{_qty(defect_units)} {_t(lang, 'шт.', 'dona')}", _t(lang, f"{len(defect_events)} событий", f"{len(defect_events)} hodisa"), PURPLE, PURPLE_SOFT)
    _kpi(c, MARGIN + card_w + gap, y2, card_w, 68, _t(lang, "Оценка нового брака", "Yangi brak bahosi"), _money(defect_value, lang), _t(lang, "по цене продажи", "sotuv narxida"), PURPLE, PURPLE_SOFT)
    cumulative_qty = sum(_num(row.get("defected")) for row in cumulative)
    _kpi(c, MARGIN + 2 * (card_w + gap), y2, card_w, 68, _t(lang, "Брак накопительно", "Jami yaroqsiz"), f"{_qty(cumulative_qty)} {_t(lang, 'шт.', 'dona')}" if loss_data_available else "-", _t(lang, "по данным Uzum", "Uzum ma’lumoti") if loss_data_available else _t(lang, "нет данных", "ma’lumot yo‘q"), BLUE if loss_data_available else ORANGE, BLUE_SOFT if loss_data_available else ORANGE_SOFT)

    _section(c, _t(lang, "Последние отмены и возвраты", "Oxirgi bekor va qaytarishlar"), _t(lang, "В таблице показаны наиболее значимые строки выбранного периода", "Jadvalda tanlangan davrning muhim qatorlari"), MARGIN, 500)
    problem_rows: list[list[str]] = []
    status_colors: list[Any] = []
    for row in sorted(problems, key=lambda value: _num(value.get("revenue")), reverse=True)[:5]:
        date_value = row.get("date")
        date_text = date_value.strftime("%d.%m") if isinstance(date_value, datetime) else str(date_value or "-")[:10]
        reason = str(row.get("reason") or _t(lang, "Причина не передана Uzum", "Sabab Uzum tomonidan berilmagan"))
        kind = _t(lang, "Отмена", "Bekor") if row.get("kind") == "cancel" else _t(lang, "Возврат", "Qaytarish")
        problem_rows.append([date_text, _short(row.get("order_id"), 18), _short(row.get("title"), 34), _qty(row.get("qty")), _money(row.get("revenue"), lang), _short(f"{kind}: {reason}", 35)])
        status_colors.append(RED if row.get("kind") == "cancel" else ORANGE)
    if not problem_rows:
        problem_rows.append(["-", "-", _t(lang, "Отмен и возвратов не найдено", "Bekor va qaytarish topilmadi"), "-", "-", _t(lang, "Норма", "Yaxshi")])
        status_colors.append(GREEN)
    _table(c, MARGIN, 472, [45, 76, 165, 38, 80, 97], [_t(lang, "ДАТА", "SANA"), _t(lang, "ЗАКАЗ", "BUYURTMA"), _t(lang, "ТОВАР", "TOVAR"), _t(lang, "ШТ.", "DONA"), _t(lang, "СУММА", "SUMMA"), _t(lang, "СТАТУС / ПРИЧИНА", "STATUS / SABAB")], problem_rows, row_h=27, right={1, 3, 4}, last_colors=status_colors)

    _section(c, _t(lang, "Новый брак FBO", "Yangi FBO yaroqsiz"), _t(lang, "Прирост за выбранный период; накопительные значения не считаются повторно", "Tanlangan davrdagi o‘sish; jami qiymatlar qayta hisoblanmaydi"), MARGIN, 305)
    defect_rows: list[list[str]] = []
    defect_colors: list[Any] = []
    for row in sorted(defect_events, key=lambda value: _num(value.get("estimated_value")), reverse=True)[:3]:
        defect_rows.append([_short(row.get("product_title"), 45), _short(row.get("sku_id") or row.get("barcode"), 22), f"+{_qty(row.get('defected_delta'))}", _qty(row.get("defected_qty")), _money(row.get("estimated_value"), lang)])
        defect_colors.append(PURPLE)
    if not defect_rows:
        defect_rows.append([_t(lang, "Новый брак за период не зафиксирован", "Davrda yangi yaroqsiz qayd etilmadi"), "-", "0", "-", _money(0, lang)])
        defect_colors.append(GREEN)
    _table(c, MARGIN, 277, [190, 105, 65, 70, 71], [_t(lang, "ТОВАР", "TOVAR"), "SKU", _t(lang, "НОВЫЙ", "YANGI"), _t(lang, "ВСЕГО", "JAMI"), _t(lang, "ОЦЕНКА", "BAHO")], defect_rows, row_h=27, right={2, 3, 4}, last_colors=defect_colors)

    _rounded(c, MARGIN, 67, W - 2 * MARGIN, 89, ORANGE_SOFT, 12)
    c.setFillColor(ORANGE)
    c.setFont(FONT_BOLD, 8.5)
    c.drawString(MARGIN + 14, 133, _t(lang, "Важно о качестве данных", "Ma’lumot sifati haqida"))
    if loss_data_available:
        note = _t(
            lang,
            "Uzum не всегда передаёт причину отмены и дату накопительного брака. В таком случае отчёт показывает доступные факты и прямо отмечает отсутствие причины. История нового брака ведётся ботом с момента установки этой версии.",
            "Uzum har doim bekor sababini va jami yaroqsiz sanasini bermaydi. Hisobot mavjud faktlarni ko‘rsatadi va sabab yo‘qligini ochiq yozadi. Yangi yaroqsiz tarixi ushbu versiya o‘rnatilgandan boshlab yuritiladi.",
        )
    else:
        note = _t(
            lang,
            "Накопительный брак временно не загрузился из Uzum API. Отмены, возвраты и уже записанная ботом история нового брака остаются в отчёте. Повторите формирование позже.",
            "Jamlangan brak Uzum API dan vaqtincha yuklanmadi. Bekor, qaytarish va bot yozgan yangi brak tarixi hisobotda qoldi. Keyinroq qayta urinib ko‘ring.",
        )
    _paragraph(c, note, MARGIN + 14, 119, W - 2 * MARGIN - 28, size=7.3, leading=10.2, color=INK)
    _footer(c, 4, lang)
    c.showPage()


def build_seller_pdf_report(
    payload: dict[str, Any],
    output: str | Path,
    *,
    lang: str = "ru",
    regular_font_path: str | Path,
    bold_font_path: str | Path,
) -> Path:
    """Render a fixed four-page management report from already collected data."""
    _register_fonts(regular_font_path, bold_font_path)
    output_path = Path(output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    normalized_lang = "uz" if str(lang).lower().startswith("uz") else "ru"
    document = canvas.Canvas(str(output_path), pagesize=A4, pageCompression=1)
    document.setTitle(_t(normalized_lang, "Seller.pro.uz - управленческий отчёт", "Seller.pro.uz - boshqaruv hisoboti"))
    document.setAuthor("Seller.pro.uz")
    document.setSubject(_t(normalized_lang, "Отчёт продавца Uzum", "Uzum sotuvchisi hisoboti"))
    _page_summary(document, payload, normalized_lang)
    _page_finance(document, payload, normalized_lang)
    _page_stock(document, payload, normalized_lang)
    _page_problems(document, payload, normalized_lang)
    document.save()
    return output_path
