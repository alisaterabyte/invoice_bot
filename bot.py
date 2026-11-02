import logging
import os
import json
import base64
from datetime import datetime
from io import BytesIO

import anthropic
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    filters,
    ContextTypes,
)
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import cm
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BOT_TOKEN        = os.environ["BOT_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
DATABASE_URL     = os.environ.get("DATABASE_URL")  # set by Railway PostgreSQL

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

ai = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

# ---------------------------------------------------------------------------
# DB layer — PostgreSQL on Railway, SQLite locally
# ---------------------------------------------------------------------------

def _connect():
    if DATABASE_URL:
        import psycopg2
        conn = psycopg2.connect(DATABASE_URL)
        return conn, "%s"
    else:
        import sqlite3
        conn = sqlite3.connect("invoices.db")
        return conn, "?"

def init_db():
    conn, ph = _connect()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS invoices (
            id          SERIAL PRIMARY KEY,
            invoice_no  TEXT,
            supplier    TEXT,
            item_name   TEXT NOT NULL,
            quantity    REAL,
            unit        TEXT,
            unit_price  REAL,
            total_price REAL,
            paid_amount REAL,
            currency    TEXT DEFAULT 'USD',
            status      TEXT NOT NULL DEFAULT 'pending',
            tracking    TEXT,
            eta         TEXT,
            notes       TEXT,
            created_at  TEXT DEFAULT (to_char(now(),'YYYY-MM-DD HH24:MI:SS'))
        )
    """ if DATABASE_URL else """
        CREATE TABLE IF NOT EXISTS invoices (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_no  TEXT,
            supplier    TEXT,
            item_name   TEXT NOT NULL,
            quantity    REAL,
            unit        TEXT,
            unit_price  REAL,
            total_price REAL,
            paid_amount REAL,
            currency    TEXT DEFAULT 'USD',
            status      TEXT NOT NULL DEFAULT 'pending',
            tracking    TEXT,
            eta         TEXT,
            notes       TEXT,
            created_at  TEXT DEFAULT (datetime('now'))
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS chat_history (
            id         SERIAL PRIMARY KEY,
            user_id    BIGINT,
            role       TEXT,
            content    TEXT,
            created_at TEXT DEFAULT (to_char(now(),'YYYY-MM-DD HH24:MI:SS'))
        )
    """ if DATABASE_URL else """
        CREATE TABLE IF NOT EXISTS chat_history (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id    INTEGER,
            role       TEXT,
            content    TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS context_notes (
            id      SERIAL PRIMARY KEY,
            note    TEXT,
            created_at TEXT DEFAULT (to_char(now(),'YYYY-MM-DD HH24:MI:SS'))
        )
    """ if DATABASE_URL else """
        CREATE TABLE IF NOT EXISTS context_notes (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            note    TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        )
    """)
    conn.commit()
    conn.close()

def db_fetchall(query: str, params: tuple = ()):
    conn, ph = _connect()
    if params:
        query = query.replace("?", ph)
    cur = conn.cursor()
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    return rows

def db_execute(query: str, params: tuple = ()):
    conn, ph = _connect()
    if params:
        query = query.replace("?", ph)
    cur = conn.cursor()
    cur.execute(query, params)
    conn.commit()
    conn.close()

def get_db_summary() -> str:
    rows = db_fetchall(
        "SELECT id, invoice_no, supplier, item_name, quantity, unit, total_price, paid_amount, currency, status, tracking, eta, notes FROM invoices ORDER BY created_at DESC"
    )
    if not rows:
        return "База данных пустая."
    header = "id|invoice_no|supplier|item_name|qty|unit|total|paid|currency|status|tracking|eta|notes"
    lines = [header]
    for r in rows:
        lines.append("|".join("" if x is None else str(x) for x in r))
    return "\n".join(lines)

def get_context_notes() -> str:
    rows = db_fetchall("SELECT note FROM context_notes ORDER BY id DESC")
    if not rows:
        return ""
    return "\n".join(r[0] for r in rows)

def execute_db_action(action: dict) -> str:
    op = action.get("operation")
    if op == "insert":
        items = action.get("items", [action.get("item", {})])
        for item in items:
            db_execute("""
                INSERT INTO invoices (invoice_no, supplier, item_name, quantity, unit, unit_price, total_price, paid_amount, currency, status, tracking, eta, notes)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                item.get("invoice_no"), item.get("supplier"), item.get("item_name", "Unknown"),
                item.get("quantity"), item.get("unit"), item.get("unit_price"),
                item.get("total_price"), item.get("paid_amount"), item.get("currency", "USD"),
                item.get("status", "pending"), item.get("tracking"), item.get("eta"), item.get("notes"),
            ))
        return f"✅ Добавлено {len(items)} позиций."
    elif op == "update":
        row_id = action.get("id")
        fields = action.get("fields", {})
        if not row_id or not fields:
            return "❌ Нужен id и fields."
        set_clause = ", ".join(f"{k}=?" for k in fields)
        values = tuple(fields.values()) + (row_id,)
        db_execute(f"UPDATE invoices SET {set_clause} WHERE id=?", values)
        return f"✅ Запись #{row_id} обновлена."
    elif op == "delete":
        db_execute("DELETE FROM invoices WHERE id=?", (action.get("id"),))
        return f"✅ Запись #{action.get('id')} удалена."
    elif op == "save_note":
        db_execute("INSERT INTO context_notes (note) VALUES (?)", (action.get("note", ""),))
        return "✅ Контекст сохранён."
    return f"❌ Неизвестная операция: {op}"

def get_history(user_id: int, limit: int = 8) -> list:
    rows = db_fetchall(
        "SELECT role, content FROM chat_history WHERE user_id=? ORDER BY id DESC LIMIT ?",
        (user_id, limit)
    )
    return [{"role": r, "content": c} for r, c in reversed(rows)]

def save_message(user_id: int, role: str, content: str):
    db_execute("INSERT INTO chat_history (user_id, role, content) VALUES (?,?,?)", (user_id, role, content))

# ---------------------------------------------------------------------------
# AI
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """Ты — умный ассистент для управления инвойсами компании UVS GROUP INC (Natalia Samailovich).
Компания продаёт конфеты и заказывает упаковочные пакеты из Китая.

━━━ ТАБЛИЦА РАЗМЕРОВ И НАЗВАНИЙ ПАКЕТОВ ━━━

Все пакеты бывают двух видов: Donty Tonty (с печатью бренда) и Clear (прозрачные без печати).
Также есть Custom пакеты — кастомные с дизайном конкретного продукта (Donty Tonty Custom Cherry Sour, HoneyBrook Sweets Custom и т.д.)

| Shortcut       | Размер (см)         | Размер (дюйм) |
|----------------|---------------------|---------------|
| Mini           | 15.5 × 20.5 + 9 cm  | 6 × 8.5"      |
| Medium         | 18 × 25.5 + 10 cm   | 7 × 10"       |
| Size 2 / Двойка| 20.5 × 30 + 10 cm   | 8 × 12"       |
| Size 3 / Тройка| 22 × 34.5 + 10 cm   | 8 × 13"       |
| Size 4 / MAX   | 25.5 × 37.5 + 10 cm | 10 × 15"      |
| Roll Film      | 48 cm ширина        | —             |

Roll Film считается в кг. Один ролл ≈ 30 кг.

ПРАВИЛО НАЗВАНИЯ: всегда пиши «Clear Size 3» или «Donty Tonty Size 2» или «Clear Mini» — не пиши сырые размеры в сантиметрах если знаешь shortcut.

━━━ ПОСТАВЩИКИ ━━━
- QD XTD Plastic Products Co., Ltd — основной поставщик пакетов (Qingdao, China). Контакт: Ben, access1@xtdplastic.com, WhatsApp: +8617663936931
- Shanghai Xiangcheng International Logistics Co., Ltd — логистический агент (доставка морем)

━━━ СПОСОБЫ ДОСТАВКИ ━━━
- АВИА — быстро, небольшие партии
- МОРЕ (морской контейнер) — большие партии, ~30-45 дней из Шанхая до Нью-Йорка

━━━ УСЛОВИЯ ОПЛАТЫ ━━━
50% аванс при заказе, 50% когда пакеты готовы перед отправкой.

━━━ ТЕКУЩИЕ ДАННЫЕ В БАЗЕ ━━━
{db_summary}

━━━ ДОПОЛНИТЕЛЬНЫЙ КОНТЕКСТ ━━━
{context_notes}

━━━ СХЕМА ТАБЛИЦЫ invoices ━━━
id | invoice_no | supplier | item_name | quantity | unit | unit_price | total_price | paid_amount | currency | status | tracking | eta | notes

СТАТУСЫ:
- pending — ожидает оплаты или отправки
- paid — оплачено (50% или 100%), ждём отправки
- ready — готово на складе поставщика, не отправлено
- in_transit — в пути (есть трекинг номер)
- delivered — доставлено на склад

━━━ КАК ОТОБРАЖАТЬ ПОЗИЦИИ ━━━
Формат: [Название shortcut] — [кол-во] шт | $[сумма] | [статус]
Пример: Clear Size 3 — 51 300 шт | $4 155 | 🚚 В пути

━━━ ТВОИ ЗАДАЧИ ━━━
1. Инвойс (текст или фото) → извлеки ВСЕ позиции, переведи размеры в shortcut-названия, внеси в базу
2. Сообщение об оплате/статусе/трекинге → обновляй записи
3. Объяснение о товарах → сохрани как контекст
4. Вопросы → отвечай по данным из базы

ЕСЛИ НУЖНО ИЗМЕНИТЬ БАЗУ, добавь JSON-блок в конце:
```json
{{"action": {{"operation": "insert", "items": [{{"invoice_no": "XTDQ-B20260905MK", "supplier": "QD XTD", "item_name": "Clear Size 3", "quantity": 51300, "unit": "шт", "unit_price": 0.081, "total_price": 4155.30, "currency": "USD", "status": "in_transit", "tracking": "CSGU6424856"}}]}}}}
```
или обновление:
```json
{{"action": {{"operation": "update", "id": 5, "fields": {{"status": "delivered", "paid_amount": 11918.25}}}}}}
```
или контекст:
```json
{{"action": {{"operation": "save_note", "note": "..."}}}}
```

Общайся на русском. Будь конкретным и кратким."""

async def ask_ai(user_id: int, user_message: str, image_data: bytes = None) -> tuple[str, dict | None]:
    system = SYSTEM_PROMPT.format(
        db_summary=get_db_summary(),
        context_notes=get_context_notes() or "Пока нет сохранённого контекста."
    )
    history = get_history(user_id)

    if image_data:
        content = [
            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": base64.b64encode(image_data).decode()}},
            {"type": "text", "text": user_message or "Это инвойс. Извлеки все данные и внеси в базу."}
        ]
    else:
        content = user_message

    messages = history + [{"role": "user", "content": content}]

    response = ai.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=2048,
        system=system,
        messages=messages,
    )
    reply = response.content[0].text

    db_action = None
    if "```json" in reply:
        try:
            j_start = reply.index("```json") + 7
            j_end   = reply.index("```", j_start)
            parsed  = json.loads(reply[j_start:j_end].strip())
            if "action" in parsed:
                db_action = parsed["action"]
            reply = reply[:reply.index("```json")].strip()
        except Exception as e:
            logger.warning(f"JSON parse error: {e}")

    save_message(user_id, "user", user_message or "[image]")
    save_message(user_id, "assistant", reply)
    return reply, db_action

# ---------------------------------------------------------------------------
# Exports
# ---------------------------------------------------------------------------

STATUS_RU = {
    "pending":    "⏳ Ожидает",
    "paid":       "💰 Оплачено",
    "ready":      "✅ Готово",
    "in_transit": "🚚 В пути",
    "delivered":  "📦 Доставлено",
}

def _get_rows():
    return db_fetchall(
        "SELECT id, invoice_no, supplier, item_name, quantity, unit, total_price, paid_amount, currency, status, tracking, eta, notes FROM invoices ORDER BY status, invoice_no"
    )

def generate_pdf(path: str):
    doc = SimpleDocTemplate(path, pagesize=landscape(A4),
        leftMargin=1*cm, rightMargin=1*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Invoice Status Report", styles["Heading1"]),
        Paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}", styles["Normal"]),
        Spacer(1, 0.4*cm),
    ]
    rows = _get_rows()
    headers = ["#", "Invoice", "Поставщик", "Товар", "Кол-во", "Итого", "Оплачено", "Статус", "Трекинг", "ETA", "Заметки"]
    data = [headers]
    for r in rows:
        id_, inv, sup, item, qty, unit, total, paid, ccy, status, track, eta, notes = r
        data.append([
            str(id_), inv or "", sup or "", item or "",
            f"{qty} {unit or ''}" if qty else "",
            f"{total} {ccy}" if total else "",
            f"{paid} {ccy}" if paid else "",
            STATUS_RU.get(status, status),
            track or "", eta or "", notes or "",
        ])
    if len(data) == 1:
        data.append(["Нет данных"] + [""] * 10)

    col_w = [1*cm, 2.5*cm, 2.5*cm, 5*cm, 2*cm, 2.5*cm, 2.5*cm, 2.5*cm, 3.5*cm, 2*cm, 4*cm]
    table = Table(data, colWidths=col_w, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#2C3E50")),
        ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
        ("FONTSIZE",   (0,0), (-1,0), 8),
        ("FONTSIZE",   (0,1), (-1,-1), 7),
        ("GRID",       (0,0), (-1,-1), 0.4, colors.grey),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#EBF5FB")]),
        ("PADDING",    (0,0), (-1,-1), 3),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
    ]))
    elements.append(table)
    doc.build(elements)

def generate_excel(path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoices"

    header_fill = PatternFill("solid", fgColor="2C3E50")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    thin = Side(style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["ID", "Invoice #", "Поставщик", "Товар", "Кол-во", "Ед.", "Цена/ед.", "Итого", "Оплачено", "Валюта", "Статус", "Трекинг", "ETA", "Заметки", "Дата добавления"]
    col_widths = [5, 14, 18, 30, 8, 6, 10, 10, 10, 8, 14, 20, 12, 30, 18]

    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 20

    status_colors = {
        "pending":    "FFF9C4",
        "paid":       "C8E6C9",
        "ready":      "B3E5FC",
        "in_transit": "FFE0B2",
        "delivered":  "DCEDC8",
    }
    rows = _get_rows()
    for row_idx, r in enumerate(rows, 2):
        id_, inv, sup, item, qty, unit, uprice, total, paid, ccy, status, track, eta, notes = (
            r[0], r[1], r[2], r[3], r[4], r[5], None, r[6], r[7], r[8], r[9], r[10], r[11], r[12]
        )
        # fetch unit_price separately
        full = db_fetchall("SELECT unit_price, created_at FROM invoices WHERE id=?", (id_,))
        uprice2 = full[0][0] if full else None
        created = full[0][1] if full else ""
        values = [id_, inv, sup, item, qty, unit, uprice2, total, paid, ccy, STATUS_RU.get(status, status), track, eta, notes, created]
        fill_color = status_colors.get(status, "FFFFFF")
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.fill = PatternFill("solid", fgColor=fill_color)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=(col == 4))
        ws.row_dimensions[row_idx].height = 16

    ws.freeze_panes = "A2"
    wb.save(path)

# ---------------------------------------------------------------------------
# Summary text
# ---------------------------------------------------------------------------

def build_summary() -> str:
    rows = db_fetchall(
        "SELECT invoice_no, item_name, quantity, unit, total_price, paid_amount, currency, status, tracking, eta FROM invoices ORDER BY status, invoice_no"
    )
    if not rows:
        return "📭 Инвойсов пока нет."

    groups: dict[str, list] = {}
    for r in rows:
        s = r[7]
        groups.setdefault(s, []).append(r)

    order = ["in_transit", "ready", "paid", "pending", "delivered"]
    lines = ["📋 *Сводка по всем заказам*\n"]
    for status in order + [s for s in groups if s not in order]:
        if status not in groups:
            continue
        lines.append(f"\n{STATUS_RU.get(status, status)}")
        for inv, item, qty, unit, total, paid, ccy, _, track, eta in groups[status]:
            line = f"  • `{inv or '—'}` {item}"
            if qty:
                line += f"  {qty} {unit or ''}"
            if total:
                line += f"  | {total} {ccy}"
                if paid:
                    line += f" (оплачено {paid})"
            if track:
                line += f"\n    🔍 `{track}`"
            if eta:
                line += f"  📅 {eta}"
            lines.append(line)
    return "\n".join(lines)

# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------

MAIN_KB = InlineKeyboardMarkup([
    [InlineKeyboardButton("📊 Сводка", callback_data="summary")],
    [InlineKeyboardButton("🚢 Проверить доставки", callback_data="tracking")],
    [
        InlineKeyboardButton("📄 PDF", callback_data="pdf"),
        InlineKeyboardButton("📊 Excel", callback_data="excel"),
    ],
    [InlineKeyboardButton("🗑 Очистить историю", callback_data="clear_history")],
])

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *Invoice Bot*\n\n"
        "Просто пиши мне что нужно:\n"
        "• Скинь фото или текст инвойса — внесу сам\n"
        "• _«INV-003 оплачен полностью»_ — обновлю статус\n"
        "• _«Что сейчас в пути?»_ — покажу\n"
        "• _«Пакеты Donti-Tonti это...»_ — запомню контекст\n\n"
        "Или используй кнопки ниже 👇",
        parse_mode="Markdown",
        reply_markup=MAIN_KB,
    )

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "summary":
        text = build_summary()
        await query.edit_message_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🔄 Обновить", callback_data="summary")],
            [InlineKeyboardButton("« Меню", callback_data="menu")],
        ]))

    elif query.data == "pdf":
        await query.edit_message_text("⏳ Генерирую PDF...")
        path = "/tmp/invoices.pdf"
        generate_pdf(path)
        await query.message.reply_document(
            document=open(path, "rb"),
            filename=f"invoices_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            caption="📄 Отчёт по инвойсам",
        )
        await query.edit_message_text("✅ PDF отправлен.", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("« Меню", callback_data="menu")]]))

    elif query.data == "excel":
        await query.edit_message_text("⏳ Генерирую Excel...")
        path = "/tmp/invoices.xlsx"
        generate_excel(path)
        await query.message.reply_document(
            document=open(path, "rb"),
            filename=f"invoices_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            caption="📊 Таблица инвойсов",
        )
        await query.edit_message_text("✅ Excel отправлен.", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("« Меню", callback_data="menu")]]))

    elif query.data == "tracking":
        await show_tracking(query)

    elif query.data == "clear_history":
        user_id = query.from_user.id
        db_execute("DELETE FROM chat_history WHERE user_id=?", (user_id,))
        await query.edit_message_text("✅ История чата очищена.", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("« Меню", callback_data="menu")]]))

    elif query.data == "menu":
        await query.edit_message_text(
            "Чем могу помочь?",
            reply_markup=MAIN_KB,
        )

TRACKING_URLS = {
    "COSCO": "https://elines.coscoshipping.com/ebusiness/cargoTracking?trackingType=CONTAINER&number={}",
    "OOCL":  "https://www.oocl.com/eng/ourservices/eservices/cargotracking/Pages/cargotracking.aspx?{}",
    "MSC":   "https://www.msc.com/en/track-a-shipment?agencyPath=mwi&searchInfo={}",
    "MAERSK":"https://www.maersk.com/tracking/{}",
    "DEFAULT":"https://www.track-trace.com/container/{}",
}

def detect_carrier(container_no: str) -> str:
    prefix = container_no[:4].upper()
    if prefix in ("CSGU", "COSU", "CBHU"): return "COSCO"
    if prefix in ("OOLU", "OOCU"): return "OOCL"
    if prefix in ("MSCU", "MEDU"): return "MSC"
    if prefix in ("MAEU", "MSKU"): return "MAERSK"
    return "DEFAULT"

async def show_tracking(query):
    rows = db_fetchall(
        "SELECT invoice_no, item_name, tracking, eta, status FROM invoices WHERE tracking IS NOT NULL AND tracking != '' ORDER BY created_at DESC"
    )
    if not rows:
        await query.edit_message_text(
            "🚢 Нет доставок с трекинг-номерами.\n\nДобавь трекинг написав боту, например:\n_«трекинг для INV-001: CSGU6424856»_",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("« Меню", callback_data="menu")]]),
        )
        return

    lines = ["🚢 *Активные доставки:*\n"]
    seen = set()
    for inv, item, tracking, eta, status in rows:
        if tracking in seen:
            continue
        seen.add(tracking)
        carrier = detect_carrier(tracking)
        url = TRACKING_URLS[carrier].format(tracking)
        status_emoji = {"in_transit": "🚚", "ready": "✅", "delivered": "📦", "paid": "💰", "pending": "⏳"}.get(status, "•")
        lines.append(f"{status_emoji} *{item}* (`{inv}`)")
        lines.append(f"  Трекинг: `{tracking}`")
        if eta:
            lines.append(f"  ETA: {eta}")
        lines.append(f"  [Проверить на сайте {carrier}]({url})")
        lines.append("")

    kb = [[InlineKeyboardButton("🔄 Обновить", callback_data="tracking")],
          [InlineKeyboardButton("« Меню", callback_data="menu")]]
    await query.edit_message_text(
        "\n".join(lines),
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb),
        disable_web_page_preview=True,
    )

async def _process_message(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str, image_data: bytes = None):
    user_id = update.effective_user.id
    await update.message.chat.send_action("typing")
    try:
        reply, db_action = await ask_ai(user_id, text, image_data=image_data)
        action_result = ""
        if db_action:
            action_result = "\n\n" + execute_db_action(db_action)
        await update.message.reply_text(
            reply + action_result,
            parse_mode="Markdown",
            reply_markup=MAIN_KB,
        )
    except Exception as e:
        logger.error(f"Error: {e}", exc_info=True)
        await update.message.reply_text("⚠️ Ошибка. Попробуй ещё раз.")

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await _process_message(update, context, update.message.text)

async def photo_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    photo = update.message.photo[-1]
    file  = await context.bot.get_file(photo.file_id)
    buf   = BytesIO()
    await file.download_to_memory(buf)
    await _process_message(update, context, update.message.caption or "", image_data=buf.getvalue())

async def document_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if doc.mime_type in ("text/plain", "text/csv") or (doc.file_name or "").endswith((".txt", ".csv")):
        file = await context.bot.get_file(doc.file_id)
        buf  = BytesIO()
        await file.download_to_memory(buf)
        text = buf.getvalue().decode("utf-8", errors="ignore")[:4000]
        await _process_message(update, context, f"Файл {doc.file_name}:\n{text}")
    else:
        await _process_message(update, context, f"Пользователь прислал файл «{doc.file_name}». Скажи что поддерживаются только фото, текстовые файлы и CSV.")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    init_db()
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.PHOTO, photo_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, document_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))
    logger.info("Bot started")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
