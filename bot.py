import asyncio
import logging
import sqlite3
import os
from datetime import datetime

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    filters,
    ContextTypes,
)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

BOT_TOKEN = os.environ["BOT_TOKEN"]
DB_PATH = os.environ.get("DB_PATH", "invoices.db")

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------------

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS invoices (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_no  TEXT NOT NULL,
            supplier    TEXT,
            item_name   TEXT NOT NULL,
            quantity    REAL,
            unit        TEXT,
            unit_price  REAL,
            total_price REAL,
            currency    TEXT DEFAULT 'USD',
            status      TEXT NOT NULL DEFAULT 'pending',
            notes       TEXT,
            created_at  TEXT DEFAULT (datetime('now'))
        )
    """)
    conn.commit()
    conn.close()

# Status labels
STATUS_LABELS = {
    "ready":     "✅ Готово, не отправлено",
    "in_transit":"🚚 В пути",
    "delivered": "📦 Доставлено",
    "pending":   "⏳ Ожидает оплаты / отправки",
    "paid":      "💰 Оплачено, ждём отправки",
}

# ---------------------------------------------------------------------------
# PDF generation
# ---------------------------------------------------------------------------

def generate_pdf(path: str):
    doc = SimpleDocTemplate(
        path,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "Title",
        parent=styles["Heading1"],
        fontSize=16,
        spaceAfter=12,
    )
    elements = []

    elements.append(Paragraph("Invoice Status Report", title_style))
    elements.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles["Normal"]))
    elements.append(Spacer(1, 0.5 * cm))

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT invoice_no, supplier, item_name, quantity, unit, unit_price, total_price, currency, status, notes
        FROM invoices
        ORDER BY status, invoice_no
    """)
    rows = cur.fetchall()
    conn.close()

    headers = ["Invoice #", "Supplier", "Item", "Qty", "Unit", "Price", "Total", "CCY", "Status", "Notes"]
    table_data = [headers]
    for row in rows:
        inv_no, supplier, item, qty, unit, uprice, total, ccy, status, notes = row
        status_label = STATUS_LABELS.get(status, status)
        table_data.append([
            inv_no or "",
            supplier or "",
            item or "",
            str(qty) if qty else "",
            unit or "",
            str(uprice) if uprice else "",
            str(total) if total else "",
            ccy or "",
            status_label,
            notes or "",
        ])

    if len(table_data) == 1:
        table_data.append(["No invoices yet", "", "", "", "", "", "", "", "", ""])

    col_widths = [2.5*cm, 2.5*cm, 3.5*cm, 1.5*cm, 1.2*cm, 2*cm, 2*cm, 1.2*cm, 3.5*cm, 3*cm]
    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
        ("FONTSIZE",   (0, 0), (-1, 0), 8),
        ("FONTSIZE",   (0, 1), (-1, -1), 7),
        ("ALIGN",      (0, 0), (-1, -1), "LEFT"),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",       (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F3F4")]),
        ("PADDING",    (0, 0), (-1, -1), 4),
    ]))
    elements.append(table)
    doc.build(elements)

# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("🔍 Статус позиции", callback_data="menu_search")],
        [InlineKeyboardButton("📋 Информация о заказах", callback_data="menu_orders")],
        [InlineKeyboardButton("📄 Скачать PDF отчёт", callback_data="menu_pdf")],
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(
        "👋 *Invoice Bot*\n\nВыбери раздел:",
        parse_mode="Markdown",
        reply_markup=reply_markup,
    )

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "menu_search":
        await query.edit_message_text(
            "🔍 *Поиск позиции*\n\nОтправь номер инвойса или название товара:",
            parse_mode="Markdown",
        )
        context.user_data["mode"] = "search"

    elif data == "menu_orders":
        await show_orders(query)

    elif data == "menu_pdf":
        await send_pdf(query)

    elif data == "back_main":
        keyboard = [
            [InlineKeyboardButton("🔍 Статус позиции", callback_data="menu_search")],
            [InlineKeyboardButton("📋 Информация о заказах", callback_data="menu_orders")],
            [InlineKeyboardButton("📄 Скачать PDF отчёт", callback_data="menu_pdf")],
        ]
        await query.edit_message_text(
            "👋 *Invoice Bot*\n\nВыбери раздел:",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(keyboard),
        )

    elif data.startswith("filter_"):
        status = data.replace("filter_", "")
        await show_orders(query, status_filter=status)

async def show_orders(query, status_filter: str = None):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    if status_filter and status_filter != "all":
        cur.execute(
            "SELECT invoice_no, item_name, quantity, unit, total_price, currency, status, notes FROM invoices WHERE status=? ORDER BY invoice_no",
            (status_filter,),
        )
    else:
        cur.execute(
            "SELECT invoice_no, item_name, quantity, unit, total_price, currency, status, notes FROM invoices ORDER BY status, invoice_no"
        )
    rows = cur.fetchall()
    conn.close()

    filter_keyboard = [
        [
            InlineKeyboardButton("Все", callback_data="filter_all"),
            InlineKeyboardButton("✅ Готово", callback_data="filter_ready"),
            InlineKeyboardButton("🚚 В пути", callback_data="filter_in_transit"),
        ],
        [
            InlineKeyboardButton("📦 Доставлено", callback_data="filter_delivered"),
            InlineKeyboardButton("⏳ Ожидает", callback_data="filter_pending"),
            InlineKeyboardButton("💰 Оплачено", callback_data="filter_paid"),
        ],
        [InlineKeyboardButton("« Назад", callback_data="back_main")],
    ]

    if not rows:
        text = "📋 *Заказы*\n\nПока нет ни одного инвойса."
    else:
        lines = ["📋 *Заказы*\n"]
        current_status = None
        for inv_no, item, qty, unit, total, ccy, status, notes in rows:
            label = STATUS_LABELS.get(status, status)
            if status != current_status:
                lines.append(f"\n*{label}*")
                current_status = status
            qty_str = f"{qty} {unit}" if qty else ""
            total_str = f"{total} {ccy}" if total else ""
            line = f"  • `{inv_no}` — {item}"
            if qty_str:
                line += f", {qty_str}"
            if total_str:
                line += f", {total_str}"
            if notes:
                line += f"\n    _{notes}_"
            lines.append(line)
        text = "\n".join(lines)

    await query.edit_message_text(
        text,
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(filter_keyboard),
    )

async def send_pdf(query):
    pdf_path = "/tmp/invoice_report.pdf"
    generate_pdf(pdf_path)
    await query.message.reply_document(
        document=open(pdf_path, "rb"),
        filename=f"invoice_report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        caption="📄 Отчёт по инвойсам",
    )
    await query.edit_message_text("✅ PDF отчёт отправлен.")

async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mode = context.user_data.get("mode")
    text = update.message.text.strip()

    if mode == "search":
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute(
            """SELECT invoice_no, supplier, item_name, quantity, unit, total_price, currency, status, notes
               FROM invoices
               WHERE invoice_no LIKE ? OR item_name LIKE ?
               ORDER BY invoice_no""",
            (f"%{text}%", f"%{text}%"),
        )
        rows = cur.fetchall()
        conn.close()
        context.user_data.pop("mode", None)

        keyboard = [[InlineKeyboardButton("« В меню", callback_data="back_main")]]
        if not rows:
            await update.message.reply_text(
                f"❌ По запросу «{text}» ничего не найдено.",
                reply_markup=InlineKeyboardMarkup(keyboard),
            )
        else:
            lines = [f"🔍 *Результаты по «{text}»:*\n"]
            for inv_no, supplier, item, qty, unit, total, ccy, status, notes in rows:
                label = STATUS_LABELS.get(status, status)
                lines.append(f"*{item}*")
                lines.append(f"  Инвойс: `{inv_no}`")
                if supplier:
                    lines.append(f"  Поставщик: {supplier}")
                if qty:
                    lines.append(f"  Кол-во: {qty} {unit or ''}")
                if total:
                    lines.append(f"  Сумма: {total} {ccy}")
                lines.append(f"  Статус: {label}")
                if notes:
                    lines.append(f"  Заметки: _{notes}_")
                lines.append("")
            await update.message.reply_text(
                "\n".join(lines),
                parse_mode="Markdown",
                reply_markup=InlineKeyboardMarkup(keyboard),
            )
    else:
        keyboard = [
            [InlineKeyboardButton("🔍 Статус позиции", callback_data="menu_search")],
            [InlineKeyboardButton("📋 Информация о заказах", callback_data="menu_orders")],
            [InlineKeyboardButton("📄 Скачать PDF отчёт", callback_data="menu_pdf")],
        ]
        await update.message.reply_text(
            "Используй меню ниже 👇",
            reply_markup=InlineKeyboardMarkup(keyboard),
        )

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    init_db()
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, message_handler))
    logger.info("Bot started")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
