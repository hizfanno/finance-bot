from telegram.ext import ApplicationBuilder, CommandHandler
from openpyxl import load_workbook
from dotenv import load_dotenv
from datetime import datetime
import os

import gspread
from google.oauth2.service_account import Credentials


load_dotenv("config.env")

TOKEN = os.getenv("BOT_TOKEN")
ALLOWED_USERS = os.getenv("ALLOWED_USERS").split(",")
OWNER = os.getenv("OWNER_NAME")
EXCEL_FILE = os.getenv("EXCEL_FILE")

print("TOKEN:", TOKEN)

def is_allowed(update):
    user_id = str(update.effective_user.id)
    return user_id in ALLOWED_USERS


def save_to_excel(category, amount):

    wb = load_workbook(EXCEL_FILE)
    ws = wb["Pengeluaran Bulanan"]

    row = 2

    while ws[f"A{row}"].value is not None:
        row += 1

    ws[f"A{row}"] = datetime.now().strftime("%d-%m-%Y")
    ws[f"B{row}"] = OWNER
    ws[f"C{row}"] = category
    ws[f"D{row}"] = 1
    ws[f"E{row}"] = amount * 1000

    wb.save(EXCEL_FILE)

def get_sheet():

    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_file(
        "financebot-key.json",
        scopes=scope
    )

    client = gspread.authorize(creds)

    return client.open("Daily Finance").worksheet("Pengeluaran")

def save_to_google_sheet(category, amount, username):

    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_file(
        "financebot-key.json",
        scopes=scope
    )

    client = gspread.authorize(creds)

    sheet = client.open("Daily Finance").worksheet("Pengeluaran")

    row = [
        datetime.now().strftime("%d-%m-%Y %H:%M"),
        username,
        category,
        1,
        amount * 1000
    ]

    sheet.append_row(
    row,
    table_range="A:E"
)

async def month(update, context):

    if not is_allowed(update):
        return

    username = update.effective_user.first_name

    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_file(
        "financebot-key.json",
        scopes=scope
    )

    client = gspread.authorize(creds)
    sheet = client.open("Daily Finance").worksheet("Pengeluaran")

    rows = sheet.get_all_values()

    current_month = datetime.now().strftime("%m-%Y")

    summary = {}
    total = 0


    for row in rows[1:]:

        try:
            nominal = int(row[4])
        except:
            continue

        tanggal = row[0]
        nama = row[1]
        kategori = row[2]

        if nama != username:
            continue

        if current_month not in tanggal:
            continue

        summary[kategori] = summary.get(kategori, 0) + nominal
        total += nominal


    message = f"📅 {datetime.now().strftime('%B %Y')} — {username}\n\n"

    for kategori, amount in summary.items():
        message += f"{kategori}: Rp {rupiah(amount)}\n"

    message += f"\n💰 Total: Rp {rupiah(total)}"

    await update.message.reply_text(message)

async def today(update, context):

    if not is_allowed(update):
        return

    username = update.effective_user.first_name

    sheet = get_sheet()
    rows = sheet.get_all_values()

    today_date = datetime.now().strftime("%d-%m-%Y")

    summary = {}
    total = 0


    for row in rows[1:]:

        try:
            nominal = int(row[4])
        except:
            continue

        tanggal = row[0]
        nama = row[1]
        kategori = row[2]

        if nama != username:
            continue

        if today_date not in tanggal:
            continue

        summary[kategori] = summary.get(kategori, 0) + nominal
        total += nominal


    message = f"📆 Hari ini — {username}\n\n"

    for kategori, amount in summary.items():
        message += f"{kategori}: Rp {rupiah(amount)}\n"

    message += f"\n💰 Total: Rp {rupiah(total)}"

    await update.message.reply_text(message)

async def top(update, context):

    if not is_allowed(update):
        return

    username = update.effective_user.first_name

    sheet = get_sheet()
    rows = sheet.get_all_values()

    current_month = datetime.now().strftime("%m-%Y")

    summary = {}


    for row in rows[1:]:

        try:
            nominal = int(row[4])
        except:
            continue

        tanggal = row[0]
        nama = row[1]
        kategori = row[2]

        if nama != username:
            continue

        if current_month not in tanggal:
            continue

        summary[kategori] = summary.get(kategori, 0) + nominal


    sorted_items = sorted(
        summary.items(),
        key=lambda x: x[1],
        reverse=True
    )


    message = "🏆 Top Spending\n\n"

    medals = ["🥇", "🥈", "🥉"]


    for i, (kategori, amount) in enumerate(sorted_items[:3]):

        medal = medals[i]

        message += f"{medal} {kategori}: Rp {rupiah(amount)}\n"


    await update.message.reply_text(message)

async def help_command(update, context):

    if not is_allowed(update):
        return

    message = """
📋 Finance Bot Menu

➕ Input Pengeluaran
☕ /kopi 25
🚬 /rokok 36
⛽ /bensin 50
🍚 /makan 20
🧴 /skincare 50
✨ /lifestyle 100
🛍️ /kebutuhan 75
📦 /misc 20

📊 Report
📅 /month    → bulan ini
📆 /today    → hari ini
👤 /me       → data pribadi
🏆 /top      → kategori terbesar

ℹ️ /help     → lihat menu
"""

    await update.message.reply_text(message)

async def summary(update, context):

    if not is_allowed(update):
        return

    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_file(
        "financebot-key.json",
        scopes=scope
    )

    client = gspread.authorize(creds)

    sheet = client.open("Daily Finance").worksheet("Pengeluaran")

    kopi = int(sheet.acell("H2").value or 0)
    rokok = int(sheet.acell("H3").value or 0)
    bensin = int(sheet.acell("H4").value or 0)
    makan = int(sheet.acell("H5").value or 0)
    total = int(sheet.acell("H6").value or 0)

    await update.message.reply_text(
        f"""📊 Total Saat Ini

☕ Kopi: Rp {rupiah(kopi)}
🚬 Rokok: Rp {rupiah(rokok)}
⛽ Bensin: Rp {rupiah(bensin)}
🍜 Makan: Rp {rupiah(makan)}

💰 Total: Rp {rupiah(total)}
"""
    )

async def me(update, context):

    if not is_allowed(update):
        return

    username = update.effective_user.first_name

    sheet = get_sheet()

    data = sheet.get_all_records()

    summary = {}
    total = 0


    for row in data:

        if row["Nama"] != username:
            continue

        kategori = row["Kategori"]
        nominal = int(row["Nominal"])

        summary[kategori] = summary.get(kategori, 0) + nominal

        total += nominal


    emoji_map = {
        "Kopi": "☕",
        "Rokok": "🚬",
        "Bensin": "⛽",
        "Makan": "🍜",
        "Skincare": "🧴",
        "Lifestyle": "✨",
        "Kebutuhan": "🛍️",
        "Misc": "📦"
    }


    message = f"👤 {username}\n\n"


    for kategori, amount in summary.items():

        emoji = emoji_map.get(kategori, "💸")

        message += f"{emoji} {kategori}: Rp {rupiah(amount)}\n"


    message += f"\n💰 Total: Rp {rupiah(total)}"


    await update.message.reply_text(message)


def rupiah(amount):
    return f"{amount:,}".replace(",", ".")


async def add(update, context, category):


    if not is_allowed(update):
        return

    print("ARGS:", context.args)

    try:
        amount = int(context.args[0])
        username = update.effective_user.first_name

        save_to_google_sheet(category, amount, username)

        await update.message.reply_text(
            f"✅ {category} Rp {rupiah(amount * 1000)} tersimpan"
        )

    except Exception as e:
        print("ERROR:", e)


        #await update.message.reply_text(
        #   "Contoh: /kopi 25"
        #)

async def kopi(update, context):
    await add(update, context, "Kopi")


async def rokok(update, context):
    await add(update, context, "Rokok")


async def kebutuhan(update, context):
    await add(update, context, "Kebutuhan")


async def misc(update, context):
    await add(update, context, "Misc")


async def skincare(update, context):
    await add(update, context, "Skincare")


async def lifestyle(update, context):
    await add(update, context, "Lifestyle")


async def makan(update, context):
    await add(update, context, "Makan")


async def bensin(update, context):
    await add(update, context, "Bensin")


app = ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("kopi", kopi))
app.add_handler(CommandHandler("rokok", rokok))
app.add_handler(CommandHandler("makan", makan))
app.add_handler(CommandHandler("bensin", bensin))
app.add_handler(CommandHandler("summary", summary))
app.add_handler(CommandHandler("me", me))
app.add_handler(CommandHandler("month", month))
app.add_handler(CommandHandler("help", help_command))
app.add_handler(CommandHandler("today", today))
app.add_handler(CommandHandler("top", top))
app.add_handler(CommandHandler("skincare", skincare))
app.add_handler(CommandHandler("lifestyle", lifestyle))
app.add_handler(CommandHandler("kebutuhan", kebutuhan))
app.add_handler(CommandHandler("misc", misc))

print("Bot running...")

app.run_polling()
