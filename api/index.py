import os
import re
import openpyxl
import asyncio
from flask import Flask, request
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler, Application
)

app = Flask(__name__)

# Token dari Environment Variable Vercel (Lebih Aman)
TOKEN = os.environ.get('TOKEN', "8445793972:AAEAtlfKNHy4VC5eYgnXtx0RJbJ8i53rjko")

# Definisi State
CHOOSING, INPUT_ODP, INPUT_PENYEBAB, INPUT_BA_TEXT = range(4)

# --- FUNGSI PROSES EXCEL (DISESUAIKAN UNTUK VERCEL) ---
def process_ba_excel(raw_text):
    # Path template di Vercel
    template_path = os.path.join(os.getcwd(), 'assets', 'template.xlsx')
    
    data = {}
    lines = raw_text.split('\n')
    for line in lines:
        if ':' in line:
            key, val = line.split(':', 1)
            data[key.strip().upper()] = val.strip()

    project_name = data.get("LOKASI", "Tanpa_Project")
    safe_project_name = re.sub(r'[\\/*?:"<>|]', "_", project_name)
    
    # Simpan di folder /tmp (Satu-satunya folder yang bisa ditulis di Vercel)
    temp_excel = f"/tmp/BA_Manual_{safe_project_name}.xlsx"

    materials = re.findall(r'-\s*(.*?)\s*=\s*(\d+)', raw_text)

    wb = openpyxl.load_workbook(template_path)
    ws_ba = wb["BA"]
    ws_code = wb["code gudang"]

    wh_user = data.get("WH", "")
    id_gudang_val = ""
    for row in ws_code.iter_rows(min_row=1, max_col=2):
        if row[1].value and row[1].value.strip().upper() == wh_user.upper():
            id_gudang_val = f"{row[0].value} {row[1].value}"
            break
    
    ws_ba['D12'] = id_gudang_val
    ws_ba['D13'] = data.get("TGL", "")
    ws_ba['D15'] = data.get("LOKASI", "")
    ws_ba['D16'] = data.get("MITRA", "")

    start_row = 21
    for i, (nama, qty) in enumerate(materials):
        if i > 11: break
        current_row = start_row + i
        ws_ba[f'B{current_row}'] = nama
        satuan = "Meter" if "AC-OF-SM-ADSS" in nama.upper() else "Pcs"
        ws_ba[f'D{current_row}'] = satuan
        ws_ba[f'E{current_row}'] = int(qty)
        ws_ba[f'F{current_row}'] = int(qty)

    if "code gudang" in wb.sheetnames:
        wb.remove(wb["code gudang"])

    wb.save(temp_excel)
    return temp_excel

# --- HANDLERS ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    reply_keyboard = [['Input Gamas', 'BA Manual']]
    await update.message.reply_text(
        "Hailurrdeeee! Pilih fitur:",
        reply_markup=ReplyKeyboardMarkup(reply_keyboard, one_time_keyboard=True, resize_keyboard=True),
    )
    return CHOOSING

async def start_ba(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Kirimkan data ringkasan material Anda:", reply_markup=ReplyKeyboardRemove())
    return INPUT_BA_TEXT

async def get_ba_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw_text = update.message.text
    await update.message.reply_text("⏳ Sedang memproses file Excel...")
    try:
        temp_file = process_ba_excel(raw_text)
        with open(temp_file, 'rb') as doc:
            await update.message.reply_document(document=doc)
        os.remove(temp_file)
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {str(e)}")
    return ConversationHandler.END

# (Tambahkan fungsi start_gamas, get_odp, get_penyebab, cancel di sini sesuai kode asli Anda)

# --- VERCEL WEBHOOK SETUP ---
application = ApplicationBuilder().token(TOKEN).build()

conv_handler = ConversationHandler(
    entry_points=[
        CommandHandler("start", start),
        MessageHandler(filters.Regex('^Input Gamas$'), start_gamas),
        MessageHandler(filters.Regex('^BA Manual$'), start_ba)
    ],
    states={
        CHOOSING: [
            MessageHandler(filters.Regex('^Input Gamas$'), start_gamas),
            MessageHandler(filters.Regex('^BA Manual$'), start_ba)
        ],
        INPUT_ODP: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_odp)],
        INPUT_PENYEBAB: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_penyebab)],
        INPUT_BA_TEXT: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_ba_text)],
    },
    fallbacks=[CommandHandler("cancel", cancel)],
)
application.add_handler(conv_handler)

@app.route('/webhook', methods=['POST'])
async def webhook():
    if request.method == "POST":
        async with application:
            update = Update.de_json(request.get_json(force=True), application.bot)
            await application.process_update(update)
        return "ok"

@app.route('/')
def index():
    return "Bot is running"