import os
import re
import openpyxl
from flask import Flask, request
import telegram
from telegram.ext import Dispatcher, CommandHandler, MessageHandler, Filters, ConversationHandler

app = Flask(__name__)

# Token & State
TOKEN = os.environ.get('TOKEN', "8445793972:AAEAtlfKNHy4VC5eYgnXtx0RJbJ8i53rjko")
CHOOSING, INPUT_ODP, INPUT_PENYEBAB, INPUT_BA_TEXT = range(4)

# --- FUNGSI PROSES EXCEL ---
def process_ba_excel(raw_text):
    template_path = os.path.join(os.getcwd(), 'assets', 'template.xlsx')
    data = {}
    lines = raw_text.split('\n')
    for line in lines:
        if ':' in line:
            key, val = line.split(':', 1)
            data[key.strip().upper()] = val.strip()

    project_name = data.get("LOKASI", "Tanpa_Project")
    safe_name = re.sub(r'[\\/*?:"<>|]', "_", project_name)
    temp_excel = f"/tmp/BA_{safe_name}.xlsx"

    materials = re.findall(r'-\s*(.*?)\s*=\s*(\d+)', raw_text)
    wb = openpyxl.load_workbook(template_path)
    ws_ba = wb["BA"]
    
    # Isi Data
    ws_ba['D12'] = data.get("WH", "")
    ws_ba['D13'] = data.get("TGL", "")
    ws_ba['D15'] = data.get("LOKASI", "")
    ws_ba['D16'] = data.get("MITRA", "")

    for i, (nama, qty) in enumerate(materials[:12]):
        row = 21 + i
        ws_ba[f'B{row}'] = nama
        ws_ba[f'E{row}'] = int(qty)
        ws_ba[f'F{row}'] = int(qty)

    if "code gudang" in wb.sheetnames:
        wb.remove(wb["code gudang"])
    wb.save(temp_excel)
    return temp_excel

# --- HANDLERS ---
def start(update, context):
    reply_keyboard = [['Input Gamas', 'BA Manual']]
    update.message.reply_text("Pilih fitur:", 
        reply_markup=telegram.ReplyKeyboardMarkup(reply_keyboard, one_time_keyboard=True, resize_keyboard=True))
    return CHOOSING

def start_ba(update, context):
    update.message.reply_text("Kirim data material:")
    return INPUT_BA_TEXT

def get_ba_text(update, context):
    raw_text = update.message.text
    update.message.reply_text("⏳ Memproses Excel...")
    try:
        temp_file = process_ba_excel(raw_text)
        with open(temp_file, 'rb') as f:
            update.message.reply_document(document=f)
        os.remove(temp_file)
    except Exception as e:
        update.message.reply_text(f"❌ Error: {str(e)}")
    return ConversationHandler.END

# (Tambahkan fungsi Gamas sederhana)
def start_gamas(update, context):
    update.message.reply_text("Datek Terdampak (ODP):")
    return INPUT_ODP

def get_odp(update, context):
    context.user_data['odp'] = update.message.text
    update.message.reply_text("Penyebab:")
    return INPUT_PENYEBAB

def get_penyebab(update, context):
    update.message.reply_text(f"Gamas Terinput: {context.user_data['odp']} - {update.message.text}")
    return ConversationHandler.END

# --- VERCEL SETUP ---
bot = telegram.Bot(token=TOKEN)
dispatcher = Dispatcher(bot, None, workers=0, use_context=True)

conv_handler = ConversationHandler(
    entry_points=[CommandHandler("start", start), 
                  MessageHandler(Filters.regex('^Input Gamas$'), start_gamas),
                  MessageHandler(Filters.regex('^BA Manual$'), start_ba)],
    states={
        CHOOSING: [MessageHandler(Filters.regex('^Input Gamas$'), start_gamas),
                   MessageHandler(Filters.regex('^BA Manual$'), start_ba)],
        INPUT_ODP: [MessageHandler(Filters.text, get_odp)],
        INPUT_PENYEBAB: [MessageHandler(Filters.text, get_penyebab)],
        INPUT_BA_TEXT: [MessageHandler(Filters.text, get_ba_text)],
    },
    fallbacks=[CommandHandler("cancel", lambda u, c: ConversationHandler.END)],
)
dispatcher.add_handler(conv_handler)

@app.route('/webhook', methods=['POST'])
def webhook():
    update = telegram.Update.de_json(request.get_json(force=True), bot)
    dispatcher.process_update(update)
    return 'ok'

@app.route('/')
def index():
    return "Bot is running..."