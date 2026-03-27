import os
import re
import openpyxl
from flask import Flask, request
import telegram

app = Flask(__name__)

# Ambil Token dari Environment Variable Vercel
TOKEN = os.environ.get('TOKEN')
bot = telegram.Bot(token=TOKEN)

@app.route('/webhook', methods=['POST'])
def webhook():
    if request.method == "POST":
        update = telegram.Update.de_json(request.get_json(force=True), bot)
        chat_id = update.message.chat.id
        text = update.message.text

        if text == "/start":
            bot.send_message(chat_id=chat_id, text="Bot Aktif di Vercel! Kirim data material dengan format 'BA Manual' untuk mencoba.")
        
        # Logika sederhana tanpa ConversationHandler agar tidak crash
        elif "WH:" in text.upper():
            bot.send_message(chat_id=chat_id, text="⏳ Sedang memproses data Anda...")
            try:
                # Path template
                template_path = os.path.join(os.getcwd(), 'assets', 'template.xlsx')
                wb = openpyxl.load_workbook(template_path)
                ws = wb["BA"]
                
                # Contoh pengisian simpel (sesuaikan dengan kebutuhanmu)
                ws['D15'] = text.split('\n')[0] # Ambil baris pertama sebagai lokasi
                
                temp_file = "/tmp/output.xlsx"
                wb.save(temp_file)
                
                with open(temp_file, 'rb') as f:
                    bot.send_document(chat_id=chat_id, document=f, filename="Hasil_BA.xlsx")
                os.remove(temp_file)
            except Exception as e:
                bot.send_message(chat_id=chat_id, text=f"❌ Terjadi kesalahan: {str(e)}")
        
        else:
            bot.send_message(chat_id=chat_id, text=f"Anda mengirim: {text}")

        return 'ok'

@app.route('/')
def index():
    return "Telegram Bot is Online"