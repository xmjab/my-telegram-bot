import os
import re
import openpyxl
from flask import Flask, request
import telegram

app = Flask(__name__)

# Ambil TOKEN dari Environment Variables Vercel
TOKEN = os.environ.get('TOKEN')
bot = telegram.Bot(token=TOKEN)

@app.route('/webhook', methods=['POST'])
def webhook():
    if request.method == "POST":
        try:
            update = telegram.Update.de_json(request.get_json(force=True), bot)
            if not update.message:
                return 'ok'
                
            chat_id = update.message.chat.id
            text = update.message.text

            if text == "/start":
                bot.send_message(chat_id=chat_id, text="Bot Aktif! Kirim data material dengan format 'WH: ...' untuk mencoba.")
            
            # LOGIKA PROSES EXCEL
            elif "WH:" in text.upper():
                bot.send_message(chat_id=chat_id, text="⏳ Sedang memproses file Excel...")
                
                # Path template (Pastikan file template.xlsx ada di folder assets)
                template_path = os.path.join(os.getcwd(), 'assets', 'template.xlsx')
                
                if not os.path.exists(template_path):
                    bot.send_message(chat_id=chat_id, text="❌ Error: File assets/template.xlsx tidak ditemukan di server.")
                    return 'ok'

                wb = openpyxl.load_workbook(template_path)
                ws = wb["BA"]
                
                # Contoh: Isi data sederhana dari baris pertama input user ke sel D15
                lines = text.split('\n')
                ws['D15'] = lines[0] 
                
                # Simpan di folder sementara Vercel
                temp_file = "/tmp/Hasil_BA.xlsx"
                wb.save(temp_file)
                
                # Kirim ke user
                with open(temp_file, 'rb') as f:
                    bot.send_document(chat_id=chat_id, document=f, filename="Hasil_BA_Manual.xlsx")
                
                os.remove(temp_file)
            
            else:
                bot.send_message(chat_id=chat_id, text=f"Anda mengirim: {text}")

        except Exception as e:
            # Jika ada error di tengah jalan, bot akan lapor ke Anda
            print(f"Error: {e}")
            
        return 'ok'

@app.route('/')
def index():
    return "TEST BERHASIL: Server Vercel Berjalan!"