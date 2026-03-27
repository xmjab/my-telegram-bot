import os
import requests
import openpyxl
import re
from flask import Flask, request

app = Flask(__name__)
TOKEN = os.environ.get('TOKEN')

def send_message(chat_id, text):
    requests.post(f"https://api.telegram.org/bot{TOKEN}/sendMessage", 
                  json={"chat_id": chat_id, "text": text})

def send_document(chat_id, file_path):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    with open(file_path, 'rb') as f:
        requests.post(url, data={"chat_id": chat_id}, files={"document": f})

def process_excel(chat_id, raw_text):
    try:
        # Path ke template di folder assets
        template_path = os.path.join(os.getcwd(), 'assets', 'template.xlsx')
        if not os.path.exists(template_path):
            send_message(chat_id, "❌ Error: File assets/template.xlsx tidak ditemukan.")
            return

        # Ambil data dari teks (Logika lama kamu)
        data = {}
        for line in raw_text.split('\n'):
            if ':' in line:
                key, val = line.split(':', 1)
                data[key.strip().upper()] = val.strip()

        # Buka dan isi Excel
        wb = openpyxl.load_workbook(template_path)
        ws = wb["BA"]
        ws['D12'] = data.get("WH", "")
        ws['D13'] = data.get("TGL", "")
        ws['D15'] = data.get("LOKASI", "")
        ws['D16'] = data.get("MITRA", "")

        # Simpan sementara di /tmp (folder wajib untuk Vercel)
        output_path = "/tmp/BA_Manual.xlsx"
        wb.save(output_path)
        
        # Kirim ke user
        send_document(chat_id, output_path)
        os.remove(output_path) # Hapus setelah dikirim
        
    except Exception as e:
        send_message(chat_id, f"❌ Terjadi kesalahan saat proses Excel: {str(e)}")

@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.get_json()
    if "message" in data and "text" in data["message"]:
        chat_id = data["message"]["chat"]["id"]
        text = data["message"]["text"]

        if text == "/start":
            send_message(chat_id, "✅ Bot Aktif!\n\nKirim data dengan format:\nWH: (Isi)\nTGL: (Isi)\nLOKASI: (Isi)\nMITRA: (Isi)")
        elif "WH:" in text.upper():
            send_message(chat_id, "⏳ Sedang memproses BA Manual Anda...")
            process_excel(chat_id, text)
        else:
            send_message(chat_id, f"Bot menerima: {text}\n(Gunakan format WH: untuk proses Excel)")

    return "ok", 200

@app.route('/')
def index():
    return "Bot is Running", 200