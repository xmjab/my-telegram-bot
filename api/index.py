import os
import re
import requests
import openpyxl
from flask import Flask, request

app = Flask(__name__)
TOKEN = os.environ.get('TOKEN')

def send_message(chat_id, text):
    requests.post(f"https://api.telegram.org/bot{TOKEN}/sendMessage", 
                  json={"chat_id": chat_id, "text": text})

def send_document(chat_id, file_path, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    with open(file_path, 'rb') as f:
        requests.post(url, data={"chat_id": chat_id, "caption": caption}, files={"document": f})

def process_ba_excel(chat_id, raw_text):
    try:
        template_path = os.path.join(os.getcwd(), 'assets', 'template.xlsx')
        
        # Ekstraksi Data (Sesuai Logika Asli Anda)
        data = {}
        lines = raw_text.split('\n')
        for line in lines:
            if ':' in line:
                key, val = line.split(':', 1)
                data[key.strip().upper()] = val.strip()

        # Ambil Nama Project untuk Nama File
        project_name = data.get("LOKASI", "Tanpa_Project")
        safe_name = re.sub(r'[\\/*?:"<>|]', "_", project_name)
        output_path = f"/tmp/BA_{safe_name}.xlsx"

        # Proses Excel (Sesuai Logika Asli Anda)
        wb = openpyxl.load_workbook(template_path)
        ws_ba = wb["BA"]
        
        # Isi Header
        ws_ba['D12'] = data.get("WH", "")
        ws_ba['D13'] = data.get("TGL", "")
        ws_ba['D15'] = data.get("LOKASI", "")
        ws_ba['D16'] = data.get("MITRA", "")

        # Isi Material (Sesuai Logika Asli Anda)
        materials = re.findall(r'-\s*(.*?)\s*=\s*(\d+)', raw_text)
        for i, (nama, qty) in enumerate(materials[:12]):
            row = 21 + i
            ws_ba[f'B{row}'] = nama
            ws_ba[f'E{row}'] = int(qty)
            ws_ba[f'F{row}'] = int(qty)

        # Hapus sheet tambahan jika ada
        if "code gudang" in wb.sheetnames:
            wb.remove(wb["code gudang"])
            
        wb.save(output_path)
        
        # Kirim balik ke user
        send_document(chat_id, output_path, f"✅ BA Manual untuk {project_name} berhasil diproses.")
        os.remove(output_path)
        
    except Exception as e:
        send_message(chat_id, f"❌ Error saat memproses Excel: {str(e)}")

@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.get_json()
    if "message" in data and "text" in data["message"]:
        chat_id = data["message"]["chat"]["id"]
        text = data["message"]["text"]

        if text == "/start":
            send_message(chat_id, "Bot Aktif! Kirim data material Anda untuk memproses BA Manual.")
        elif "WH:" in text.upper():
            send_message(chat_id, "⏳ Memproses data...")
            process_ba_excel(chat_id, text)
            
    return "ok", 200

@app.route('/')
def index():
    return "Bot is Running", 200