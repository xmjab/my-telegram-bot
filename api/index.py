import os
import requests
from flask import Flask, request

app = Flask(__name__)

# Ambil TOKEN dari Environment Variables
TOKEN = os.environ.get('TOKEN')
TELEGRAM_API = f"https://api.telegram.org/bot{TOKEN}"

def send_message(chat_id, text):
    url = f"{TELEGRAM_API}/sendMessage"
    payload = {"chat_id": chat_id, "text": text}
    requests.post(url, json=payload)

@app.route('/webhook', methods=['POST'])
def webhook():
    if request.method == "POST":
        data = request.get_json()
        
        # Cek jika ada pesan masuk
        if "message" in data:
            chat_id = data["message"]["chat"]["id"]
            text = data["message"].get("text", "")

            if text == "/start":
                send_message(chat_id, "✅ Bot BERHASIL Aktif di Vercel tanpa library berat! Kirim pesan apapun untuk tes.")
            else:
                send_message(chat_id, f"Bot menerima pesan: {text}")
                
        return "ok", 200

@app.route('/')
def index():
    return "Bot Server is Running", 200