import os
import requests
from flask import Flask, request

app = Flask(__name__)

# Ambil Token dari Environment Variable Vercel
TOKEN = os.environ.get('TOKEN')

def kirim_pesan(chat_id, teks):
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    payload = {"chat_id": chat_id, "text": teks}
    requests.post(url, json=payload)

@app.route('/webhook', methods=['POST'])
def webhook():
    try:
        data = request.get_json()
        if "message" in data:
            chat_id = data["message"]["chat"]["id"]
            text = data["message"].get("text", "")

            if text == "/start":
                kirim_pesan(chat_id, "✅ AKHIRNYA JALAN!\nIni adalah mode tanpa library. Bot sudah bisa mendengar kamu.")
            else:
                kirim_pesan(chat_id, f"Kamu mengirim: {text}")
    except Exception as e:
        print(f"Error: {e}")
        
    return "ok", 200

@app.route('/')
def index():
    return "Server Bot Aktif!", 200