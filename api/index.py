from flask import Flask, request

app = Flask(__name__)

@app.route('/webhook', methods=['POST'])
def webhook():
    return 'OK', 200

@app.route('/')
def index():
    return "TEST BERHASIL: Server Vercel Berjalan!", 200