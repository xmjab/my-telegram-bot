import os
import re
import requests
import openpyxl
from flask import Flask, request
from fpdf import FPDF

app = Flask(__name__)
TOKEN = os.environ.get('TOKEN')

def bot_do(method, data, files=None):
    url = f"https://api.telegram.org/bot{TOKEN}/{method}"
    return requests.post(url, data=data, files=files) if files else requests.post(url, json=data)

# --- FITUR 1: LOGIKA GAMAS ---
def handle_gamas(cid, txt):
    odc = re.sub(r'\d', '', txt.replace("ODP", "ODC").replace("/", ""))
    try: sto = odc.split("-")[1][:3]
    except: sto = "..."
    
    res = (f"#request\nSTO : {sto}\nDatek Terdampak (ODC): {odc}\n"
           f"Datek Terdampak (ODP): {txt}\nKeterangan: ODC LOSS\n"
           f"Segmentasi: ODC\nPenyebab: Sedang Dicek\nEstimasi: 1 hari\n"
           f"PIC: yusuf 08112229796\nClassification: Z_NEW_MANUAL_GAMAS_002_004_001_004")
    bot_do("sendMessage", {"chat_id": cid, "text": res})

# --- FITUR 2: LOGIKA BA & PDF ---
def handle_ba(cid, txt):
    try:
        # 1. Parsing & VLOOKUP Gudang
        d = {k.strip().upper(): v.strip() for k, v in [l.split(':', 1) for l in txt.split('\n') if ':' in l]}
        wh_id = d.get("WH", "")
        tp = os.path.join(os.getcwd(), 'assets', 'template.xlsx')
        
        if os.path.exists(tp):
            wb = openpyxl.load_workbook(tp)
            if "code gudang" in wb.sheetnames:
                for r in wb["code gudang"].iter_rows(min_row=1, max_col=2):
                    if r[1].value and r[1].value.strip().upper() == wh_id.upper():
                        wh_id = f"{r[0].value} {r[1].value}"
                        break

        # 2. Generate PDF (Logika Tampilan)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(190, 10, "BERITA ACARA PENGAMBILAN MATERIAL", ln=True, align='C')
        pdf.ln(5)
        pdf.set_font("Arial", size=10)
        pdf.cell(100, 7, f"ID GUDANG : {wh_id}", ln=True)
        pdf.cell(100, 7, f"TANGGAL   : {d.get('TGL', '-')}", ln=True)
        pdf.cell(100, 7, f"LOKASI    : {d.get('LOKASI', '-')}", ln=True)
        pdf.cell(100, 7, f"MITRA     : {d.get('MITRA', '-')}", ln=True)
        pdf.ln(5)

        # Header Tabel & Material (Regex Satuan)
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(10, 8, "No", 1); pdf.cell(110, 8, "Nama Material", 1); pdf.cell(25, 8, "Satuan", 1); pdf.cell(20, 8, "Qty", 1); pdf.ln()
        pdf.set_font("Arial", size=9)
        mats = re.findall(r'-\s*(.*?)\s*=\s*(\d+)', txt)
        for i, (nm, qty) in enumerate(mats[:12]):
            u = nm.upper()
            sat = "Meter" if "AC-OF-SM-ADSS" in u else ("Batang" if any(x in u for x in ["PU-S7.0-400NM", "PU-S9.0-140"]) else "Pcs")
            pdf.cell(10, 8, str(i+1), 1); pdf.cell(110, 8, nm[:55], 1); pdf.cell(25, 8, sat, 1); pdf.cell(20, 8, qty, 1); pdf.ln()

        # 3. Output & Kirim
        path = f"/tmp/BA_{d.get('LOKASI', 'Manual')}.pdf".replace(" ", "_")
        pdf.output(path)
        with open(path, 'rb') as f:
            bot_do("sendDocument", {"chat_id": cid, "caption": "✅ PDF Berhasil Dibuat"}, files={"document": f})
        os.remove(path)
    except Exception as e:
        bot_do("sendMessage", {"chat_id": cid, "text": f"❌ Error: {str(e)}"})

@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.get_json()
    if "message" in data and "text" in data["message"]:
        cid, txt = data["message"]["chat"]["id"], data["message"]["text"]
        if txt == "/start":
            bot_do("sendMessage", {"chat_id": cid, "text": "Hailurrdeeee! Bot Aktif.\n- Ketik ODP... (Gamas)\n- Ketik WH:... (BA Manual)"})
        elif txt.upper().startswith("ODP"):
            handle_gamas(cid, txt)
        elif "WH:" in txt.upper():
            bot_do("sendMessage", {"chat_id": cid, "text": "⏳ Memproses PDF..."})
            handle_ba(cid, txt)
    return "ok", 200

@app.route('/')
def index(): return "Bot Online", 200