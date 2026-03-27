import os
import re
import requests
import openpyxl
from flask import Flask, request
from fpdf import FPDF

app = Flask(__name__)
TOKEN = os.environ.get('TOKEN')

def send_telegram(method, payload, files=None):
    url = f"https://api.telegram.org/bot{TOKEN}/{method}"
    if files:
        return requests.post(url, data=payload, files=files)
    return requests.post(url, json=payload)

# --- LOGIKA FITUR 1: GAMAS ---
def proses_gamas(chat_id, text):
    odp_raw = text
    odc = re.sub(r'\d', '', odp_raw.replace("ODP", "ODC").replace("/", ""))
    try: sto = odc.split("-")[1][:3]
    except: sto = "..."
    
    output = (
        f"#request\nSTO : {sto}\nDatek Terdampak (ODC): {odc}\n"
        f"Datek Terdampak (ODP): {odp_raw}\nKeterangan: ODC LOSS\n"
        f"Segmentasi: ODC\nPenyebab: Sedang Dicek\nEstimasi: 1 hari\n"
        f"PIC: yusuf 08112229796\nClassification: Z_NEW_MANUAL_GAMAS_002_004_001_004"
    )
    send_telegram("sendMessage", {"chat_id": chat_id, "text": output})

# --- LOGIKA FITUR 2: BA MANUAL (KONVERSI PDF) ---
def proses_ba_pdf(chat_id, raw_text):
    try:
        # 1. Parsing Data Dasar
        data = {k.strip().upper(): v.strip() for k, v in [line.split(':', 1) for line in raw_text.split('\n') if ':' in line]}
        
        # 2. Logika VLOOKUP Gudang (Sesuai Code Colab Anda)
        id_gudang_val = data.get("WH", "")
        tp = os.path.join(os.getcwd(), 'assets', 'template.xlsx')
        if os.path.exists(tp):
            wb = openpyxl.load_workbook(tp)
            if "code gudang" in wb.sheetnames:
                for r in wb["code gudang"].iter_rows(min_row=1, max_col=2):
                    if r[1].value and r[1].value.strip().upper() == id_gudang_val.upper():
                        id_gudang_val = f"{r[0].value} {r[1].value}"
                        break

        # 3. Create PDF (Menggantikan LibreOffice)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(190, 10, "BERITA ACARA PENGAMBILAN MATERIAL", ln=True, align='C')
        pdf.ln(5)
        
        pdf.set_font("Arial", size=10)
        pdf.cell(100, 7, f"ID GUDANG : {id_gudang_val}", ln=True)
        pdf.cell(100, 7, f"TANGGAL   : {data.get('TGL', '-')}", ln=True)
        pdf.cell(100, 7, f"LOKASI    : {data.get('LOKASI', '-')}", ln=True)
        pdf.cell(100, 7, f"MITRA     : {data.get('MITRA', '-')}", ln=True)
        pdf.ln(5)

        # Tabel Header
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(10, 8, "No", 1); pdf.cell(110, 8, "Nama Material", 1); pdf.cell(25, 8, "Satuan", 1); pdf.cell(20, 8, "Qty", 1); pdf.ln()

        # Isi Material (Regex & Penentuan Satuan sesuai code asli)
        pdf.set_font("Arial", size=9)
        mats = re.findall(r'-\s*(.*?)\s*=\s*(\d+)', raw_text)
        for i, (nm, qty) in enumerate(mats[:12]):
            nm_u = nm.upper()
            sat = "Meter" if "AC-OF-SM-ADSS" in nm_u else ("Batang" if any(x in nm_u for x in ["PU-S7.0-400NM", "PU-S9.0-140"]) else "Pcs")
            pdf.cell(10, 8, str(i+1), 1); pdf.cell(110, 8, nm[:55], 1); pdf.cell(25, 8, sat, 1); pdf.cell(20, 8, qty, 1); pdf.ln()

        # Simpan & Kirim
        out_name = f"BA_{data.get('LOKASI', 'Manual')}.pdf".replace(" ", "_")
        pdf_path = f"/tmp/{out_name}"
        pdf.output(pdf_path)
        
        with open(pdf_path, 'rb') as f:
            send_telegram("sendDocument", {"chat_id": chat_id, "caption": "✅ PDF Berhasil Dibuat"}, files={"document": f})
        os.remove(pdf_path)

    except Exception as e:
        send_telegram("sendMessage", {"chat_id": chat_id, "text": f