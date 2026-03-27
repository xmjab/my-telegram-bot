const axios = require('axios');
const ExcelJS = require('exceljs');
const { jsPDF } = require('jspdf');
require('jspdf-autotable');
const FormData = require('form-data');
const path = require('path');

const TOKEN = process.env.TOKEN;
const API_URL = `https://api.telegram.org/bot${TOKEN}`;

module.exports = async (req, res) => {
    if (req.method !== 'POST') return res.status(200).send('Bot Node.js Online');

    const { message } = req.body;
    if (!message || !message.text) return res.status(200).send('ok');

    const chatId = message.chat.id;
    const text = message.text;

    try {
        if (text === '/start') {
            await axios.post(`${API_URL}/sendMessage`, {
                chat_id: chatId,
                text: "Hailurrdeeee! Bot Node.js Aktif.\n\n- Ketik ODP... (Gamas)\n- Ketik WH:... (BA Manual)"
            });
        } 
        else if (text.toUpperCase().startsWith('ODP')) {
            await handleGamas(chatId, text);
        } 
        else if (text.toUpperCase().includes('WH:')) {
            await handleBA(chatId, text);
        }
    } catch (error) {
        console.error("Error utama:", error.message);
    }

    return res.status(200).send('ok');
};

async function handleGamas(chatId, text) {
    const odc = text.replace(/ODP/g, 'ODC').replace(/\//g, '').replace(/\d/g, '');
    const sto = odc.split('-')[1]?.substring(0, 3) || '...';
    const output = `#request\nSTO : ${sto}\nDatek Terdampak (ODC): ${odc}\nDatek Terdampak (ODP): ${text}\nKeterangan: ODC LOSS\nSegmentasi: ODC\nPenyebab: Sedang Dicek\nEstimasi: 1 hari\nPIC: yusuf 08112229796\nClassification: Z_NEW_MANUAL_GAMAS_002_004_001_004`;
    await axios.post(`${API_URL}/sendMessage`, { chat_id: chatId, text: output });
}

async function handleBA(chatId, text) {
    try {
        // 1. Parsing Data
        const data = {};
        text.split('\n').forEach(line => {
            if (line.includes(':')) {
                const [key, ...val] = line.split(':');
                data[key.trim().toUpperCase()] = val.join(':').trim();
            }
        });

        // 2. Logika VLOOKUP Gudang (ExcelJS)
        let idGudang = data["WH"] || "";
        const templatePath = path.join(process.cwd(), 'assets', 'template.xlsx');
        
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);
        const sheet = workbook.getWorksheet('code gudang');
        
        sheet.eachRow((row) => {
            if (row.getCell(2).value?.toString().toUpperCase() === idGudang.toUpperCase()) {
                idGudang = `${row.getCell(1).value} ${row.getCell(2).value}`;
            }
        });

        // 3. Generate PDF (jsPDF)
        const doc = new jsPDF();
        doc.setFontSize(14);
        doc.text("BERITA ACARA PENGAMBILAN MATERIAL", 105, 20, { align: 'center' });
        
        doc.setFontSize(10);
        doc.text(`ID GUDANG : ${idGudang}`, 20, 40);
        doc.text(`TANGGAL   : ${data.TGL || '-'}`, 20, 47);
        doc.text(`LOKASI    : ${data.LOKASI || '-'}`, 20, 54);
        doc.text(`MITRA     : ${data.MITRA || '-'}`, 20, 61);

        // 4. Extract Materials & Logic Satuan
        const tableRows = [];
        const materialRegex = /-\s*(.*?)\s*=\s*(\d+)/g;
        let match, i = 1;
        while ((match = materialRegex.exec(text)) !== null) {
            const nm = match[1].toUpperCase();
            let sat = "Pcs";
            if (nm.includes("AC-OF-SM-ADSS")) sat = "Meter";
            else if (nm.includes("PU-S7.0-400NM") || nm.includes("PU-S9.0-140")) sat = "Batang";
            
            tableRows.push([i++, match[1], sat, match[2]]);
        }

        doc.autoTable({
            startY: 70,
            head: [['No', 'Nama Material', 'Satuan', 'Qty']],
            body: tableRows,
            theme: 'grid'
        });

        // 5. Kirim ke Telegram
        const pdfData = doc.output('arraybuffer');
        const form = new FormData();
        form.append('chat_id', chatId);
        form.append('document', Buffer.from(pdfData), { filename: `BA_${data.LOKASI || 'Manual'}.pdf` });

        await axios.post(`${API_URL}/sendDocument`, form, { headers: form.getHeaders() });

    } catch (e) {
        await axios.post(`${API_URL}/sendMessage`, { chat_id: chatId, text: "❌ Error: " + e.message });
    }
}