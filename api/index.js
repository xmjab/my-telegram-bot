const axios = require('axios');
const ExcelJS = require('exceljs');
const path = require('path');
const FormData = require('form-data');

const TOKEN = process.env.TOKEN;
const API_URL = `https://api.telegram.org/bot${TOKEN}`;

module.exports = async (req, res) => {
    if (req.method !== 'POST') return res.status(200).send('Bot Active');
    const { message } = req.body;
    if (!message || !message.text) return res.status(200).send('ok');

    const chatId = message.chat.id;
    const text = message.text;

    try {
        if (text === '/start') {
            await axios.post(`${API_URL}/sendMessage`, {
                chat_id: chatId,
                text: "Hailurrdeeee! Kirim data WH: untuk mengisi template Excel asli Anda."
            });
        } 
        else if (text.toUpperCase().includes('WH:')) {
            await handleExcelTemplate(chatId, text);
        }
    } catch (e) { console.error(e); }
    return res.status(200).send('ok');
};

async function handleExcelTemplate(chatId, text) {
    try {
        // 1. Parsing Data
        const data = {};
        text.split('\n').forEach(line => {
            if (line.includes(':')) {
                const [key, ...val] = line.split(':');
                data[key.trim().toUpperCase()] = val.join(':').trim();
            }
        });

        // 2. Load Template Asli
        const templatePath = path.join(process.cwd(), 'assets', 'template.xlsx');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);
        
        const wsBA = workbook.getWorksheet('BA');
        const wsCode = workbook.getWorksheet('code gudang');

        // 3. Logika VLOOKUP Gudang
        let idGudang = data["WH"] || "";
        wsCode.eachRow((row) => {
            if (row.getCell(2).value?.toString().toUpperCase() === idGudang.toUpperCase()) {
                idGudang = `${row.getCell(1).value} ${row.getCell(2).value}`;
            }
        });

        // 4. Isi Sel Spesifik (Sama dengan Logika Python Anda)
        wsBA.getCell('D12').value = idGudang;
        wsBA.getCell('D13').value = data["TGL"] || "";
        wsBA.getCell('D15').value = data["LOKASI"] || "";
        wsBA.getCell('D16').value = data["MITRA"] || "";

        // 5. Isi Tabel Material (B21:F32)
        const materialRegex = /-\s*(.*?)\s*=\s*(\d+)/g;
        let match, i = 0;
        while ((match = materialRegex.exec(text)) && i < 12) {
            const rowNum = 21 + i;
            const nm = match[1];
            const qty = parseInt(match[2]);
            
            let sat = "Pcs";
            if (nm.toUpperCase().includes("AC-OF-SM-ADSS")) sat = "Meter";
            else if (nm.toUpperCase().includes("PU-S7.0-400NM") || nm.toUpperCase().includes("PU-S9.0-140")) sat = "Batang";

            wsBA.getCell(`B${rowNum}`).value = nm;
            wsBA.getCell(`D${rowNum}`).value = sat;
            wsBA.getCell(`E${rowNum}`).value = qty;
            wsBA.getCell(`F${rowNum}`).value = qty;
            i++;
        }

        // 6. Hapus Sheet Code Gudang agar bersih
        workbook.removeWorksheet('code gudang');

        // 7. Generate Buffer & Kirim (Tetap dalam format .xlsx agar gambar aman)
        const buffer = await workbook.xlsx.writeBuffer();
        const form = new FormData();
        form.append('chat_id', chatId);
        form.append('document', buffer, { filename: `BA_${data.LOKASI || 'Manual'}.xlsx` });

        await axios.post(`${API_URL}/sendDocument`, form, { headers: form.getHeaders() });

    } catch (e) {
        await axios.post(`${API_URL}/sendMessage`, { chat_id: chatId, text: "❌ Error: " + e.message });
    }
}