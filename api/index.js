const axios = require('axios');
const ExcelJS = require('exceljs');
const path = require('path');
const FormData = require('form-data');
const CloudConvert = require('cloudconvert');

const TOKEN = process.env.TOKEN;
const CLOUDCONVERT_KEY = process.env.CLOUDCONVERT_API_KEY;
const API_URL = `https://api.telegram.org/bot${TOKEN}`;

// Menghapus spasi putih yang mungkin terbawa dari Environment Variables
const CLOUDCONVERT_KEY = process.env.CLOUDCONVERT_API_KEY ? process.env.CLOUDCONVERT_API_KEY.trim() : "";
const cloudConvert = new CloudConvert(CLOUDCONVERT_KEY);

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
                text: "Hailurrdeeee! Kirim data WH: untuk membuat PDF BA Manual yang sempurna."
            });
        } 
        else if (text.toUpperCase().includes('WH:')) {
            await axios.post(`${API_URL}/sendMessage`, { chat_id: chatId, text: "⏳ Sedang mengisi template & mengonversi ke PDF..." });
            await handlePerfectPDF(chatId, text);
        }
    } catch (e) { console.error(e); }
    return res.status(200).send('ok');
};

async function handlePerfectPDF(chatId, text) {
    try {
        // 1. Parsing Data Input
        const data = {};
        text.split('\n').forEach(line => {
            if (line.includes(':')) {
                const [key, ...val] = line.split(':');
                data[key.trim().toUpperCase()] = val.join(':').trim();
            }
        });

        // 2. Load & Isi Template Excel (Menjaga Gambar Aman)
        const templatePath = path.join(process.cwd(), 'assets', 'template.xlsx');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);
        
        const wsBA = workbook.getWorksheet('BA');
        const wsCode = workbook.getWorksheet('code gudang');

        let idGudang = data["WH"] || "";
        wsCode.eachRow((row) => {
            if (row.getCell(2).value?.toString().toUpperCase() === idGudang.toUpperCase()) {
                idGudang = `${row.getCell(1).value} ${row.getCell(2).value}`;
            }
        });

        wsBA.getCell('D12').value = idGudang;
        wsBA.getCell('D13').value = data["TGL"] || "";
        wsBA.getCell('D15').value = data["LOKASI"] || "";
        wsBA.getCell('D16').value = data["MITRA"] || "";

        const materialRegex = /-\s*(.*?)\s*=\s*(\d+)/g;
        let match, i = 0;
        while ((match = materialRegex.exec(text)) && i < 12) {
            const rowNum = 21 + i;
            const nm = match[1];
            let sat = "Pcs";
            if (nm.toUpperCase().includes("AC-OF-SM-ADSS")) sat = "Meter";
            else if (nm.toUpperCase().includes("PU-S7.0-400NM") || nm.toUpperCase().includes("PU-S9.0-140")) sat = "Batang";

            wsBA.getCell(`B${rowNum}`).value = nm;
            wsBA.getCell(`D${rowNum}`).value = sat;
            wsBA.getCell(`E${rowNum}`).value = parseInt(match[2]);
            wsBA.getCell(`F${rowNum}`).value = parseInt(match[2]);
            i++;
        }
        workbook.removeWorksheet('code gudang');

        // 3. Simpan ke Buffer (Bukan File Fisik)
        const buffer = await workbook.xlsx.writeBuffer();

        // 4. Konversi ke PDF via CloudConvert API
        const job = await cloudConvert.jobs.create({
            tasks: {
                'upload-file': { operation: 'import/base64', file: buffer.toString('base64'), filename: 'BA.xlsx' },
                'convert-file': { operation: 'convert', input: 'upload-file', output_format: 'pdf' },
                'export-file': { operation: 'export/url', input: 'convert-file' }
            }
        });

        // Tunggu proses konversi selesai (biasanya 2-5 detik)
        const finishedJob = await cloudConvert.jobs.wait(job.id);
        const exportTask = finishedJob.tasks.filter(t => t.operation === 'export/url' && t.status === 'finished')[0];
        const pdfUrl = exportTask.result.files[0].url;

        // 5. Kirim PDF Kembali ke Telegram via URL (Sangat Cepat)
        await axios.post(`${API_URL}/sendDocument`, {
            chat_id: chatId,
            document: pdfUrl,
            caption: `✅ BA Manual ${data.LOKASI || ''} berhasil dikonversi.`
        });

    } catch (e) {
        await axios.post(`${API_URL}/sendMessage`, { chat_id: chatId, text: "❌ Error: " + e.message });
    }
}