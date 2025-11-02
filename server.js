// 1. Impor Pustaka (Library)
const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

// 2. Konfigurasi Aplikasi
const app = express();
const port = 3000;
app.use(express.json());

// === PENGATURAN VLOOKUP ===
const KEY_COLUMN = 'BNI_CIF_KEY'; 
// ==========================

// Konfigurasi Multer
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/');
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
    }
});
const upload = multer({ storage: storage });

// Helper function untuk membaca header
function getHeaders(filePath) {
    try {
        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0]; 
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return [];
        
        const data = xlsx.utils.sheet_to_json(worksheet, { header: 1, range: 0 });
        if (data.length > 0) {
            return data[0];
        }
        return [];
    } catch (e) {
        console.error("Gagal membaca header:", e);
        return [];
    }
}

// 3. Rute (Routes)

// Rute utama (GET /)
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

/**
 * ENDPOINT 1: INSPEKSI FILE
 * (Tidak ada perubahan dari sebelumnya)
 */
app.post('/inspect-files', upload.fields([
    { name: 'fileMain', maxCount: 1 },
    { name: 'fileLookup', maxCount: 1 }
]), (req, res) => {
    
    if (!req.files || !req.files.fileMain || !req.files.fileLookup) {
        return res.status(400).json({ error: 'Harap unggah kedua file.' });
    }

    const fileMainPath = req.files.fileMain[0].path;
    const fileLookupPath = req.files.fileLookup[0].path;

    const mainHeaders = getHeaders(fileMainPath);
    const lookupHeaders = getHeaders(fileLookupPath);

    if (mainHeaders.length === 0 || lookupHeaders.length === 0) {
        fs.unlinkSync(fileMainPath);
        fs.unlinkSync(fileLookupPath);
        return res.status(400).json({ error: 'Gagal membaca header dari file. Pastikan file tidak kosong.' });
    }

    res.json({
        mainHeaders: mainHeaders,
        lookupHeaders: lookupHeaders,
        fileMainName: req.files.fileMain[0].filename, 
        fileLookupName: req.files.fileLookup[0].filename 
    });
});

/**
 * ENDPOINT 2: PROSES VLOOKUP & GENERATE
 * !!! PERUBAHAN UTAMA DI SINI !!!
 * Menerima 'orderedColumns' sebagai array objek.
 */
app.post('/generate-report', async (req, res) => {
    // Ambil data dari body, 'orderedColumns' adalah yang baru
    const { fileMainName, fileLookupName, orderedColumns } = req.body;

    if (!fileMainName || !fileLookupName || !orderedColumns || orderedColumns.length === 0) {
        return res.status(400).json({ error: 'Data tidak lengkap untuk generate laporan.' });
    }

    const fileMainPath = path.join(__dirname, 'uploads', fileMainName);
    const fileLookupPath = path.join(__dirname, 'uploads', fileLookupName);
    
    if (!fs.existsSync(fileMainPath) || !fs.existsSync(fileLookupPath)) {
        return res.status(400).json({ error: 'Sesi file tidak ditemukan. Harap unggah ulang file.' });
    }

    try {
        // --- Membaca File Excel (Seluruhnya) ---
        const wbMain = xlsx.readFile(fileMainPath);
        const wbLookup = xlsx.readFile(fileLookupPath);

        const sheet1Data = xlsx.utils.sheet_to_json(wbMain.Sheets[wbMain.SheetNames[0]]);
        const sheet2Data = xlsx.utils.sheet_to_json(wbLookup.Sheets[wbLookup.SheetNames[0]]);

        // --- Optimasi VLOOKUP (membuat Map) ---
        const lookupMap = new Map();
        for (const row of sheet2Data) {
            const key = row[KEY_COLUMN];
            if (key) {
                lookupMap.set(key, row);
            }
        }

        // --- Proses Penggabungan Data (VLOOKUP) ---
        const combinedData = [];

        for (const rowSheet1 of sheet1Data) {
            const lookupValue = rowSheet1[KEY_COLUMN];
            const matchSheet2 = lookupMap.get(lookupValue);

            let newRow = {};

            // === LOGIKA BARU ===
            // Loop melalui array 'orderedColumns' yang dikirim dari frontend
            for (const colInfo of orderedColumns) {
                const colName = colInfo.column;

                if (colInfo.source === 'main') {
                    // Jika sumbernya 'main', ambil dari rowSheet1
                    newRow[colName] = rowSheet1[colName];
                
                } else if (colInfo.source === 'lookup') {
                    // Jika sumbernya 'lookup', ambil dari data yang cocok (matchSheet2)
                    if (matchSheet2) {
                        newRow[colName] = matchSheet2[colName]; // Data ditemukan
                    } else {
                        newRow[colName] = "data_tidak_ditemukan"; // Data tidak ditemukan
                    }
                }
            }
            // ===================
            
            combinedData.push(newRow);
        }

        // --- Membuat File Excel Baru (Output) ---
        const newWorkbook = xlsx.utils.book_new();
        // Penting: json_to_sheet akan otomatis mengikuti urutan key di objek newRow,
        // yang mana sudah kita tentukan urutannya saat membangun newRow.
        const newWorksheet = xlsx.utils.json_to_sheet(combinedData);
        xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Hasil_VLOOKUP');
        const outputBuffer = xlsx.write(newWorkbook, { bookType: 'xlsx', type: 'buffer' });

        // --- Mengirim File Hasil ke Pengguna ---
        res.setHeader('Content-Disposition', 'attachment; filename="Hasil_VLOOKUP.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(outputBuffer);

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Terjadi error internal saat memproses file: ' + error.message });
    } finally {
        // --- PENTING: Hapus file sementara ---
        fs.unlinkSync(fileMainPath);
        fs.unlinkSync(fileLookupPath);
    }
});

// 4. Jalankan Server
app.listen(port, () => {
    console.log(`Server berjalan di http://localhost:${port}`);
});