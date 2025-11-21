// 1. Impor Pustaka
const { exec } = require('child_process');
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
// !!! PERUBAHAN UTAMA DI SINI !!!
function getHeaders(filePath) {
    try {
        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0]; 
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return [];
        
        const data = xlsx.utils.sheet_to_json(worksheet, { header: 1, range: 0 });
        
        if (data.length > 0 && data[0]) {
            // BARU: Membersihkan (trim) spasi dari setiap nama header
            return data[0].map(header => {
                if (typeof header === 'string') {
                    return header.trim(); // Menghapus spasi di awal/akhir
                }
                return header; // Kembalikan apa adanya (jika angka atau null)
            });
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
 * (Tidak ada perubahan, tapi sekarang menggunakan getHeaders yang baru)
 */
app.post('/inspect-files', upload.fields([
    { name: 'fileMain', maxCount: 1 },
    { name: 'fileLookupA', maxCount: 1 },
    { name: 'fileLookupB', maxCount: 1 }
]), (req, res) => {
    
    if (!req.files || !req.files.fileMain || !req.files.fileLookupA || !req.files.fileLookupB) {
        return res.status(400).json({ error: 'Harap unggah ketiga file.' });
    }

    const fileMainPath = req.files.fileMain[0].path;
    const fileLookupAPath = req.files.fileLookupA[0].path;
    const fileLookupBPath = req.files.fileLookupB[0].path;

    const mainHeaders = getHeaders(fileMainPath);
    const lookupAHeaders = getHeaders(fileLookupAPath);
    const lookupBHeaders = getHeaders(fileLookupBPath);

    if (mainHeaders.length === 0 || lookupAHeaders.length === 0 || lookupBHeaders.length === 0) {
        // Hapus file jika gagal baca header
        try {
            fs.unlinkSync(fileMainPath);
            fs.unlinkSync(fileLookupAPath);
            fs.unlinkSync(fileLookupBPath);
        } catch (e) { console.error("Gagal hapus file:", e); }
        return res.status(400).json({ error: 'Gagal membaca header dari satu atau lebih file.' });
    }

    res.json({
        mainHeaders: mainHeaders,
        lookupAHeaders: lookupAHeaders,
        lookupBHeaders: lookupBHeaders,
        fileMainName: req.files.fileMain[0].filename, 
        fileLookupAName: req.files.fileLookupA[0].filename,
        fileLookupBName: req.files.fileLookupB[0].filename 
    });
});

/**
 * ENDPOINT 2: PROSES VLOOKUP & GENERATE
 * (Tidak ada perubahan di sini)
 */
app.post('/generate-report', async (req, res) => {
    const { fileMainName, fileLookupAName, fileLookupBName, orderedColumns } = req.body;

    if (!fileMainName || !fileLookupAName || !fileLookupBName || !orderedColumns || orderedColumns.length === 0) {
        return res.status(400).json({ error: 'Data tidak lengkap untuk generate laporan.' });
    }

    const fileMainPath = path.join(__dirname, 'uploads', fileMainName);
    const fileLookupAPath = path.join(__dirname, 'uploads', fileLookupAName);
    const fileLookupBPath = path.join(__dirname, 'uploads', fileLookupBName);
    
    if (!fs.existsSync(fileMainPath) || !fs.existsSync(fileLookupAPath) || !fs.existsSync(fileLookupBPath)) {
        return res.status(400).json({ error: 'Sesi file tidak ditemukan. Harap unggah ulang file.' });
    }

    try {
        // --- Membaca 3 File Excel ---
        const wbMain = xlsx.readFile(fileMainPath);
        const wbLookupA = xlsx.readFile(fileLookupAPath);
        const wbLookupB = xlsx.readFile(fileLookupBPath);

        const sheet1Data = xlsx.utils.sheet_to_json(wbMain.Sheets[wbMain.SheetNames[0]]);
        const sheet2DataA = xlsx.utils.sheet_to_json(wbLookupA.Sheets[wbLookupA.SheetNames[0]]);
        const sheet2DataB = xlsx.utils.sheet_to_json(wbLookupB.Sheets[wbLookupB.SheetNames[0]]);

        // --- Optimasi VLOOKUP (membuat 2 Map) ---
        const lookupMapA = new Map();
        for (const row of sheet2DataA) {
            // BERSIHKAN KUNCI SAAT MEMBUAT MAP
            const key = (row[KEY_COLUMN] && typeof row[KEY_COLUMN] === 'string') ? row[KEY_COLUMN].trim() : row[KEY_COLUMN];
            if (key) lookupMapA.set(key, row);
        }
        const lookupMapB = new Map();
        for (const row of sheet2DataB) {
            // BERSIHKAN KUNCI SAAT MEMBUAT MAP
            const key = (row[KEY_COLUMN] && typeof row[KEY_COLUMN] === 'string') ? row[KEY_COLUMN].trim() : row[KEY_COLUMN];
            if (key) lookupMapB.set(key, row);
        }

        // --- Proses Penggabungan Data (VLOOKUP) ---
        const combinedData = [];

        for (const rowSheet1 of sheet1Data) {
            // BERSIHKAN KUNCI SAAT MENCARI
            const lookupValueRaw = rowSheet1[KEY_COLUMN];
            const lookupValue = (lookupValueRaw && typeof lookupValueRaw === 'string') ? lookupValueRaw.trim() : lookupValueRaw;
            
            const matchSheetA = lookupMapA.get(lookupValue);
            const matchSheetB = lookupMapB.get(lookupValue);

            let newRow = {};

            // === Hitung Formula DULU ===
            let calculatedStatus = "TAGIH";
            let saldo = (matchSheetA && matchSheetA.SALDO_AKHIR_AFILIASI_NEW !== undefined) 
                ? matchSheetA.SALDO_AKHIR_AFILIASI_NEW 
                : (matchSheetB && matchSheetB.SALDO_AKHIR_AFILIASI_NEW);
            let kewajiban = (matchSheetA && matchSheetA.Total_Kewajiban_New !== undefined) 
                ? matchSheetA.Total_Kewajiban_New
                : (matchSheetB && matchSheetB.Total_Kewajiban_New);

            const saldoNum = parseFloat(String(saldo).replace(/,/g, '')) || 0;
            const kewajibanNum = parseFloat(String(kewajiban).replace(/,/g, '')) || 0;

            if (saldoNum > kewajibanNum) {
                calculatedStatus = "AMAN";
            }
            // =============================

            // === Loop melalui kolom pilihan pengguna ===
            for (const colInfo of orderedColumns) {
                const colName = colInfo.column; // colName ini sudah di-trim dari frontend

                if (colInfo.column === 'STATUS' && colInfo.useFormula === true) {
                    newRow[colName] = calculatedStatus;
                
                } else if (colInfo.source === 'main') {
                    newRow[colName] = rowSheet1[colName];
                
                } else if (colInfo.source === 'lookupA') {
                    if (matchSheetA) {
                        newRow[colName] = matchSheetA[colName];
                    } else {
                        newRow[colName] = "data_tidak_ditemukan_A";
                    }
                
                } else if (colInfo.source === 'lookupB') {
                    if (matchSheetB) {
                        newRow[colName] = matchSheetB[colName];
                    } else {
                        newRow[colName] = "data_tidak_ditemukan_B";
                    }
                }
            }
            
            combinedData.push(newRow);
        }

        // --- Membuat File Excel Baru (Output) ---
        const newWorkbook = xlsx.utils.book_new();
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
        try {
            fs.unlinkSync(fileMainPath);
            fs.unlinkSync(fileLookupAPath);
            fs.unlinkSync(fileLookupBPath);
        } catch(e) { console.error("Gagal hapus file:", e); }
    }
});

// 4. Jalankan Server
app.listen(3000, () => {
  console.log('Server berjalan di http://localhost:3000');
  
  // Perintah ajaib untuk membuka browser otomatis di Windows
  exec('explorer "http://localhost:3000"'); 
});