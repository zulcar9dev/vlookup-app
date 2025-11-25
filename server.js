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

// Buat folder uploads jika belum ada
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir);
}

const storage = multer.diskStorage({
    destination: (req, file, cb) => { cb(null, 'uploads/'); },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
    }
});
const upload = multer({ storage: storage });

// Helper: Normalisasi Key (PENTING: Memaksa jadi String agar 100% cocok)
function normalizeKey(value) {
    if (value === undefined || value === null) return null;
    return String(value).trim(); // Paksa jadi string dan hapus spasi
}

// Helper: Baca Header
function getHeaders(filePath) {
    try {
        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0]; 
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return [];
        const data = xlsx.utils.sheet_to_json(worksheet, { header: 1, range: 0 });
        if (data.length > 0 && data[0]) {
            return data[0].map(h => (typeof h === 'string' ? h.trim() : h));
        }
        return [];
    } catch (e) { return []; }
}

// 3. Rute
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

app.post('/inspect-files', upload.fields([
    { name: 'fileMain', maxCount: 1 },
    { name: 'fileLookupA', maxCount: 1 },
    { name: 'fileLookupB', maxCount: 1 }
]), (req, res) => {
    if (!req.files || !req.files.fileMain || !req.files.fileLookupA || !req.files.fileLookupB) {
        return res.status(400).json({ error: 'Harap unggah ketiga file.' });
    }
    const paths = [req.files.fileMain[0].path, req.files.fileLookupA[0].path, req.files.fileLookupB[0].path];
    const headers = paths.map(p => getHeaders(p));

    if (headers.some(h => h.length === 0)) {
        paths.forEach(p => { if (fs.existsSync(p)) fs.unlinkSync(p); });
        return res.status(400).json({ error: 'Gagal membaca file.' });
    }

    res.json({
        mainHeaders: headers[0],
        lookupAHeaders: headers[1],
        lookupBHeaders: headers[2],
        fileMainName: req.files.fileMain[0].filename, 
        fileLookupAName: req.files.fileLookupA[0].filename,
        fileLookupBName: req.files.fileLookupB[0].filename 
    });
});

app.post('/generate-report', async (req, res) => {
    const { fileMainName, fileLookupAName, fileLookupBName, orderedColumns } = req.body;
    if (!fileMainName) return res.status(400).json({ error: 'Data tidak lengkap.' });

    const paths = {
        main: path.join(__dirname, 'uploads', fileMainName),
        lookupA: path.join(__dirname, 'uploads', fileLookupAName),
        lookupB: path.join(__dirname, 'uploads', fileLookupBName)
    };

    if (!fs.existsSync(paths.main)) return res.status(400).json({ error: 'File kedaluwarsa.' });

    try {
        // 1. Baca File
        const wbMain = xlsx.readFile(paths.main);
        const wbLookupA = xlsx.readFile(paths.lookupA);
        const wbLookupB = xlsx.readFile(paths.lookupB);

        const dataMain = xlsx.utils.sheet_to_json(wbMain.Sheets[wbMain.SheetNames[0]]);
        const dataA = xlsx.utils.sheet_to_json(wbLookupA.Sheets[wbLookupA.SheetNames[0]]);
        const dataB = xlsx.utils.sheet_to_json(wbLookupB.Sheets[wbLookupB.SheetNames[0]]);

        console.log(`\n--- MULAI PROSES BARU ---`);
        console.log(`Total Baris Utama: ${dataMain.length}`);
        console.log(`Total Baris Ref A: ${dataA.length}`);

        // 2. Buat MAP dengan Sistem ANTRIAN (Array)
        // Kita gunakan Object Map agar lebih cepat
        const mapA = new Map();
        const mapB = new Map();

        // Fungsi pembantu untuk mengisi Map
        const populateMap = (dataset, mapTarget, sourceName) => {
            dataset.forEach(row => {
                const key = normalizeKey(row[KEY_COLUMN]);
                if (key) {
                    if (!mapTarget.has(key)) {
                        mapTarget.set(key, []); // Inisialisasi Array jika key baru
                    }
                    mapTarget.get(key).push(row); // Masukkan data ke antrian
                }
            });
            console.log(`Map ${sourceName} selesai. Contoh Key teratas:`, mapTarget.keys().next().value);
        };

        populateMap(dataA, mapA, "Ref A");
        populateMap(dataB, mapB, "Ref B");

        // 3. Proses Pencocokan (Looping File Utama)
        const combinedData = [];

        for (let i = 0; i < dataMain.length; i++) {
            const rowMain = dataMain[i];
            const key = normalizeKey(rowMain[KEY_COLUMN]);
            
            let matchA = undefined;
            let matchB = undefined;

            // --- LOGIKA PENGAMBILAN (SHIFT) ---
            // Cek Map A
            if (mapA.has(key)) {
                const queueA = mapA.get(key); // Ambil antrian
                if (queueA.length > 0) {
                    matchA = queueA.shift(); // AMBIL yang pertama, lalu HAPUS dari antrian
                    // Log untuk debug khusus key yang bermasalah
                    if (key === '10389134685') {
                        console.log(`[DEBUG] Baris Utama #${i+1}: Mengambil data Ref A. Sisa antrian: ${queueA.length}`);
                    }
                }
            }

            // Cek Map B
            if (mapB.has(key)) {
                const queueB = mapB.get(key);
                if (queueB.length > 0) {
                    matchB = queueB.shift();
                }
            }

            // --- Penyusunan Baris Baru ---
            let newRow = {};

            // Logika Status/Formula
            let calculatedStatus = "TAGIH";
            let saldo = (matchA?.SALDO_AKHIR_AFILIASI_NEW) ?? (matchB?.SALDO_AKHIR_AFILIASI_NEW) ?? 0;
            let kewajiban = (matchA?.Total_Kewajiban_New) ?? (matchB?.Total_Kewajiban_New) ?? 0;

            // Bersihkan format angka (hapus koma)
            const parseNum = (val) => parseFloat(String(val).replace(/,/g, '')) || 0;
            if (parseNum(saldo) > parseNum(kewajiban)) {
                calculatedStatus = "AMAN";
            }

            // Isi Kolom Sesuai Urutan
            orderedColumns.forEach(colInfo => {
                const col = colInfo.column;
                if (col === 'STATUS' && colInfo.useFormula) {
                    newRow[col] = calculatedStatus;
                } else if (colInfo.source === 'main') {
                    newRow[col] = rowMain[col];
                } else if (colInfo.source === 'lookupA') {
                    newRow[col] = matchA ? matchA[col] : "TIDAK_ADA_DATA";
                } else if (colInfo.source === 'lookupB') {
                    newRow[col] = matchB ? matchB[col] : "TIDAK_ADA_DATA";
                }
            });

            combinedData.push(newRow);
        }

        // 4. Buat File Output
        const newBook = xlsx.utils.book_new();
        const newSheet = xlsx.utils.json_to_sheet(combinedData);
        xlsx.utils.book_append_sheet(newBook, newSheet, 'Hasil_VLOOKUP');
        const buffer = xlsx.write(newBook, { bookType: 'xlsx', type: 'buffer' });

        res.setHeader('Content-Disposition', 'attachment; filename="Hasil_VLOOKUP_Antrian.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);

    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    } finally {
        // Bersihkan file
        Object.values(paths).forEach(p => { if (fs.existsSync(p)) fs.unlinkSync(p); });
    }
});

// Global Error Handler
app.use((err, req, res, next) => {
    if (err instanceof multer.MulterError) return res.status(500).json({ error: `Upload Error: ${err.message}` });
    res.status(500).json({ error: err.message });
});

app.listen(port, () => {
    console.log(`Server siap di http://localhost:${port}`);
    exec(`explorer "http://localhost:${port}"`);
});