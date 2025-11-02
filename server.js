// 1. Impor Pustaka (Library)
const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');

// 2. Konfigurasi Aplikasi
const app = express();
const port = 3000;

// Konfigurasi Multer untuk menyimpan file di memori (bukan di disk)
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// === PENGATURAN VLOOKUP (SESUAI PERMINTAAN ANDA) ===
const SHEET1_NAME = 'Sheet1'; // Nama sheet utama
const SHEET2_NAME = 'Sheet2'; // Nama sheet referensi
const KEY_COLUMN = 'BNI_CIF_KEY'; // Kunci pencocokan

// Daftar 34 kolom yang akan diambil dari Sheet2
const COLUMNS_TO_FETCH = [
    'BNI_CIF_KEY', 'ID_NUMBER', 'No_Rekening_Afiliasi', 'CUSTOMER_NAME', 
    'Product_Name5', 'GOLONGAN', 'Cycle', 'BAKI_DEBET_NEW', 'Saldo_Blokir', 
    'Total_Tunggakkan_New', 'ANGSURAN_BUNGA', 'ANGSURAN_POKOK', 'Total_Angsuran', 
    'Total_Kewajiban_New', 'SALDO_AKHIR_AFILIASI_NEW', 'program', 
    'CUSTOMER_ADDRESS_1', 'CUSTOMER_ADDRESS_2', 'MobilePhone_elo', 'Nama_Perusahaan', 
    'ALAMAT_KANTOR', 'CUSTOMER_ADDRESS_1', 'CUSTOMER_ADDRESS_2', 'NAMA_AKK', 
    'TELP_RUMAH', 'TELP_KANTOR', 'TELP_HP1', 'NO_HP_ELO', 'NO_HP_WONDR', 
    'NO_HP_MBANK', 'HomePhone_elo', 'OfficePhone_elo', 'SpousePhone', 
    'EmergencyContact'
];
// =====================================================

const COLUMNS_TO_ADD = COLUMNS_TO_FETCH.filter(col => col !== KEY_COLUMN);

// 3. Rute (Routes)

// Rute utama (GET /) untuk menyajikan file index.html
app.get('/', (req, res) => {
    // Mengirim file index.html yang ada di folder yang sama
    res.sendFile(__dirname + '/index.html');
});

// Rute (POST /upload) untuk menangani file yang diunggah
app.post('/upload', upload.single('excelFile'), (req, res) => {
    
    // Validasi: Pastikan file ada
    if (!req.file) {
        return res.status(400).send('Tidak ada file yang diunggah.');
    }

    try {
        // --- Membaca File Excel ---
        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });

        // Validasi: Pastikan sheet yang dibutuhkan ada
        if (!workbook.SheetNames.includes(SHEET1_NAME) || !workbook.SheetNames.includes(SHEET2_NAME)) {
            return res.status(400).send(`Error: Pastikan file Excel Anda memiliki sheet bernama '${SHEET1_NAME}' dan '${SHEET2_NAME}'.`);
        }
        
        // Ubah sheet menjadi data JSON
        const sheet1Data = xlsx.utils.sheet_to_json(workbook.Sheets[SHEET1_NAME]);
        const sheet2Data = xlsx.utils.sheet_to_json(workbook.Sheets[SHEET2_NAME]);
        
        // --- Optimasi VLOOKUP ---
        // Kita ubah Sheet2 menjadi "Map" (seperti kamus) agar pencarian lebih cepat.
        // Ini jauh lebih efisien daripada melakukan loop di dalam loop.
        const lookupMap = new Map();
        for (const row of sheet2Data) {
            const key = row[KEY_COLUMN];
            if (key) {
                // Menyimpan seluruh baris data Sheet2 dengan kuncinya
                lookupMap.set(key, row);
            }
        }

        // --- Proses Penggabungan Data (VLOOKUP) ---
        const combinedData = [];

        for (const rowSheet1 of sheet1Data) {
            const lookupValue = rowSheet1[KEY_COLUMN];
            const matchSheet2 = lookupMap.get(lookupValue);

            // Kita mulai dengan menyalin semua data dari baris Sheet1
            const newRow = { ...rowSheet1 };

           // Ambil data dari Sheet2 jika ada kecocokan
        // Kita loop 33 kolom (daftar COLUMNS_TO_ADD, TANPA kunci)
        for (const colName of COLUMNS_TO_ADD) { 
            if (matchSheet2) {
                // Jika cocok, salin data dari Sheet2
                newRow[colName] = matchSheet2[colName];
            } else {
                // MODIFIKASI: Jika tidak cocok, isi dengan "data_tidak_ditemukan"
                newRow[colName] = "data_tidak_ditemukan"; 
            }
        }
            
            combinedData.push(newRow);
        }

        // --- Membuat File Excel Baru (Output) ---
        const newWorkbook = xlsx.utils.book_new();
        const newWorksheet = xlsx.utils.json_to_sheet(combinedData);
        
        // Menambahkan sheet hasil ke workbook baru
        xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Hasil_VLOOKUP');

        // Mengubah workbook menjadi buffer (format yang bisa dikirim lewat internet)
        const outputBuffer = xlsx.write(newWorkbook, { bookType: 'xlsx', type: 'buffer' });

        // --- Mengirim File Hasil ke Pengguna ---
        
        // Atur header respons agar browser tahu ini adalah file download
        res.setHeader(
            'Content-Disposition', 
            'attachment; filename="Hasil_VLOOKUP.xlsx"'
        );
        res.setHeader(
            'Content-Type', 
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );

        // Kirim filenya!
        res.send(outputBuffer);

    } catch (error) {
        console.error(error);
        res.status(500).send('Terjadi error saat memproses file: ' + error.message);
    }
});

// 4. Jalankan Server
app.listen(port, () => {
    console.log(`Server berjalan di http://localhost:${port}`);
});