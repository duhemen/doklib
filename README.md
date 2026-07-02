<div align="center">
  <img src="https://img.shields.io/badge/version-2.0.0-blue.svg" alt="Version">
  <img src="https://img.shields.io/badge/Python-3.8+-green.svg" alt="Python">
  <img src="https://img.shields.io/badge/Streamlit-1.28+-red.svg" alt="Streamlit">
  <img src="https://img.shields.io/badge/TNDE-Compliant-brightgreen.svg" alt="TNDE">
  <img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License">
</div>

<br>

<div align="center">
  <h1>🚀 APSO Ultimate</h1>
  <h3>Aplikasi Surat Otomatis</h3>
  <p><strong>Sesuai Standar TNDE (Tata Naskah Dinas Elektronik)</strong></p>
  
  <br>
  
  <p>
    <a href="#-fitur-unggulan">Fitur</a> •
    <a href="#-instalasi">Instalasi</a> •
    <a href="#-panduan-penggunaan">Panduan</a> •
    <a href="#-struktur-proyek">Struktur</a> •
    <a href="#-kontribusi">Kontribusi</a>
  </p>
</div>

---

### 🎯 Tentang APSO Ultimate

**APSO Ultimate** (Aplikasi Surat Otomatis) adalah sistem berbasis web yang dirancang khusus untuk mengotomatiskan pembuatan surat dinas sesuai dengan **Standar TNDE (Tata Naskah Dinas Elektronik)**.

Aplikasi ini mengeliminasi proses manual yang rentan error dengan menyediakan antarmuka yang intuitif untuk mengisi data surat, mengganti placeholder secara otomatis, dan menghasilkan dokumen Word yang siap pakai dengan format yang konsisten.

### 🎯 Tujuan

- **Efisiensi**: Mengurangi waktu pembuatan surat dari 30 menit menjadi 5 menit
- **Konsistensi**: Menjamin keseragaman format sesuai standar TNDE
- **Akurasi**: Meminimalkan kesalahan penulisan dan format dokumen
- **Produktivitas**: Memungkinkan staf fokus pada konten而非 format

---

### ✨ Fitur Unggulan

### 📝 Standar TNDE Penuh
- **Margin Otomatis**: Kiri 4cm, Kanan 3cm, Atas 4cm, Bawah 3cm
- **Font Standar**: Times New Roman 12pt dengan spasi 1.5
- **Format Nomor Surat**: 001/BP2JK-KT/01/2026
- **Alignment**: Justify untuk semua paragraf

### 🧠 Smart Features
| Fitur | Deskripsi |
|-------|-----------|
| **Smart Placeholder** | Support multiple formats: `[NAMA]`, `{pembuka}`, `$NIP$` |
| **Split-Run Handling** | Mengatasi placeholder yang terpecah di XML Word |
| **Dynamic Table Injection** | Mapping berdasarkan nama kolom (bukan urutan) |
| **In-Body Table Detection** | Deteksi otomatis tabel di dalam template Word |

### 🖥️ User Interface
- **Multi-Tab Layout**: Teks, Tabel, Gambar, Penandatangan, Pengaturan
- **Live Preview**: Lihat hasil sebelum mencetak
- **Download Ready**: Satu klik untuk download dokumen .docx
- **Responsive Design**: Optimal di desktop dan tablet

### 🔒 Keamanan & Logging
- **Audit Trail**: Setiap aktivitas tercatat di log
- **Validasi Input**: Pengecekan nomor surat dan data wajib
- **Session Management**: Data tetap aman selama sesi berlangsung

---

### 🛠️ Teknologi yang Digunakan

<div align="center">
  
| Teknologi | Versi | Kegunaan |
|-----------|-------|----------|
| ![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white) | 3.8+ | Bahasa pemrograman utama |
| ![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white) | 1.28+ | Framework web untuk UI |
| ![python-docx](https://img.shields.io/badge/python--docx-3776AB?style=for-the-badge&logo=python&logoColor=white) | 0.8.11+ | Manipulasi dokumen Word |
| ![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white) | 1.5+ | Manipulasi data tabel |

</div>

---

## 💻 Persyaratan Sistem

### Minimum Requirements
- **OS**: Windows 10/11, macOS 10.15+, atau Linux (Ubuntu 20.04+)
- **RAM**: 4 GB (8 GB recommended)
- **Storage**: 500 MB untuk aplikasi + space untuk template dan output
- **Python**: 3.8 atau lebih tinggi

### Recommended
- **OS**: Windows 11 atau macOS 12+
- **RAM**: 8 GB
- **Storage**: 2 GB
- **Python**: 3.10+

---

## 📥 Instalasi

### 1. Clone Repository
```bash
git clone https://github.com/duhemen/doklib.git
cd apso-ultimate
```

### 2. Buat Virtual Environment (Opsional tapi Recommended)
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Setup Folder Struktur
```bash
# Buat folder yang diperlukan
mkdir C:\doklib\tnd  # Windows
# atau
mkdir ~/doklib/tnd   # macOS/Linux

# Masukkan template surat (.docx) ke folder tnd/
```

### 5. Jalankan Aplikasi
```bash
streamlit run app.py
```

Aplikasi akan terbuka di browser Anda pada `http://localhost:8501`

---

## 📖 Panduan Penggunaan

### Langkah 1: Persiapkan Template Word

1. **Buat template Word** dengan placeholder:
   - `[NOMOR_SURAT]` untuk nomor surat
   - `{pembuka}` untuk alinea pembuka
   - `$NAMA_ATASAN$` untuk nama atasan
   - `(Alinea isi)` untuk alinea isi

2. **Simpan template** di folder `C:\doklib\tnd\` dengan ekstensi `.docx`

3. **Template harus memiliki**:
   - Kop surat (logo, instansi, alamat)
   - Tempat untuk nomor surat dan perihal
   - Area untuk alinea (pembuka, isi, penutup)
   - Tempat tanda tangan (opsional)

### Langkah 2: Buka Aplikasi

1. Jalankan `streamlit run app.py`
2. Browser akan terbuka otomatis
3. Pilih template surat dari dropdown

### Langkah 3: Isi Data Surat

#### Tab 1: Teks & Isi 📝
1. **Identitas Surat**:
   - Pilih Tahun, Bulan, dan Nomor Urut (otomatis generate nomor surat)
   - Isi Perihal/Hal Surat
   - Pilih Jenis Surat dan Sifat Surat

2. **Alinea Surat**:
   - Gunakan Template Alinea Cepat atau isi manual
   - Isi Alinea Pembuka, Isi, dan Penutup
   - Tentukan Tempat dan Tanggal Surat

#### Tab 2: Tabel Data 📊
1. **Tentukan Kolom**: Masukkan nama kolom (pisahkan dengan koma)
   - Contoh: `NO, Jam, Hari Ke-1, Hari Ke-2`
   
2. **Isi Data**: Edit tabel langsung di antarmuka
   - Tambah/ Hapus baris sesuai kebutuhan
   - Data akan otomatis mengisi tabel di Word (jika nama kolom cocok)

#### Tab 3: Gambar 🖼️
1. Upload gambar pendukung (PNG/JPG)
2. Preview gambar akan tampil
3. Gambar akan otomatis ditempatkan di lampiran

#### Tab 4: Penandatangan 👥
1. Pilih jumlah penandatangan (1-3)
2. Isi:
   - Jabatan
   - Nama Lengkap
   - NIP

#### Tab 5: Pengaturan ⚙️
- Atur folder template dan output
- Aktifkan auto-open setelah selesai
- Lihat log aktivitas

### Langkah 4: Preview & Cetak

1. **Preview**: Klik "👁️ Preview Dokumen" untuk melihat hasil
   - Download file preview untuk dicek di Word
   - Lihat ringkasan data yang akan diproses

2. **Cetak**: Klik "🚀 Cetak Dokumen" untuk menghasilkan file final
   - Validasi otomatis akan memeriksa data
   - File akan disimpan di folder output

3. **Download**: File .docx siap diunduh atau dibuka langsung di Word

### 🎯 Tips & Trik

| Tip | Deskripsi |
|-----|-----------|
| **Template Placeholder** | Gunakan `[NAMA]` untuk teks, `{pembuka}` untuk alinea, `$NIP$` untuk format khusus |
| **Tabel Dinamis** | Pastikan nama kolom di Word sama dengan di aplikasi (case-insensitive) |
| **Gambar** | Ukuran gambar akan otomatis diresize ke 4 inci untuk final, 2.5 inci untuk preview |
| **Format Dokumen** | Format Bold/Italic/Color akan dipertahankan saat penggantian placeholder |
| **Split-Run** | Jika placeholder tidak terganti, ketik ulang placeholder di Word (jangan copy-paste) |

---

## 📁 Struktur Proyek

```
apso-ultimate/
│
├── app.py                      # File utama aplikasi
├── requirements.txt            # Dependencies
├── README.md                   # Dokumentasi
├── .gitignore                 # Git ignore rules
├── setup.py                   # Setup untuk packaging
│
├── C:\doklib\                 # Folder data (default)
│   ├── tnd\                   # Template surat (.docx)
│   ├── logs\                  # Log aktivitas
│   └── [Hasil_Surat].docx     # Output dokumen
│
├── .streamlit/                # Konfigurasi Streamlit
│   └── config.toml
│
└── tests/                     # Unit testing (coming soon)
    └── test_app.py
```

---

## 🐛 Troubleshooting

### 1. Folder 'C:\doklib\tnd' Tidak Ditemukan
**Solusi**:
- Buat folder secara manual: `mkdir C:\doklib\tnd`
- Atau buat melalui aplikasi dengan tombol "Buat Folder Sekarang"

### 2. Placeholder Tidak Terganti di Word
**Penyebab**: Word memecah teks menjadi beberapa run (split-run)
**Solusi**:
- Hapus placeholder dan ketik ulang tanpa mengubah format
- Atau gunakan fitur "Optimasi Split-Run" yang sudah terintegrasi

### 3. Error Saat Upload Gambar
**Solusi**:
- Pastikan format file: PNG, JPG, atau JPEG
- Periksa ukuran file (max 10 MB)
- Jika error, refresh halaman dan coba lagi

### 4. Aplikasi Tidak Bisa Menyimpan File
**Solusi**:
- Periksa izin folder `C:\doklib\`
- Jalankan Streamlit sebagai administrator (Windows)
- Atau ubah folder output di tab Pengaturan

### 5. Tombol "Buka di Word" Tidak Berfungsi
**Penyebab**: Fitur hanya untuk Windows lokal
**Solusi**:
- Gunakan download button untuk menyimpan file
- Buka file secara manual dari folder output

---

## 🤝 Kontribusi

Kami menyambut kontribusi dari pengembang lain! Berikut panduan kontribusi:

### 1. Fork Repository
```bash
git fork https://github.com/your-username/doklib.git
```

### 2. Buat Branch Fitur
```bash
git checkout -b feature/your-feature-name
```

### 3. Lakukan Perubahan
- Ikuti style guide Python (PEP 8)
- Tambahkan komentar yang jelas
- Update dokumentasi jika perlu

### 4. Commit dan Push
```bash
git commit -m "feat: add your feature description"
git push origin feature/your-feature-name
```

### 5. Buat Pull Request
- Jelaskan perubahan yang dilakukan
- Lampirkan screenshot jika ada perubahan UI

### Area yang Perlu Dikembangkan
- [ ] Unit testing
- [ ] Multi-user support
- [ ] Cloud deployment (Docker)
- [ ] Export ke PDF
- [ ] Template management UI
- [ ] Database integration

---

## 📄 Lisensi

Proyek ini dilisensikan di bawah **MIT License** - lihat file [LICENSE](LICENSE) untuk detail.

```
MIT License

Copyright (c) 2026 BP2JK Wilayah Kalimantan Tengah

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:
...
```

---
