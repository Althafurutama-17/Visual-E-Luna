# Visual E-LUNA: Dashboard Operator Gudang Logistik BPBD Jawa Barat

Sistem antarmuka (*dashboard*) pemantauan operasional logistik kebencanaan untuk **Badan Penanggulangan Bencana Daerah (BPBD) Provinsi Jawa Barat**, terintegrasi langsung dengan basis data spreadsheet Excel (*two-way client-side parsing & auto-sync*).

---

## 🚀 Fitur Utama Dashboard

1. **Integrasi Langsung Berkas Excel (.xlsx)**:
   * Menggunakan pustaka *SheetJS* (`xlsx.full.min.js`) yang beroperasi 100% secara offline di sisi peramban (*client-side*).
   * Fitur **Muat Ulang Excel**, **Unggah Excel Kustom**, dan **Unduh Basis Data**.
   * Indikator status koneksi basis data aktif pada bilah navigasi atas.

2. **Ringkasan Arus Logistik & Kartu Metrik KPI**:
   * Total Penerimaan Masuk (Unit & Jumlah Dokumen BAST).
   * Total Distribusi Keluar (Unit & Jumlah Surat Jalan).
   * Total Saldo Stok Akhir Gudang.
   * Jumlah Batch Barang Kadaluwarsa & Kategori Peringatan (*Warning*).

3. **Neraca Stok 4 Kategori Barang Logistik**:
   * Menampilkan neraca stok per 4 kategori resmi: **Sandang, Pangan, Papan, dan Logistik Lainnya**.
   * Stok Awal dijaga konsisten pada baseline saldo master.
   * Penerimaan (+) dan Distribusi (-) merespons filter aktif.
   * Angka Stok Akhir tampil dengan aksen warna biru khas E-LUNA.

4. **Visualisasi Statistik Interaktif (Chart.js)**:
   * **Statistik Penerimaan Berdasarkan Sumber Anggaran**: Visualisasi volume barang masuk menurut sumber pendanaan (**APBD, APBN, CSR**) yang dinamis mengikuti filter waktu.
   * **Statistik Distribusi ke 27 Kabupaten/Kota**: Visualisasi peringkat alokasi penyaluran logistik ke seluruh daerah rawan bencana di Jawa Barat.

5. **Modul Pengawasan Kadaluwarsa Dinamis**:
   * Parameter peringatan dini dinamis ($X$ Hari sebelum tanggal kadaluwarsa asli).
   * Filter tab status kelayakan (*Semua, Kadaluwarsa, Warning, Aman*).

6. **Tabel Riwayat Mutasi Operator & Filter**:
   * Mode Tampilan: **10 Terakhir (Default)** vs **Semua Data Transaksi Kumulatif**.
   * Filter waktu murni berdasarkan **Tahun** dan **Bulan**.
   * Tab filter jenis mutasi: **Semua Mutasi**, **Masuk Saja (BAST)**, dan **Keluar Saja (Surat Jalan)**.

---

## 💻 Panduan Menjalankan Aplikasi

### Opsi 1: Menjalankan Langsung (Tanpa Instalasi Server)
Cukup buka berkas **`index.html`** langsung menggunakan peramban modern (Google Chrome, Microsoft Edge, Mozilla Firefox, dll.).

### Opsi 2: Menjalankan via Smart Server Python (Direkomendasikan)
Server pintar ini dilengkapi dengan penonaktifan *cache* browser serta *auto-watcher* file Excel:
```bash
python server.py
```
Akses dashboard pada peramban melalui alamat:
```text
http://localhost:3456
```

---

## 📁 Struktur Direktori

```text
├── index.html                           # Halaman utama dashboard operator E-LUNA
├── styles.css                           # Lembar gaya CSS terpadu desain resmi E-LUNA
├── app.js                               # Pengendali logika, parser SheetJS, Chart.js, & filter
├── xlsx.full.min.js                     # Mesin parser Excel biner offline (SheetJS)
├── database_logistik_bpbd_jabar.xlsx    # Basis data spreadsheet utama (250 Transaksi mutasi)
├── server.py                            # Server lokal pintar (Anti-cache & Auto-sync)
├── sync_excel_to_web.py                 # Skrip CLI sinkronisasi data Excel ke JSON/JS
├── data.js                              # Cadangan basis data JavaScript offline (fallback)
├── database_eluna.json                  # Cadangan basis data JSON
├── .gitignore                           # Konfigurasi pengabaian file Git
└── README.md                            # Dokumentasi teknis repositori
```

---

## 🏛️ Instansi & Kontributor
* **Instansi Mitra**: BPBD Provinsi Jawa Barat (Bidang Kedaruratan & Logistik)
* **Pengembang**: Nadhif Althafu Rutama (S1 Teknik Logistik - *Digital Supply Chain*, Universitas Telkom)
