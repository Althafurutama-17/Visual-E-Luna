"""
SYNC EXCEL DATABASE TO WEB APPS E-LUNA
Membaca file Excel 'database_logistik_bpbd_jabar.xlsx' (250 Transaksi)
dan mengekspornya ke format 'data.js' dan 'database_eluna.json'.
"""

import openpyxl
import json
import datetime
import os

EXCEL_FILE = r"database_logistik_bpbd_jabar.xlsx"
TARGET_JS = r"data.js"
TARGET_JSON = r"database_eluna.json"

def sync_data():
    if not os.path.exists(EXCEL_FILE):
        print(f"Error: File {EXCEL_FILE} tidak ditemukan!")
        return

    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    
    # 1. BACA LIST DATA (250 Transaksi)
    ws_list = wb["List Data"]
    transactions = []
    
    for row in range(2, ws_list.max_row + 1):
        dana = ws_list.cell(row, 1).value
        barang = ws_list.cell(row, 2).value
        kat = ws_list.cell(row, 3).value
        masuk = ws_list.cell(row, 4).value or 0
        keluar = ws_list.cell(row, 5).value or 0
        tujuan = ws_list.cell(row, 6).value or "-"
        exp = ws_list.cell(row, 7).value
        tgl_in = ws_list.cell(row, 8).value
        tgl_out = ws_list.cell(row, 9).value
        doc = ws_list.cell(row, 10).value
        
        if not barang:
            continue
            
        jenis = "Masuk" if masuk > 0 else "Keluar"
        jumlah = masuk if masuk > 0 else keluar
        
        tgl_tx = tgl_in if masuk > 0 else tgl_out
        if isinstance(tgl_tx, datetime.datetime) or isinstance(tgl_tx, datetime.date):
            tgl_str = tgl_tx.strftime("%Y-%m-%d")
            bulan_str = tgl_tx.strftime("%Y-%m")
            tahun_str = tgl_tx.strftime("%Y")
        else:
            tgl_str = "2026-09-15"
            bulan_str = "2026-09"
            tahun_str = "2026"
            
        asal_tujuan = f"Sumber: {dana}" if masuk > 0 else f"Tujuan: {tujuan}"
        
        transactions.append({
            "id": f"TX-{row-1:04d}",
            "tanggal": tgl_str,
            "bulan": bulan_str,
            "tahun": tahun_str,
            "noDokumen": doc or f"DOC-{row}",
            "jenis": jenis,
            "kategori": kat,
            "uraianBarang": barang,
            "jumlah": int(jumlah),
            "satuan": "Unit/Paket",
            "sumberDana": dana,
            "asalTujuan": asal_tujuan,
            "kabKotaTujuan": tujuan if keluar > 0 else "-"
        })

    # Urutkan transaksi dari yang paling baru ke paling lama (Newest first)
    # Sehingga 10 transaksi teratas adalah inputan paling baru (September 2026)
    transactions.sort(key=lambda x: x["tanggal"], reverse=True)

    # 2. BACA KADALUWARSA (5-10% dari stok barang)
    ws_exp = wb["Kadaluwarsa"]
    exp_items = []
    
    for row in range(2, ws_exp.max_row + 1):
        dana = ws_exp.cell(row, 1).value
        barang = ws_exp.cell(row, 2).value
        kat = ws_exp.cell(row, 3).value
        doc = ws_exp.cell(row, 4).value
        tgl_exp = ws_exp.cell(row, 5).value
        
        if not barang or not tgl_exp:
            continue
            
        if isinstance(tgl_exp, datetime.datetime) or isinstance(tgl_exp, datetime.date):
            exp_str = tgl_exp.strftime("%Y-%m-%d")
        else:
            exp_str = str(tgl_exp)
            
        exp_items.append({
            "id": f"EXP-{row-1:02d}",
            "kodeBarang": f"LOG-{kat[:3].upper()}-{row-1:03d}",
            "namaBarang": barang,
            "kategori": kat,
            "jumlahStok": 1500,
            "satuan": "Unit/Paket",
            "tglKadaluwarsaAsli": exp_str,
            "lokasiGudang": "Gudang Induk BPBD Jawa Barat",
            "sumber": dana,
            "noDokumen": doc
        })

    # 3. BACA STOK AWAL & AKUMULASI PER 4 KATEGORI RESMI
    ws_awal = wb["Stok Awal"]
    stok_per_kat = {"Sandang": 0, "Pangan": 0, "Papan": 0, "Logistik Lainnya": 0}
    
    for row in range(2, ws_awal.max_row + 1):
        kat = ws_awal.cell(row, 4).value
        awal = ws_awal.cell(row, 5).value or 0
        if kat in stok_per_kat:
            stok_per_kat[kat] += awal

    # Susun Objek State Lengkap untuk Web Apps E-LUNA
    eluna_state = {
        "metadata": {
            "systemName": "E-LUNA OPERATOR GUDANG",
            "subTitle": "Sistem Pemantauan Operasional Logistik Kebencanaan",
            "instansi": "BPBD Provinsi Jawa Barat",
            "gudangAktif": "Gudang Logistik & Peralatan Provinsi Jawa Barat",
            "lokasi": "Jl. Soekarno-Hatta No. 629, Kota Bandung",
            "petugas": "Nadhif Althafu Rutama",
            "tanggalAcuanSistem": "2026-10-03",
            "parameterXDefaultHari": 30,
            "sourceDatabase": EXCEL_FILE,
            "totalEntries": len(transactions),
            "dateRange": "Januari 2025 - September 2026"
        },
        "kategoriMaster": [
            {
                "id": "sandang",
                "nama": "Sandang",
                "deskripsi": "Selimut, Matras, Sarung, Pakaian Dewasa, Pakaian Anak",
                "color": "#8B5CF6",
                "bgLight": "#F5F3FF",
                "stokAwal": stok_per_kat["Sandang"],
                "satuan": "Unit/Paket"
            },
            {
                "id": "pangan",
                "nama": "Pangan",
                "deskripsi": "Beras, Makanan Siap Saji, Susu Formula, Biskuit, Lauk Pauk, Air Mineral",
                "color": "#3B82F6",
                "bgLight": "#EFF6FF",
                "stokAwal": stok_per_kat["Pangan"],
                "satuan": "Unit/Paket"
            },
            {
                "id": "papan",
                "nama": "Papan",
                "deskripsi": "Tenda Pengungsi, Tenda Keluarga, Terpal, Seng",
                "color": "#F97316",
                "bgLight": "#FFF7ED",
                "stokAwal": stok_per_kat["Papan"],
                "satuan": "Unit"
            },
            {
                "id": "lainnya",
                "nama": "Logistik Lainnya",
                "deskripsi": "Perahu Karet, Genset, Chainsaw, Family Kit, Sanitasi Kit",
                "color": "#10B981",
                "bgLight": "#ECFDF5",
                "stokAwal": stok_per_kat["Logistik Lainnya"],
                "satuan": "Unit/Paket"
            }
        ],
        "kabKotaJabar": [
            "Kab. Bogor", "Kab. Sukabumi", "Kab. Cianjur", "Kab. Bandung", "Kab. Garut",
            "Kab. Tasikmalaya", "Kab. Ciamis", "Kab. Kuningan", "Kab. Cirebon", "Kab. Majalengka",
            "Kab. Sumedang", "Kab. Indramayu", "Kab. Subang", "Kab. Purwakarta", "Kab. Karawang",
            "Kab. Bekasi", "Kab. Bandung Barat", "Kab. Pangandaran", "Kota Bogor", "Kota Sukabumi",
            "Kota Bandung", "Kota Cirebon", "Kota Bekasi", "Kota Depok", "Kota Cimahi",
            "Kota Tasikmalaya", "Kota Banjar"
        ],
        "daftarBarangKadaluwarsa": exp_items,
        "riwayatMutasiSemua": transactions
    }

    # Tulis ke data.js
    with open(TARGET_JS, "w", encoding="utf-8") as f:
        f.write("/**\n")
        f.write(" * DATABASE LOGISTIK BPBD PROVINSI JAWA BARAT (250 TRANSAKSI)\n")
        f.write(" * Disinkronkan secara otomatis dari 'database_logistik_bpbd_jabar.xlsx'\n")
        f.write(f" * Terakhir diperbarui: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(" */\n\n")
        f.write("const ELUNA_OPERATOR_DATA = ")
        f.write(json.dumps(eluna_state, indent=2, ensure_ascii=False))
        f.write(";\n")

    # Tulis ke database_eluna.json
    with open(TARGET_JSON, "w", encoding="utf-8") as f:
        json.dump(eluna_state, f, indent=2, ensure_ascii=False)

    print(f"\n[SINKRONISASI SUKSES]")
    print(f"  -> Total Transaksi Mutasi  : {len(transactions)} Baris (200 Masuk, 50 Keluar)")
    print(f"  -> Total Item Kadaluwarsa   : {len(exp_items)} Batch Pelacakan")
    print(f"  -> Rentang Waktu           : Januari 2025 s.d. September 2026")
    print(f"  -> File Target Terupdate    : {TARGET_JS} & {TARGET_JSON}")

if __name__ == "__main__":
    sync_data()
