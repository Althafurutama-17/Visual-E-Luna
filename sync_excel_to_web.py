"""
SYNC EXCEL DATABASE TO WEB APPS E-LUNA (VERSION 2.0 - KABID & KALAK SUPPORT)
Membaca file Excel 'database_logistik_bpbd_jabar.xlsx' (250 Transaksi, 324 Permohonan Daerah,
kolom valuasi finansial Rupiah) dan mengekspornya ke 'data.js' dan 'database_eluna.json'.
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
    
    # 1. BACA LIST DATA (250 Transaksi + Kolom Finansial)
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
        harga_satuan = ws_list.cell(row, 11).value or 0
        nilai_total = ws_list.cell(row, 12).value or 0
        
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
            "kabKotaTujuan": tujuan if keluar > 0 else "-",
            "hargaSatuan": int(harga_satuan),
            "nilaiTotal": int(nilai_total if nilai_total else int(jumlah) * int(harga_satuan))
        })

    # Urutkan transaksi dari yang paling baru ke paling lama (Newest first)
    transactions.sort(key=lambda x: x["tanggal"], reverse=True)

    # 2. BACA KADALUWARSA
    ws_exp = wb["Kadaluwarsa"]
    exp_items = []
    
    for row in range(2, ws_exp.max_row + 1):
        dana = ws_exp.cell(row, 1).value
        barang = ws_exp.cell(row, 2).value
        kat = ws_exp.cell(row, 3).value
        doc = ws_exp.cell(row, 4).value
        tgl_exp = ws_exp.cell(row, 5).value
        status = ws_exp.cell(row, 6).value
        sisa = ws_exp.cell(row, 7).value
        harga = ws_exp.cell(row, 8).value or 0
        
        if not barang or not tgl_exp:
            continue
            
        if isinstance(tgl_exp, datetime.datetime) or isinstance(tgl_exp, datetime.date):
            exp_str = tgl_exp.strftime("%Y-%m-%d")
        else:
            exp_str = str(tgl_exp)
            
        matched_tx = next((t for t in transactions if t["noDokumen"] == doc), None)
        qty_stok = matched_tx["jumlah"] if matched_tx else 1200

        exp_items.append({
            "id": f"EXP-{row-1:02d}",
            "kodeBarang": f"LOG-{kat[:3].upper()}-{row-1:03d}",
            "namaBarang": barang,
            "kategori": kat,
            "jumlahStok": qty_stok,
            "satuan": "Unit/Paket",
            "tglKadaluwarsaAsli": exp_str,
            "lokasiGudang": "Gudang Induk BPBD Jawa Barat",
            "sumber": dana,
            "noDokumen": doc,
            "hargaSatuan": int(harga),
            "nilaiTotal": int(qty_stok * int(harga))
        })

    # 3. BACA STOK AWAL & AKUMULASI PER 4 KATEGORI RESMI
    ws_awal = wb["Stok Awal"]
    stok_per_kat = {"Sandang": 0, "Pangan": 0, "Papan": 0, "Logistik Lainnya": 0}
    nilai_awal_per_kat = {"Sandang": 0, "Pangan": 0, "Papan": 0, "Logistik Lainnya": 0}
    daftar_stok_awal_barang = []
    
    for row in range(2, ws_awal.max_row + 1):
        sumber = ws_awal.cell(row, 2).value
        barang = ws_awal.cell(row, 3).value
        kat = ws_awal.cell(row, 4).value
        awal = ws_awal.cell(row, 5).value or 0
        harga = ws_awal.cell(row, 6).value or 0
        nilai = ws_awal.cell(row, 7).value or (int(awal) * int(harga))
        
        if kat in stok_per_kat:
            stok_per_kat[kat] += int(awal)
            nilai_awal_per_kat[kat] += int(nilai)
            
        if barang and sumber:
            daftar_stok_awal_barang.append({
                "sumber": str(sumber).strip(),
                "barang": str(barang).strip(),
                "kategori": str(kat).strip(),
                "stokAwal": int(awal),
                "hargaSatuan": int(harga),
                "nilaiTotal": int(nilai)
            })

    # 4. BACA PERMOHONAN DAERAH (PENGUATAN LOGISTIK TAHUNAN 2025 & 2026)
    ws_perm = wb["Permohonan Daerah"]
    daftar_permohonan = []
    
    for row in range(2, ws_perm.max_row + 1):
        no_surat = ws_perm.cell(row, 1).value
        thn = ws_perm.cell(row, 2).value
        tgl_p = ws_perm.cell(row, 3).value
        kab_kota = ws_perm.cell(row, 4).value
        kat = ws_perm.cell(row, 5).value
        barang = ws_perm.cell(row, 6).value
        qty_minta = ws_perm.cell(row, 7).value or 0
        harga_satuan = ws_perm.cell(row, 8).value or 0
        est_nilai = ws_perm.cell(row, 9).value or 0
        
        if not kab_kota or not barang or str(no_surat).startswith("TOTAL"):
            continue
            
        if isinstance(tgl_p, datetime.datetime) or isinstance(tgl_p, datetime.date):
            tgl_str = tgl_p.strftime("%Y-%m-%d")
        else:
            tgl_str = str(tgl_p)
            
        daftar_permohonan.append({
            "id": f"REQ-{row-1:04d}",
            "noSurat": str(no_surat),
            "tahunAnggaran": int(thn),
            "tanggalPermohonan": tgl_str,
            "kabKota": str(kab_kota),
            "kategori": str(kat),
            "barang": str(barang),
            "jumlahDiminta": int(qty_minta),
            "hargaSatuan": int(harga_satuan),
            "estimasiNilai": int(est_nilai if est_nilai else int(qty_minta) * int(harga_satuan))
        })

    # 5. HITUNG REKAPITULASI FINANSIAL MAKRO & SUMBER DANA
    # Transaksi Masuk (Penerimaan) & Keluar (Distribusi)
    total_nilai_masuk = sum(t["nilaiTotal"] for t in transactions if t["jenis"] == "Masuk")
    total_nilai_keluar = sum(t["nilaiTotal"] for t in transactions if t["jenis"] == "Keluar")
    total_nilai_awal = sum(i["nilaiTotal"] for i in daftar_stok_awal_barang)
    
    # Nilai koreksi penyesuaian opname / depresiasi kedaluwarsa
    koreksi_tambah_rp = 450000000 # Rp 450 Juta
    koreksi_kurang_rp = 285000000 # Rp 285 Juta
    
    # Nilai Stok Akhir = Stok Awal + Penerimaan + Koreksi(+) - Koreksi(-) - Distribusi
    total_nilai_akhir = total_nilai_awal + total_nilai_masuk + koreksi_tambah_rp - koreksi_kurang_rp - total_nilai_keluar

    # Sumber Dana Breakdown
    sumber_dana_finansial = {}
    for dana in ["APBN", "APBD", "CSR"]:
        masuk_dana = sum(t["nilaiTotal"] for t in transactions if t["jenis"] == "Masuk" and t["sumberDana"] == dana)
        keluar_dana = sum(t["nilaiTotal"] for t in transactions if t["jenis"] == "Keluar" and t["sumberDana"] == dana)
        awal_dana = sum(i["nilaiTotal"] for i in daftar_stok_awal_barang if i["sumber"] == dana)
        sumber_dana_finansial[dana] = {
            "nilaiAwal": awal_dana,
            "nilaiMasuk": masuk_dana,
            "nilaiKeluar": keluar_dana,
            "nilaiAkhir": awal_dana + masuk_dana - keluar_dana
        }

    # Neraca per Kategori Barang
    neraca_kategori = {}
    for kat in ["Sandang", "Pangan", "Papan", "Logistik Lainnya"]:
        awal_kat = nilai_awal_per_kat.get(kat, 0)
        masuk_kat = sum(t["nilaiTotal"] for t in transactions if t["jenis"] == "Masuk" and t["kategori"] == kat)
        keluar_kat = sum(t["nilaiTotal"] for t in transactions if t["jenis"] == "Keluar" and t["kategori"] == kat)
        
        # Proporsi koreksi
        proporsi = masuk_kat / total_nilai_masuk if total_nilai_masuk > 0 else 0.25
        kor_tambah = int(koreksi_tambah_rp * proporsi)
        kor_kurang = int(koreksi_kurang_rp * proporsi)
        
        akhir_kat = awal_kat + masuk_kat + kor_tambah - kor_kurang - keluar_kat
        
        neraca_kategori[kat] = {
            "nilaiAwal": awal_kat,
            "nilaiMasuk": masuk_kat,
            "koreksiTambah": kor_tambah,
            "koreksiKurang": kor_kurang,
            "koreksiNetto": kor_tambah - kor_kurang,
            "nilaiKeluar": keluar_kat,
            "nilaiAkhir": akhir_kat
        }

    # Susun Objek State Lengkap untuk Web Apps E-LUNA
    eluna_state = {
        "metadata": {
            "systemName": "E-LUNA LOGISTIK BPBD JAWA BARAT",
            "subTitle": "Sistem Informasi & Pemantauan Operasional Logistik Kebencanaan",
            "instansi": "BPBD Provinsi Jawa Barat",
            "gudangAktif": "Gudang Logistik & Peralatan Provinsi Jawa Barat",
            "lokasi": "Jl. Soekarno-Hatta No. 629, Kota Bandung",
            "petugas": "Nadhif Althafu Rutama",
            "tanggalAcuanSistem": "2026-10-03",
            "parameterXDefaultHari": 60,
            "sourceDatabase": EXCEL_FILE,
            "totalEntries": len(transactions),
            "dateRange": "Januari 2025 - September 2026",
            "fokusOperasional": "Penguatan Cadangan Logistik Kesiapsiagaan Daerah Tahunan"
        },
        "kategoriMaster": [
            {
                "id": "sandang",
                "nama": "Sandang",
                "deskripsi": "Selimut, Matras, Sarung, Pakaian Dewasa, Pakaian Anak",
                "color": "#8B5CF6",
                "bgLight": "#F5F3FF",
                "stokAwal": stok_per_kat["Sandang"],
                "nilaiStokAwal": nilai_awal_per_kat["Sandang"],
                "satuan": "Unit/Paket"
            },
            {
                "id": "pangan",
                "nama": "Pangan",
                "deskripsi": "Beras, Makanan Siap Saji, Susu Formula, Biskuit, Lauk Pauk, Air Mineral",
                "color": "#3B82F6",
                "bgLight": "#EFF6FF",
                "stokAwal": stok_per_kat["Pangan"],
                "nilaiStokAwal": nilai_awal_per_kat["Pangan"],
                "satuan": "Unit/Paket"
            },
            {
                "id": "papan",
                "nama": "Papan",
                "deskripsi": "Tenda Pengungsi, Tenda Keluarga, Terpal, Seng",
                "color": "#F97316",
                "bgLight": "#FFF7ED",
                "stokAwal": stok_per_kat["Papan"],
                "nilaiStokAwal": nilai_awal_per_kat["Papan"],
                "satuan": "Unit"
            },
            {
                "id": "lainnya",
                "nama": "Logistik Lainnya",
                "deskripsi": "Perahu Karet, Genset, Chainsaw, Family Kit, Sanitasi Kit",
                "color": "#10B981",
                "bgLight": "#ECFDF5",
                "stokAwal": stok_per_kat["Logistik Lainnya"],
                "nilaiStokAwal": nilai_awal_per_kat["Logistik Lainnya"],
                "satuan": "Unit/Paket"
            }
        ],
        "kabKotaJabar": [
            "Kabupaten Bandung", "Kabupaten Bandung Barat", "Kabupaten Bekasi", "Kabupaten Bogor",
            "Kabupaten Ciamis", "Kabupaten Cianjur", "Kabupaten Cirebon", "Kabupaten Garut",
            "Kabupaten Indramayu", "Kabupaten Karawang", "Kabupaten Kuningan", "Kabupaten Majalengka",
            "Kabupaten Pangandaran", "Kabupaten Purwakarta", "Kabupaten Subang", "Kabupaten Sukabumi",
            "Kabupaten Sumedang", "Kabupaten Tasikmalaya", "Kota Bandung", "Kota Banjar",
            "Kota Bekasi", "Kota Bogor", "Kota Cimahi", "Kota Cirebon", "Kota Depok",
            "Kota Sukabumi", "Kota Tasikmalaya"
        ],
        "rekapFinansialMakro": {
            "totalNilaiStokAwal": total_nilai_awal,
            "totalNilaiPenerimaan": total_nilai_masuk,
            "totalNilaiDistribusi": total_nilai_keluar,
            "koreksiTambahRp": koreksi_tambah_rp,
            "koreksiKurangRp": koreksi_kurang_rp,
            "totalNilaiStokAkhir": total_nilai_akhir,
            "sumberDana": sumber_dana_finansial,
            "neracaKategori": neraca_kategori
        },
        "daftarStokAwalPerBarang": daftar_stok_awal_barang,
        "daftarBarangKadaluwarsa": exp_items,
        "daftarPermohonanDaerah": daftar_permohonan,
        "riwayatMutasiSemua": transactions
    }

    # Tulis ke data.js
    with open(TARGET_JS, "w", encoding="utf-8") as f:
        f.write("/**\n")
        f.write(" * DATABASE LOGISTIK BPBD PROVINSI JAWA BARAT\n")
        f.write(" * Terpadu: View Operasional, View Kabid (Kuantitas Penguatan), View KALAK (Finansial Makro)\n")
        f.write(" * Disinkronkan secara otomatis dari 'database_logistik_bpbd_jabar.xlsx'\n")
        f.write(f" * Terakhir diperbarui: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(" */\n\n")
        f.write("var ELUNA_OPERATOR_DATA = ")
        f.write(json.dumps(eluna_state, indent=2, ensure_ascii=False))
        f.write(";\n\n")
        f.write("if (typeof window !== 'undefined') {\n")
        f.write("  window.ELUNA_OPERATOR_DATA = ELUNA_OPERATOR_DATA;\n")
        f.write("}\n")

    # Tulis ke database_eluna.json
    with open(TARGET_JSON, "w", encoding="utf-8") as f:
        json.dump(eluna_state, f, indent=2, ensure_ascii=False)

    print(f"\n[SINKRONISASI SUKSES]")
    print(f"  -> Total Transaksi Mutasi    : {len(transactions)} Baris")
    print(f"  -> Total Permohonan Daerah   : {len(daftar_permohonan)} Baris (Penguatan Tahunan)")
    print(f"  -> Total Nilai Penerimaan    : Rp {total_nilai_masuk:,.0f}")
    print(f"  -> Total Nilai Distribusi    : Rp {total_nilai_keluar:,.0f}")
    print(f"  -> Total Nilai Stok Akhir    : Rp {total_nilai_akhir:,.0f}")
    print(f"  -> File Target Terupdate      : {TARGET_JS} & {TARGET_JSON}")

if __name__ == "__main__":
    sync_data()
