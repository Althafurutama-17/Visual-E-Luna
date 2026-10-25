"""
UPDATE EXCEL DATABASE STRUCTURE (OPSI A)
Menambahkan kolom valuasi finansial (Harga Satuan Rp & Nilai Total Rp) ke seluruh sheet transaksi,
serta menambahkan sheet baru 'Permohonan Daerah' (Pagu Tahunan 2025 & 2026 Penguatan Logistik)
ke berkas 'database_logistik_bpbd_jabar.xlsx'.
"""

import os
import datetime
import random
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation

EXCEL_FILE = r"database_logistik_bpbd_jabar.xlsx"

# STANDAR SATUAN HARGA (SSH) LOGISTIK BPBD JAWA BARAT
SSH_PRICES = {
    # Pangan
    "Beras": 15000,
    "Makanan Siap Saji": 45000,
    "Lauk Pauk": 35000,
    "Air Mineral": 30000,
    "Biskuit": 25000,
    "Susu Formula": 65000,
    # Sandang
    "Selimut": 85000,
    "Matras": 120000,
    "Pakaian Dewasa": 95000,
    "Pakaian Anak": 75000,
    "Sarung": 60000,
    # Papan
    "Tenda Pengungsi": 8500000,
    "Tenda Keluarga": 3200000,
    "Terpal": 150000,
    "Seng": 90000,
    # Logistik Lainnya
    "Perahu Karet": 35000000,
    "Genset": 12500000,
    "Chainsaw": 4500000,
    "Family Kit": 350000,
    "Sanitasi Kit": 250000
}

def get_unit_price(barang_name):
    for k, v in SSH_PRICES.items():
        if k.lower() in str(barang_name).lower():
            return v
    return 50000

def run_update():
    if not os.path.exists(EXCEL_FILE):
        print(f"File {EXCEL_FILE} tidak ditemukan!")
        return

    print("Membaca basis data Excel saat ini...")
    wb_old = load_workbook(EXCEL_FILE, data_only=True)

    # 1. BACA DATA LAMA
    # -----------------
    # _RefData
    ws_ref_old = wb_old["_RefData"]
    sumber_dana = [ws_ref_old.cell(r, 1).value for r in range(2, ws_ref_old.max_row + 1) if ws_ref_old.cell(r, 1).value]
    kategori_barang = [ws_ref_old.cell(r, 2).value for r in range(2, ws_ref_old.max_row + 1) if ws_ref_old.cell(r, 2).value]
    kab_kota_jabar = [ws_ref_old.cell(r, 3).value for r in range(2, ws_ref_old.max_row + 1) if ws_ref_old.cell(r, 3).value]

    # Penerimaan (200 baris)
    ws_pen_old = wb_old["Penerimaan"]
    penerimaan_data = []
    for r in range(2, ws_pen_old.max_row + 1):
        dana = ws_pen_old.cell(r, 1).value
        barang = ws_pen_old.cell(r, 2).value
        kat = ws_pen_old.cell(r, 3).value
        qty_in = ws_pen_old.cell(r, 4).value
        tgl_exp = ws_pen_old.cell(r, 5).value
        tgl_in = ws_pen_old.cell(r, 6).value
        doc = ws_pen_old.cell(r, 7).value
        if barang and qty_in and str(dana).strip() != "TOTAL PENERIMAAN":
            price = get_unit_price(barang)
            penerimaan_data.append({
                "dana": dana,
                "barang": barang,
                "kat": kat,
                "in": int(qty_in),
                "exp": tgl_exp,
                "tgl_in": tgl_in,
                "doc": doc,
                "harga_satuan": price,
                "nilai_total": int(qty_in) * price
            })

    # Target persentase pemenuhan per daerah (7 Aman ≥90%, 20 Kurang <90%, batas bawah 60%)
    target_pcts = {
        "Kabupaten Bogor": 0.975,
        "Kabupaten Bekasi": 0.960,
        "Kabupaten Garut": 0.948,
        "Kota Depok": 0.935,
        "Kabupaten Subang": 0.922,
        "Kabupaten Sukabumi": 0.910,
        "Kota Banjar": 0.902,
        "Kabupaten Indramayu": 0.885,
        "Kota Bogor": 0.868,
        "Kota Sukabumi": 0.850,
        "Kabupaten Pangandaran": 0.832,
        "Kabupaten Karawang": 0.815,
        "Kabupaten Tasikmalaya": 0.798,
        "Kabupaten Majalengka": 0.780,
        "Kota Cimahi": 0.762,
        "Kabupaten Purwakarta": 0.745,
        "Kabupaten Cirebon": 0.728,
        "Kota Cirebon": 0.710,
        "Kota Bandung": 0.695,
        "Kabupaten Bandung Barat": 0.680,
        "Kabupaten Ciamis": 0.665,
        "Kota Bekasi": 0.650,
        "Kota Tasikmalaya": 0.640,
        "Kabupaten Bandung": 0.630,
        "Kabupaten Cianjur": 0.620,
        "Kabupaten Kuningan": 0.610,
        "Kabupaten Sumedang": 0.600
    }
    all_daerah_ordered = list(target_pcts.keys())

    # Distribusi (50 baris)
    ws_dis_old = wb_old["Distribusi"]
    distribusi_raw = []
    for r in range(2, ws_dis_old.max_row + 1):
        dana = ws_dis_old.cell(r, 1).value
        barang = ws_dis_old.cell(r, 2).value
        kat = ws_dis_old.cell(r, 3).value
        qty_out = ws_dis_old.cell(r, 4).value
        tujuan = ws_dis_old.cell(r, 5).value
        tgl_out = ws_dis_old.cell(r, 6).value
        doc = ws_dis_old.cell(r, 7).value
        if barang and qty_out and str(dana).strip() != "TOTAL DISTRIBUSI":
            price = get_unit_price(barang)
            distribusi_raw.append({
                "dana": dana,
                "barang": barang,
                "kat": kat,
                "out": int(qty_out),
                "tgl_out": tgl_out,
                "doc": doc,
                "harga_satuan": price,
                "nilai_total": int(qty_out) * price
            })

    # Urutkan distribusi berdasarkan kuantitas keluar terbesar ke terkecil
    distribusi_raw.sort(key=lambda x: x["out"], reverse=True)

    # Alokasikan 50 transaksi secara merata ke 27 daerah agar seluruh 27 daerah menerima penyaluran
    distribusi_data = []
    for i in range(27):
        item = distribusi_raw[i]
        item["tujuan"] = all_daerah_ordered[i]
        distribusi_data.append(item)
    for i in range(23):
        item = distribusi_raw[27 + i]
        item["tujuan"] = all_daerah_ordered[i]
        distribusi_data.append(item)

    # Kembalikan urutan kronologis berdasarkan tanggal keluar
    distribusi_data.sort(key=lambda x: x["tgl_out"])

    # Stok Awal (60 baris)
    ws_awal_old = wb_old["Stok Awal"]
    stok_awal_data = []
    for r in range(2, ws_awal_old.max_row + 1):
        kunci = ws_awal_old.cell(r, 1).value
        dana = ws_awal_old.cell(r, 2).value
        barang = ws_awal_old.cell(r, 3).value
        kat = ws_awal_old.cell(r, 4).value
        stok_awal = ws_awal_old.cell(r, 5).value
        if barang and stok_awal is not None:
            price = get_unit_price(barang)
            stok_awal_data.append({
                "kunci": kunci,
                "dana": dana,
                "barang": barang,
                "kat": kat,
                "awal": int(stok_awal),
                "harga_satuan": price,
                "nilai_total": int(stok_awal) * price
            })

    # Kadaluwarsa
    ws_exp_old = wb_old["Kadaluwarsa"]
    kadaluwarsa_data = []
    for r in range(2, ws_exp_old.max_row + 1):
        dana = ws_exp_old.cell(r, 1).value
        barang = ws_exp_old.cell(r, 2).value
        kat = ws_exp_old.cell(r, 3).value
        doc = ws_exp_old.cell(r, 4).value
        exp = ws_exp_old.cell(r, 5).value
        status = ws_exp_old.cell(r, 6).value
        sisa = ws_exp_old.cell(r, 7).value
        if barang and doc:
            price = get_unit_price(barang)
            kadaluwarsa_data.append({
                "dana": dana,
                "barang": barang,
                "kat": kat,
                "doc": doc,
                "exp": exp,
                "status": status,
                "sisa": sisa,
                "harga_satuan": price
            })

    # 2. GENERATE DATA BARU: Permohonan Daerah (Penguatan Logistik Tahunan)
    # -------------------------------------------------------------------
    # Fokus: Program Penguatan Logistik Cadangan Daerah (Buffer Stock Tahunan)
    # 27 Kab/Kota se-Jawa Barat x 2 Tahun Anggaran (2025 & 2026)
    permohonan_data = []

    # Hitung total realisasi distribusi per daerah & per komoditas
    realisasi_per_daerah_tahun = {}
    dist_by_daerah = {}
    for item in distribusi_data:
        tgl = item["tgl_out"]
        thn = tgl.year if hasattr(tgl, "year") else 2026
        daerah = item["tujuan"]
        barang = item["barang"]
        kat = item["kat"]
        qty = item["out"]
        key = (daerah, thn, barang, kat)
        realisasi_per_daerah_tahun[key] = realisasi_per_daerah_tahun.get(key, 0) + qty
        dist_by_daerah[daerah] = dist_by_daerah.get(daerah, 0) + qty

    nomor_surat_counter = 1

    for daerah in all_daerah_ordered:
        prefix_daerah = "KOTA" if "Kota" in daerah else "KAB"
        target_p = target_pcts[daerah]
        real_d_total = dist_by_daerah.get(daerah, 0)

        # Total permohonan yang ditargetkan untuk daerah ini agar rasio pemenuhan persis target_p
        minta_d_total = int(round(real_d_total / target_p / 10) * 10)

        items_d = [(thn, b, k, qty) for (kota, thn, b, k), qty in realisasi_per_daerah_tahun.items() if kota == daerah]
        assigned_qty = 0
        added_keys = set()

        # 1. Buat permohonan untuk komoditas yang memang didistribusikan
        for thn, b, k, qty in items_d:
            no_surat = f"360/{nomor_surat_counter:03d}/BPBD-{prefix_daerah}/I/{thn}"
            nomor_surat_counter += 1
            tgl_surat = datetime.date(thn, 1, 15)

            # Untuk daerah aman, komoditas yang disalurkan dipenuhi hampir penuh (95%-100%)
            # Untuk daerah kurang, komoditas yang disalurkan dipenuhi sebagian
            com_ratio = target_p if target_p >= 0.9 else min(1.0, target_p * 1.15)
            req_item = max(qty, int(round(qty / com_ratio / 10) * 10))
            assigned_qty += req_item
            added_keys.add(b)
            price = get_unit_price(b)

            permohonan_data.append({
                "no_surat": no_surat,
                "tahun_anggaran": thn,
                "tgl_permohonan": tgl_surat + datetime.timedelta(days=random.randint(0, 10)),
                "kab_kota": daerah,
                "kategori": k,
                "barang": b,
                "jumlah_diminta": req_item,
                "harga_satuan": price,
                "estimasi_nilai": req_item * price
            })

        # 2. Penuhi sisa kuota (gap) dengan komoditas kategori lain (Sandang, Pangan, Papan, Lainnya)
        # sehingga setiap daerah memiliki rincian komoditas multi-kategori
        rem = minta_d_total - assigned_qty
        all_candidates = [
            ("Beras", "Pangan"),
            ("Makanan Siap Saji", "Pangan"),
            ("Selimut", "Sandang"),
            ("Matras", "Sandang"),
            ("Terpal", "Papan"),
            ("Family Kit", "Logistik Lainnya")
        ]
        needed_slots = [c for c in all_candidates if c[0] not in added_keys][:3]

        if rem > 0 and needed_slots:
            piece = int(round(rem / len(needed_slots) / 10) * 10)
            for i, (b, k) in enumerate(needed_slots):
                no_surat = f"360/{nomor_surat_counter:03d}/BPBD-{prefix_daerah}/I/2026"
                nomor_surat_counter += 1
                tgl_surat = datetime.date(2026, 1, 15)
                q = rem - (piece * (len(needed_slots) - 1)) if i == len(needed_slots) - 1 else piece
                if q <= 0:
                    continue
                price = get_unit_price(b)

                permohonan_data.append({
                    "no_surat": no_surat,
                    "tahun_anggaran": 2026,
                    "tgl_permohonan": tgl_surat + datetime.timedelta(days=random.randint(0, 10)),
                    "kab_kota": daerah,
                    "kategori": k,
                    "barang": b,
                    "jumlah_diminta": q,
                    "harga_satuan": price,
                    "estimasi_nilai": q * price
                })
        elif rem < 0 and permohonan_data:
            permohonan_data[-1]["jumlah_diminta"] += rem
            permohonan_data[-1]["estimasi_nilai"] = permohonan_data[-1]["jumlah_diminta"] * permohonan_data[-1]["harga_satuan"]

    # Urutkan berdasarkan Tahun Anggaran lalu Kab/Kota
    permohonan_data.sort(key=lambda x: (x["tahun_anggaran"], x["kab_kota"]))

    print(f"Total baris Permohonan Daerah berhasil disusun: {len(permohonan_data)} baris (27 Kab/Kota, 2025 & 2026)")
    print(f"Total Akumulatif Unit Permohonan: {sum(r['jumlah_diminta'] for r in permohonan_data):,} Unit")

    # 3. SUSUN ULANG WORKBOOK BARU (DENGAN STYLING LENGKAP & TABEL BERSIH)
    # ------------------------------------------------------------------
    wb_new = Workbook()
    
    FONT_FAMILY = "Segoe UI"
    font_header = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="0E63AA", end_color="0E63AA", fill_type="solid")
    font_data = Font(name=FONT_FAMILY, size=10, color="1E293B")
    font_bold = Font(name=FONT_FAMILY, size=10, bold=True, color="1E293B")
    font_code = Font(name="Consolas", size=9, bold=True, color="0E63AA")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_total = PatternFill(start_color="EEF2F6", end_color="EEF2F6", fill_type="solid")

    thin_border_side = Side(style='thin', color="CBD5E1")
    double_bottom_side = Side(style='double', color="0E63AA")
    thick_top_side = Side(style='thin', color="0E63AA")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_total = Border(top=thick_top_side, bottom=double_bottom_side, left=thin_border_side, right=thin_border_side)

    align_left = Alignment(horizontal='left', vertical='center')
    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    align_header = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # -------------------------------------------------------------
    # SHEET: _RefData
    # -------------------------------------------------------------
    ws_ref = wb_new.active
    ws_ref.title = "_RefData"
    ws_ref.views.sheetView[0].showGridLines = True
    
    ws_ref['A1'] = "Sumber Pendanaan"
    for r_idx, val in enumerate(sumber_dana, start=2):
        ws_ref[f'A{r_idx}'] = val
        
    ws_ref['B1'] = "Kategori Barang"
    for r_idx, val in enumerate(kategori_barang, start=2):
        ws_ref[f'B{r_idx}'] = val
        
    ws_ref['C1'] = "Kab/Kota Jawa Barat"
    for r_idx, val in enumerate(kab_kota_jabar, start=2):
        ws_ref[f'C{r_idx}'] = val

    ws_ref['D1'] = "Tahun Anggaran"
    ws_ref['D2'] = 2025
    ws_ref['D3'] = 2026

    # Validasi
    dv_dana = DataValidation(type="list", formula1="'_RefData'!$A$2:$A$4", allow_blank=True)
    dv_kat = DataValidation(type="list", formula1="'_RefData'!$B$2:$B$5", allow_blank=True)
    dv_tujuan = DataValidation(type="list", formula1="'_RefData'!$C$2:$C$29", allow_blank=True)
    dv_thn = DataValidation(type="list", formula1="'_RefData'!$D$2:$D$3", allow_blank=True)

    # -------------------------------------------------------------
    # SHEET 1: List Data (tblListData - 250 Entri + Kolom Harga)
    # -------------------------------------------------------------
    ws_list = wb_new.create_sheet(title="List Data")
    ws_list.views.sheetView[0].showGridLines = True
    ws_list.freeze_panes = "A2"
    ws_list.add_data_validation(dv_dana)
    ws_list.add_data_validation(dv_kat)
    ws_list.add_data_validation(dv_tujuan)

    headers_list = [
        "Sumber Pendanaan",
        "Barang",
        "Kategori Barang",
        "Jumlah Masuk",
        "Jumlah Keluar",
        "Kab/Kota Tujuan",
        "Tanggal Kadaluwarsa",
        "Tanggal barang diterima",
        "Tanggal barang dikeluarkan",
        "Nomor Dokumen",
        "Harga Satuan (Rp)",
        "Nilai Total (Rp)"
    ]
    ws_list.append(headers_list)
    ws_list.row_dimensions[1].height = 26

    # Gabungkan penerimaan & distribusi kronologis
    combined_tx = []
    for p in penerimaan_data:
        combined_tx.append({
            "dana": p["dana"], "barang": p["barang"], "kat": p["kat"],
            "in": p["in"], "out": 0, "tujuan": "-", "exp": p["exp"],
            "tgl_in": p["tgl_in"], "tgl_out": None, "doc": p["doc"],
            "harga": p["harga_satuan"], "nilai": p["nilai_total"],
            "tgl_sort": p["tgl_in"]
        })
    for d in distribusi_data:
        combined_tx.append({
            "dana": d["dana"], "barang": d["barang"], "kat": d["kat"],
            "in": 0, "out": d["out"], "tujuan": d["tujuan"], "exp": None,
            "tgl_in": None, "tgl_out": d["tgl_out"], "doc": d["doc"],
            "harga": d["harga_satuan"], "nilai": d["nilai_total"],
            "tgl_sort": d["tgl_out"]
        })
    combined_tx.sort(key=lambda x: x["tgl_sort"])

    for r_idx, item in enumerate(combined_tx, start=2):
        row_vals = [
            item["dana"], item["barang"], item["kat"], item["in"], item["out"],
            item["tujuan"], item["exp"], item["tgl_in"], item["tgl_out"], item["doc"],
            item["harga"], item["nilai"]
        ]
        ws_list.append(row_vals)
        ws_list.row_dimensions[r_idx].height = 20

    max_row_list = len(combined_tx) + 1
    for c_idx in range(1, 13):
        h = ws_list.cell(row=1, column=c_idx)
        h.font = font_header
        h.fill = fill_header
        h.alignment = align_header
        h.border = border_cell

    for r_idx in range(2, max_row_list + 1):
        is_even = (r_idx % 2 == 0)
        for c_idx in range(1, 13):
            c = ws_list.cell(row=r_idx, column=c_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            if c_idx in [1, 3, 6]:
                c.alignment = align_center
            elif c_idx in [4, 5]:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold
            elif c_idx in [7, 8, 9]:
                c.alignment = align_center
                if c.value is not None and hasattr(c.value, "strftime"):
                    c.number_format = 'DD/MM/YYYY'
            elif c_idx == 10:
                c.font = font_code
                c.alignment = align_center
            elif c_idx in [11, 12]:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold

    dv_dana.add(f"A2:A{max_row_list}")
    dv_kat.add(f"C2:C{max_row_list}")
    dv_tujuan.add(f"F2:F{max_row_list}")

    tab_list = Table(displayName="tblListData", ref=f"A1:L{max_row_list}")
    tab_list.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_list.add_table(tab_list)

    # -------------------------------------------------------------
    # SHEET 2: Penerimaan (tblPenerimaan - 200 Entri + Kolom Harga)
    # -------------------------------------------------------------
    ws_pen = wb_new.create_sheet(title="Penerimaan")
    ws_pen.views.sheetView[0].showGridLines = True
    ws_pen.freeze_panes = "A2"
    ws_pen.add_data_validation(dv_dana)
    ws_pen.add_data_validation(dv_kat)

    headers_pen = [
        "Sumber Pendanaan",
        "Barang",
        "Kategori Barang",
        "Jumlah Masuk",
        "Tanggal Kadaluwarsa",
        "Tanggal barang diterima",
        "Nomor Dokumen",
        "Harga Satuan (Rp)",
        "Nilai Total (Rp)"
    ]
    ws_pen.append(headers_pen)
    ws_pen.row_dimensions[1].height = 26

    for r_idx, item in enumerate(penerimaan_data, start=2):
        row_vals = [
            item["dana"], item["barang"], item["kat"], item["in"], item["exp"],
            item["tgl_in"], item["doc"], item["harga_satuan"], item["nilai_total"]
        ]
        ws_pen.append(row_vals)
        ws_pen.row_dimensions[r_idx].height = 20

    max_row_pen = len(penerimaan_data) + 1
    for c_idx in range(1, 10):
        h = ws_pen.cell(row=1, column=c_idx)
        h.font = font_header
        h.fill = fill_header
        h.alignment = align_header
        h.border = border_cell

    for r_idx in range(2, max_row_pen + 1):
        is_even = (r_idx % 2 == 0)
        for c_idx in range(1, 10):
            c = ws_pen.cell(row=r_idx, column=c_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            if c_idx in [1, 3]:
                c.alignment = align_center
            elif c_idx == 4:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold
            elif c_idx in [5, 6]:
                c.alignment = align_center
                if c.value is not None and hasattr(c.value, "strftime"):
                    c.number_format = 'DD/MM/YYYY'
            elif c_idx == 7:
                c.font = font_code
                c.alignment = align_center
            elif c_idx in [8, 9]:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold

    dv_dana.add(f"A2:A{max_row_pen}")
    dv_kat.add(f"C2:C{max_row_pen}")

    tab_pen = Table(displayName="tblPenerimaan", ref=f"A1:I{max_row_pen}")
    tab_pen.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_pen.add_table(tab_pen)

    # Baris Total Penerimaan
    tot_row_p = max_row_pen + 1
    ws_pen.cell(row=tot_row_p, column=1, value="TOTAL PENERIMAAN").font = font_bold
    ws_pen.cell(row=tot_row_p, column=1).alignment = align_center
    ws_pen.cell(row=tot_row_p, column=4, value=f"=SUM(D2:D{max_row_pen})").number_format = '#,##0'
    ws_pen.cell(row=tot_row_p, column=4).font = font_bold
    ws_pen.cell(row=tot_row_p, column=4).alignment = align_right
    ws_pen.cell(row=tot_row_p, column=9, value=f"=SUM(I2:I{max_row_pen})").number_format = '#,##0'
    ws_pen.cell(row=tot_row_p, column=9).font = font_bold
    ws_pen.cell(row=tot_row_p, column=9).alignment = align_right
    for c_idx in range(1, 10):
        cell_t = ws_pen.cell(row=tot_row_p, column=c_idx)
        cell_t.fill = fill_total
        cell_t.border = border_total

    # -------------------------------------------------------------
    # SHEET 3: Distribusi (tblDistribusi - 50 Entri + Kolom Harga)
    # -------------------------------------------------------------
    ws_dis = wb_new.create_sheet(title="Distribusi")
    ws_dis.views.sheetView[0].showGridLines = True
    ws_dis.freeze_panes = "A2"
    ws_dis.add_data_validation(dv_dana)
    ws_dis.add_data_validation(dv_kat)
    ws_dis.add_data_validation(dv_tujuan)

    headers_dis = [
        "Sumber Pendanaan",
        "Barang",
        "Kategori Barang",
        "Jumlah Keluar",
        "Kab/Kota Tujuan",
        "Tanggal barang dikeluarkan",
        "Nomor Dokumen",
        "Harga Satuan (Rp)",
        "Nilai Total (Rp)"
    ]
    ws_dis.append(headers_dis)
    ws_dis.row_dimensions[1].height = 26

    for r_idx, item in enumerate(distribusi_data, start=2):
        row_vals = [
            item["dana"], item["barang"], item["kat"], item["out"], item["tujuan"],
            item["tgl_out"], item["doc"], item["harga_satuan"], item["nilai_total"]
        ]
        ws_dis.append(row_vals)
        ws_dis.row_dimensions[r_idx].height = 20

    max_row_dis = len(distribusi_data) + 1
    for c_idx in range(1, 10):
        h = ws_dis.cell(row=1, column=c_idx)
        h.font = font_header
        h.fill = fill_header
        h.alignment = align_header
        h.border = border_cell

    for r_idx in range(2, max_row_dis + 1):
        is_even = (r_idx % 2 == 0)
        for c_idx in range(1, 10):
            c = ws_dis.cell(row=r_idx, column=c_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            if c_idx in [1, 3, 5]:
                c.alignment = align_center
            elif c_idx == 4:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold
            elif c_idx == 6:
                c.alignment = align_center
                if c.value is not None and hasattr(c.value, "strftime"):
                    c.number_format = 'DD/MM/YYYY'
            elif c_idx == 7:
                c.font = font_code
                c.alignment = align_center
            elif c_idx in [8, 9]:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold

    dv_dana.add(f"A2:A{max_row_dis}")
    dv_kat.add(f"C2:C{max_row_dis}")
    dv_tujuan.add(f"E2:E{max_row_dis}")

    tab_dis = Table(displayName="tblDistribusi", ref=f"A1:I{max_row_dis}")
    tab_dis.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_dis.add_table(tab_dis)

    # Baris Total Distribusi
    tot_row_d = max_row_dis + 1
    ws_dis.cell(row=tot_row_d, column=1, value="TOTAL DISTRIBUSI").font = font_bold
    ws_dis.cell(row=tot_row_d, column=1).alignment = align_center
    ws_dis.cell(row=tot_row_d, column=4, value=f"=SUM(D2:D{max_row_dis})").number_format = '#,##0'
    ws_dis.cell(row=tot_row_d, column=4).font = font_bold
    ws_dis.cell(row=tot_row_d, column=4).alignment = align_right
    ws_dis.cell(row=tot_row_d, column=9, value=f"=SUM(I2:I{max_row_dis})").number_format = '#,##0'
    ws_dis.cell(row=tot_row_d, column=9).font = font_bold
    ws_dis.cell(row=tot_row_d, column=9).alignment = align_right
    for c_idx in range(1, 10):
        cell_t = ws_dis.cell(row=tot_row_d, column=c_idx)
        cell_t.fill = fill_total
        cell_t.border = border_total

    # -------------------------------------------------------------
    # SHEET 4: Stok Awal (tblStokAwal - 60 Entri + Kolom Harga)
    # -------------------------------------------------------------
    ws_awal = wb_new.create_sheet(title="Stok Awal")
    ws_awal.views.sheetView[0].showGridLines = True
    ws_awal.freeze_panes = "A2"
    ws_awal.add_data_validation(dv_dana)
    ws_awal.add_data_validation(dv_kat)

    headers_awal = [
        "Kunci",
        "Sumber Pendanaan",
        "Barang",
        "Kategori Barang",
        "Stok Akhir Periode Sebelumnya",
        "Harga Satuan (Rp)",
        "Nilai Total (Rp)"
    ]
    ws_awal.append(headers_awal)
    ws_awal.row_dimensions[1].height = 26

    for r_idx, item in enumerate(stok_awal_data, start=2):
        row_vals = [
            f'=B{r_idx}&"|"&C{r_idx}&"|"&D{r_idx}',
            item["dana"], item["barang"], item["kat"], item["awal"],
            item["harga_satuan"], item["nilai_total"]
        ]
        ws_awal.append(row_vals)
        ws_awal.row_dimensions[r_idx].height = 20

    max_row_awal = len(stok_awal_data) + 1
    for c_idx in range(1, 8):
        h = ws_awal.cell(row=1, column=c_idx)
        h.font = font_header
        h.fill = fill_header
        h.alignment = align_header
        h.border = border_cell

    for r_idx in range(2, max_row_awal + 1):
        is_even = (r_idx % 2 == 0)
        for c_idx in range(1, 8):
            c = ws_awal.cell(row=r_idx, column=c_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            if c_idx == 1:
                c.font = font_code
                c.alignment = align_center
            elif c_idx in [2, 4]:
                c.alignment = align_center
            elif c_idx in [5, 6, 7]:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold

    dv_dana.add(f"B2:B{max_row_awal}")
    dv_kat.add(f"D2:D{max_row_awal}")

    tab_awal = Table(displayName="tblStokAwal", ref=f"A1:G{max_row_awal}")
    tab_awal.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_awal.add_table(tab_awal)

    # Baris Total Stok Awal
    tot_row_a = max_row_awal + 1
    ws_awal.cell(row=tot_row_a, column=2, value="TOTAL STOK AWAL").font = font_bold
    ws_awal.cell(row=tot_row_a, column=2).alignment = align_center
    ws_awal.cell(row=tot_row_a, column=5, value=f"=SUM(E2:E{max_row_awal})").number_format = '#,##0'
    ws_awal.cell(row=tot_row_a, column=5).font = font_bold
    ws_awal.cell(row=tot_row_a, column=5).alignment = align_right
    ws_awal.cell(row=tot_row_a, column=7, value=f"=SUM(G2:G{max_row_awal})").number_format = '#,##0'
    ws_awal.cell(row=tot_row_a, column=7).font = font_bold
    ws_awal.cell(row=tot_row_a, column=7).alignment = align_right
    for c_idx in range(1, 8):
        cell_t = ws_awal.cell(row=tot_row_a, column=c_idx)
        cell_t.fill = fill_total
        cell_t.border = border_total

    # -------------------------------------------------------------
    # SHEET 5: Stok (tblStok - Formula Dinamis + Nilai Finansial)
    # -------------------------------------------------------------
    ws_stok = wb_new.create_sheet(title="Stok")
    ws_stok.views.sheetView[0].showGridLines = True
    ws_stok.freeze_panes = "A2"
    ws_stok.add_data_validation(dv_dana)
    ws_stok.add_data_validation(dv_kat)

    headers_stok = [
        "Sumber Pendanaan",
        "Barang",
        "Kategori Barang",
        "Stok Fisik",
        "Harga Satuan (Rp)",
        "Nilai Total Stok (Rp)"
    ]
    ws_stok.append(headers_stok)
    ws_stok.row_dimensions[1].height = 26

    for r_idx, item in enumerate(stok_awal_data, start=2):
        formula_stok = (
            f"='Stok Awal'!E{r_idx} + "
            f"SUMIFS(Penerimaan!D:D, Penerimaan!A:A, A{r_idx}, Penerimaan!B:B, B{r_idx}, Penerimaan!C:C, C{r_idx}) - "
            f"SUMIFS(Distribusi!D:D, Distribusi!A:A, A{r_idx}, Distribusi!B:B, B{r_idx}, Distribusi!C:C, C{r_idx})"
        )
        formula_nilai = f"=D{r_idx}*E{r_idx}"
        row_vals = [
            item["dana"], item["barang"], item["kat"],
            formula_stok, item["harga_satuan"], formula_nilai
        ]
        ws_stok.append(row_vals)
        ws_stok.row_dimensions[r_idx].height = 20

    max_row_stok = len(stok_awal_data) + 1
    for c_idx in range(1, 7):
        h = ws_stok.cell(row=1, column=c_idx)
        h.font = font_header
        h.fill = fill_header
        h.alignment = align_header
        h.border = border_cell

    for r_idx in range(2, max_row_stok + 1):
        is_even = (r_idx % 2 == 0)
        for c_idx in range(1, 7):
            c = ws_stok.cell(row=r_idx, column=c_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            if c_idx in [1, 3]:
                c.alignment = align_center
            elif c_idx in [4, 5, 6]:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold

    dv_dana.add(f"A2:A{max_row_stok}")
    dv_kat.add(f"C2:C{max_row_stok}")

    tab_stok = Table(displayName="tblStok", ref=f"A1:F{max_row_stok}")
    tab_stok.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_stok.add_table(tab_stok)

    tot_row_s = max_row_stok + 1
    ws_stok.cell(row=tot_row_s, column=1, value="TOTAL KESELURUHAN STOK").font = font_bold
    ws_stok.cell(row=tot_row_s, column=1).alignment = align_center
    ws_stok.cell(row=tot_row_s, column=4, value=f"=SUM(D2:D{max_row_stok})").number_format = '#,##0'
    ws_stok.cell(row=tot_row_s, column=4).font = font_bold
    ws_stok.cell(row=tot_row_s, column=4).alignment = align_right
    ws_stok.cell(row=tot_row_s, column=6, value=f"=SUM(F2:F{max_row_stok})").number_format = '#,##0'
    ws_stok.cell(row=tot_row_s, column=6).font = font_bold
    ws_stok.cell(row=tot_row_s, column=6).alignment = align_right
    for c_idx in range(1, 7):
        cell_t = ws_stok.cell(row=tot_row_s, column=c_idx)
        cell_t.fill = fill_total
        cell_t.border = border_total

    # -------------------------------------------------------------
    # SHEET 6: Kadaluwarsa (tblKadaluwarsa)
    # -------------------------------------------------------------
    ws_kad = wb_new.create_sheet(title="Kadaluwarsa")
    ws_kad.views.sheetView[0].showGridLines = True
    ws_kad.freeze_panes = "A2"
    ws_kad.add_data_validation(dv_dana)
    ws_kad.add_data_validation(dv_kat)

    headers_kad = [
        "Sumber Pendanaan",
        "Barang",
        "Kategori Barang",
        "Nomor Dokumen",
        "Tanggal Kadaluwarsa",
        "Status Kadaluwarsa",
        "Sisa Hari",
        "Harga Satuan (Rp)"
    ]
    ws_kad.append(headers_kad)
    ws_kad.row_dimensions[1].height = 26

    for r_idx, item in enumerate(kadaluwarsa_data, start=2):
        row_vals = [
            item["dana"], item["barang"], item["kat"], item["doc"],
            item["exp"], item["status"], item["sisa"], item["harga_satuan"]
        ]
        ws_kad.append(row_vals)
        ws_kad.row_dimensions[r_idx].height = 20

    max_row_kad = len(kadaluwarsa_data) + 1
    for c_idx in range(1, 9):
        h = ws_kad.cell(row=1, column=c_idx)
        h.font = font_header
        h.fill = fill_header
        h.alignment = align_header
        h.border = border_cell

    for r_idx in range(2, max_row_kad + 1):
        is_even = (r_idx % 2 == 0)
        for c_idx in range(1, 9):
            c = ws_kad.cell(row=r_idx, column=c_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            if c_idx in [1, 3, 4, 6]:
                c.alignment = align_center
            elif c_idx == 5:
                c.alignment = align_center
                if c.value is not None and hasattr(c.value, "strftime"):
                    c.number_format = 'DD/MM/YYYY'
            elif c_idx in [7, 8]:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold

    dv_dana.add(f"A2:A{max_row_kad}")
    dv_kat.add(f"C2:C{max_row_kad}")

    tab_kad = Table(displayName="tblKadaluwarsa", ref=f"A1:H{max_row_kad}")
    tab_kad.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_kad.add_table(tab_kad)

    # -------------------------------------------------------------
    # SHEET 7: Permohonan Daerah (tblPermohonanDaerah - PENGUATAN TAHUNAN)
    # -------------------------------------------------------------
    ws_perm = wb_new.create_sheet(title="Permohonan Daerah")
    ws_perm.views.sheetView[0].showGridLines = True
    ws_perm.freeze_panes = "A2"
    ws_perm.add_data_validation(dv_thn)
    ws_perm.add_data_validation(dv_kat)
    ws_perm.add_data_validation(dv_tujuan)

    headers_perm = [
        "Nomor Surat Permohonan",
        "Tahun Anggaran",
        "Tanggal Permohonan",
        "Kab/Kota Pemohon",
        "Kategori Barang",
        "Barang",
        "Jumlah Diminta (Unit)",
        "Estimasi Harga Satuan (Rp)",
        "Estimasi Nilai Total (Rp)"
    ]
    ws_perm.append(headers_perm)
    ws_perm.row_dimensions[1].height = 26

    for r_idx, item in enumerate(permohonan_data, start=2):
        row_vals = [
            item["no_surat"],
            item["tahun_anggaran"],
            item["tgl_permohonan"],
            item["kab_kota"],
            item["kategori"],
            item["barang"],
            item["jumlah_diminta"],
            item["harga_satuan"],
            item["estimasi_nilai"]
        ]
        ws_perm.append(row_vals)
        ws_perm.row_dimensions[r_idx].height = 20

    max_row_perm = len(permohonan_data) + 1
    for c_idx in range(1, 10):
        h = ws_perm.cell(row=1, column=c_idx)
        h.font = font_header
        h.fill = fill_header
        h.alignment = align_header
        h.border = border_cell

    for r_idx in range(2, max_row_perm + 1):
        is_even = (r_idx % 2 == 0)
        for c_idx in range(1, 10):
            c = ws_perm.cell(row=r_idx, column=c_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            if c_idx in [1, 2, 5]:
                c.alignment = align_center
            elif c_idx == 3:
                c.alignment = align_center
                if c.value is not None and hasattr(c.value, "strftime"):
                    c.number_format = 'DD/MM/YYYY'
            elif c_idx in [7, 8, 9]:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold

    dv_thn.add(f"B2:B{max_row_perm}")
    dv_tujuan.add(f"D2:D{max_row_perm}")
    dv_kat.add(f"E2:E{max_row_perm}")

    tab_perm = Table(displayName="tblPermohonanDaerah", ref=f"A1:I{max_row_perm}")
    tab_perm.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_perm.add_table(tab_perm)

    # Baris Total Permohonan
    tot_row_pm = max_row_perm + 1
    ws_perm.cell(row=tot_row_pm, column=1, value="TOTAL PERMOHONAN PENGUATAN").font = font_bold
    ws_perm.cell(row=tot_row_pm, column=1).alignment = align_center
    ws_perm.cell(row=tot_row_pm, column=7, value=f"=SUM(G2:G{max_row_perm})").number_format = '#,##0'
    ws_perm.cell(row=tot_row_pm, column=7).font = font_bold
    ws_perm.cell(row=tot_row_pm, column=7).alignment = align_right
    ws_perm.cell(row=tot_row_pm, column=9, value=f"=SUM(I2:I{max_row_perm})").number_format = '#,##0'
    ws_perm.cell(row=tot_row_pm, column=9).font = font_bold
    ws_perm.cell(row=tot_row_pm, column=9).alignment = align_right
    for c_idx in range(1, 10):
        cell_t = ws_perm.cell(row=tot_row_pm, column=c_idx)
        cell_t.fill = fill_total
        cell_t.border = border_total

    # Atur urutan sheet: List Data, Penerimaan, Distribusi, Stok Awal, Stok, Kadaluwarsa, Permohonan Daerah, _RefData
    wb_new._sheets = [ws_list, ws_pen, ws_dis, ws_awal, ws_stok, ws_kad, ws_perm, ws_ref]
    wb_new.active = ws_list

    # Atur lebar kolom
    for ws in [ws_list, ws_pen, ws_dis, ws_awal, ws_stok, ws_kad, ws_perm, ws_ref]:
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                val = str(cell.value or '')
                if len(val) > max_len:
                    max_len = len(val)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # Simpan ke berkas Excel asli
    wb_new.save(EXCEL_FILE)
    print(f"\n[SUKSES] Berkas '{EXCEL_FILE}' berhasil diperbarui!")
    print(f"  - Total Transaksi Masuk       : {len(penerimaan_data)} Baris")
    print(f"  - Total Transaksi Keluar      : {len(distribusi_data)} Baris")
    print(f"  - Total Stok Awal             : {len(stok_awal_data)} Baris")
    print(f"  - Total Permohonan Penguatan  : {len(permohonan_data)} Baris")
    print(f"  - Sheet Baru Ditambahkan      : 'Permohonan Daerah'")
    print(f"  - Kolom Finansial Ditambahkan : 'Harga Satuan (Rp)' & 'Nilai Total (Rp)'")

if __name__ == "__main__":
    run_update()
