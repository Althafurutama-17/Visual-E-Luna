import datetime
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

def generate_expanded_database(output_path):
    # Set random seed agar data deterministik dan konsisten setiap kali dieksekusi
    random.seed(42)

    wb = Workbook()

    # -------------------------------------------------------------
    # STYLING RESMI E-LUNA & BPBD JAWA BARAT
    # -------------------------------------------------------------
    FONT_FAMILY = "Segoe UI"
    font_header = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="0E63AA", end_color="0E63AA", fill_type="solid") # Biru E-Luna
    
    font_sub_header = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
    fill_sub_header = PatternFill(start_color="11538C", end_color="11538C", fill_type="solid")
    
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
    # 1. SHEET _RefData (Master Dropdown References)
    # -------------------------------------------------------------
    ws_ref = wb.active
    ws_ref.title = "_RefData"
    ws_ref.views.sheetView[0].showGridLines = True
    
    sumber_dana = ["APBN", "APBD", "CSR"]
    ws_ref['A1'] = "Sumber Pendanaan"
    for r_idx, val in enumerate(sumber_dana, start=2):
        ws_ref[f'A{r_idx}'] = val
        
    kategori_barang = ["Sandang", "Pangan", "Papan", "Logistik Lainnya"]
    ws_ref['B1'] = "Kategori Barang"
    for r_idx, val in enumerate(kategori_barang, start=2):
        ws_ref[f'B{r_idx}'] = val
        
    kab_kota_jabar = [
        "-",
        "Kabupaten Bandung",
        "Kabupaten Bandung Barat",
        "Kabupaten Bekasi",
        "Kabupaten Bogor",
        "Kabupaten Ciamis",
        "Kabupaten Cianjur",
        "Kabupaten Cirebon",
        "Kabupaten Garut",
        "Kabupaten Indramayu",
        "Kabupaten Karawang",
        "Kabupaten Kuningan",
        "Kabupaten Majalengka",
        "Kabupaten Pangandaran",
        "Kabupaten Purwakarta",
        "Kabupaten Subang",
        "Kabupaten Sukabumi",
        "Kabupaten Sumedang",
        "Kabupaten Tasikmalaya",
        "Kota Bandung",
        "Kota Banjar",
        "Kota Bekasi",
        "Kota Bogor",
        "Kota Cimahi",
        "Kota Cirebon",
        "Kota Depok",
        "Kota Sukabumi",
        "Kota Tasikmalaya"
    ]
    ws_ref['C1'] = "Kab/Kota Tujuan"
    for r_idx, val in enumerate(kab_kota_jabar, start=2):
        ws_ref[f'C{r_idx}'] = val

    for col in ['A', 'B', 'C']:
        cell = ws_ref[f'{col}1']
        cell.font = font_sub_header
        cell.fill = fill_sub_header
        cell.alignment = align_center

    # -------------------------------------------------------------
    # 2. DEFINISI MASTER BARANG UNIK SEDERHANA (3-6 per Kategori)
    # -------------------------------------------------------------
    master_items = {
        "Sandang": [
            {"nama": "Selimut", "qty_min": 200, "qty_max": 2000, "perishable": False},
            {"nama": "Matras", "qty_min": 150, "qty_max": 1500, "perishable": False},
            {"nama": "Sarung", "qty_min": 100, "qty_max": 1000, "perishable": False},
            {"nama": "Pakaian Dewasa", "qty_min": 200, "qty_max": 2500, "perishable": False},
            {"nama": "Pakaian Anak", "qty_min": 150, "qty_max": 1800, "perishable": False}
        ],
        "Pangan": [
            {"nama": "Beras", "qty_min": 1000, "qty_max": 15000, "perishable": True, "shelf_months": 12},
            {"nama": "Makanan Siap Saji", "qty_min": 500, "qty_max": 5000, "perishable": True, "shelf_months": 8},
            {"nama": "Susu Formula", "qty_min": 100, "qty_max": 1200, "perishable": True, "shelf_months": 6},
            {"nama": "Biskuit", "qty_min": 300, "qty_max": 3000, "perishable": True, "shelf_months": 9},
            {"nama": "Lauk Pauk", "qty_min": 400, "qty_max": 4500, "perishable": True, "shelf_months": 10},
            {"nama": "Air Mineral", "qty_min": 500, "qty_max": 6000, "perishable": True, "shelf_months": 18}
        ],
        "Papan": [
            {"nama": "Tenda Pengungsi", "qty_min": 10, "qty_max": 150, "perishable": False},
            {"nama": "Tenda Keluarga", "qty_min": 20, "qty_max": 250, "perishable": False},
            {"nama": "Terpal", "qty_min": 100, "qty_max": 1200, "perishable": False},
            {"nama": "Seng", "qty_min": 200, "qty_max": 2000, "perishable": False}
        ],
        "Logistik Lainnya": [
            {"nama": "Perahu Karet", "qty_min": 2, "qty_max": 25, "perishable": False},
            {"nama": "Genset", "qty_min": 5, "qty_max": 40, "perishable": False},
            {"nama": "Chainsaw", "qty_min": 5, "qty_max": 50, "perishable": False},
            {"nama": "Family Kit", "qty_min": 100, "qty_max": 1500, "perishable": False},
            {"nama": "Sanitasi Kit", "qty_min": 100, "qty_max": 1200, "perishable": True, "shelf_months": 12}
        ]
    }

    # Flat list semua komoditas
    all_commodities = []
    for kat, items in master_items.items():
        for it in items:
            all_commodities.append({"kategori": kat, **it})

    # -------------------------------------------------------------
    # 3. GENERATE 200 TRANSAKSI PENERIMAAN (MASUK)
    # Rentang: 01 Januari 2025 s.d. 30 September 2026 (638 hari)
    # -------------------------------------------------------------
    start_date = datetime.date(2025, 1, 1)
    end_date = datetime.date(2026, 9, 30)
    total_days = (end_date - start_date).days

    penerimaan_entries = []
    bast_counter_2025 = 1
    bast_counter_2026 = 1

    # Kita buat 200 entri masuk dengan tanggal yang tersebar merata
    # Buat 200 tanggal terdistribusi
    random_day_offsets_in = sorted([random.randint(0, total_days) for _ in range(200)])

    # Pengaturan proporsi kadaluwarsa:
    # Sekitar 5% - 10% dari barang perishable akan diset tanggal kadaluwarsanya:
    # - Sebagian expired (sebelum Okt 2026)
    # - Sebagian warning (antara Okt s.d. Nov 2026)
    # - Sebagian aman (2027+)
    
    for i, day_offset in enumerate(random_day_offsets_in):
        tgl_terima = start_date + datetime.timedelta(days=day_offset)
        dana = random.choice(["APBN", "APBN", "APBD", "APBD", "CSR"]) # APBN & APBD lebih dominan
        com = random.choice(all_commodities)
        
        # Kuantitas masuk realistis dibulatkan ke puluhan/ratusan terdekat
        step = 50 if com["qty_max"] >= 1000 else (10 if com["qty_max"] >= 100 else 1)
        raw_qty = random.randint(com["qty_min"], com["qty_max"])
        qty_in = max(com["qty_min"], (raw_qty // step) * step)

        # Nomor BAST
        thn = tgl_terima.year
        if thn == 2025:
            doc_no = f"BAST-2025-IN-{bast_counter_2025:04d}"
            bast_counter_2025 += 1
        else:
            doc_no = f"BAST-2026-IN-{bast_counter_2026:04d}"
            bast_counter_2026 += 1

        # Tanggal Kadaluwarsa
        tgl_exp = None
        if com["perishable"]:
            # Kita berikan tanggal kadaluwarsa pada barang pangan/sanitasi
            # Sekitar 50% dari barang pangan memiliki tanggal kadaluwarsa tercatat di BAST
            if random.random() < 0.65:
                # Variasikan tanggal kadaluwarsa:
                # 8% expired (misal exp di pertengahan 2026)
                # 8% warning (misal exp di Okt 2026)
                # 84% aman (exp di 2027-2028)
                rand_scenario = random.random()
                if rand_scenario < 0.08:
                    # Expired: antara Mei s.d. Sep 2026 (atau sebelum Okt 2026)
                    tgl_exp = datetime.date(2026, random.randint(5, 9), random.randint(1, 28))
                elif rand_scenario < 0.16:
                    # Warning: Oktober s.d. November 2026
                    tgl_exp = datetime.date(2026, 10, random.randint(5, 30))
                else:
                    # Aman: 2027 s.d. 2028
                    tgl_exp = tgl_terima + datetime.timedelta(days=int(com["shelf_months"] * 30.5) + random.randint(90, 360))

        penerimaan_entries.append({
            "dana": dana,
            "barang": com["nama"],
            "kat": com["kategori"],
            "in": qty_in,
            "out": 0,
            "tujuan": "-",
            "exp": tgl_exp,
            "tgl_in": tgl_terima,
            "tgl_out": None,
            "doc": doc_no
        })

    # -------------------------------------------------------------
    # 4. GENERATE 50 TRANSAKSI DISTRIBUSI (KELUAR)
    # -------------------------------------------------------------
    distribusi_entries = []
    sj_counter_2025 = 1
    sj_counter_2026 = 1

    # 50 tanggal terdistribusi dari Maret 2025 s.d. September 2026 (setelah barang masuk tersedia)
    random_day_offsets_out = sorted([random.randint(60, total_days) for _ in range(50)])
    daerah_list = kab_kota_jabar[1:] # 27 Kab/Kota resmi (tanpa "-")

    for i, day_offset in enumerate(random_day_offsets_out):
        tgl_keluar = start_date + datetime.timedelta(days=day_offset)
        dana = random.choice(["APBN", "APBD", "CSR"])
        com = random.choice(all_commodities)
        
        # Kuantitas pengeluaran (porsi wajar dari penerimaan agar stok tetap positif)
        step = 25 if com["qty_max"] >= 1000 else (5 if com["qty_max"] >= 100 else 1)
        raw_qty = random.randint(int(com["qty_min"] * 0.5), int(com["qty_max"] * 0.6))
        qty_out = max(1, (raw_qty // step) * step)

        tujuan_daerah = random.choice(daerah_list)

        thn = tgl_keluar.year
        if thn == 2025:
            doc_no = f"SJ-2025-OUT-{sj_counter_2025:04d}"
            sj_counter_2025 += 1
        else:
            doc_no = f"SJ-2026-OUT-{sj_counter_2026:04d}"
            sj_counter_2026 += 1

        distribusi_entries.append({
            "dana": dana,
            "barang": com["nama"],
            "kat": com["kategori"],
            "in": 0,
            "out": qty_out,
            "tujuan": tujuan_daerah,
            "exp": None,
            "tgl_in": None,
            "tgl_out": tgl_keluar,
            "doc": doc_no
        })

    # -------------------------------------------------------------
    # 5. GABUNGKAN MENJADI MASTER LIST DATA (250 TRANSAKSI)
    # Urutkan secara kronologis berdasarkan tanggal transaksi
    # -------------------------------------------------------------
    all_transactions = penerimaan_entries + distribusi_entries
    all_transactions.sort(key=lambda x: (x["tgl_in"] or x["tgl_out"]))

    print(f"Total Transaksi Masuk: {len(penerimaan_entries)}")
    print(f"Total Transaksi Keluar: {len(distribusi_entries)}")
    print(f"Total Master Transaksi: {len(all_transactions)}")

    # -------------------------------------------------------------
    # 6. HITUNG STOK AWAL AGAR STOK AKHIR SELALU AMAN & POSITIF
    # Hitung total masuk & keluar per kombinasi (Dana, Barang, Kat)
    # -------------------------------------------------------------
    mutasi_counter = {}
    for tx in all_transactions:
        key = (tx["dana"], tx["barang"], tx["kat"])
        if key not in mutasi_counter:
            mutasi_counter[key] = {"in": 0, "out": 0}
        mutasi_counter[key]["in"] += tx["in"]
        mutasi_counter[key]["out"] += tx["out"]

    # Buat master kombinasi (3 Sumber Dana x 20 Barang = 60 entri di Stok Awal & Stok)
    master_stok_baseline = []
    for dana in sumber_dana:
        for com in all_commodities:
            key = (dana, com["nama"], com["kategori"])
            stats = mutasi_counter.get(key, {"in": 0, "out": 0})
            
            # Baseline stok awal: minimal cukup menutupi pengeluaran ditambah penyangga aman (buffer)
            buffer_stock = com["qty_max"] * 2
            stok_awal_val = max(buffer_stock, stats["out"] * 2)
            
            # Bulatkan ke puluhan/ratusan
            if stok_awal_val >= 1000:
                stok_awal_val = (stok_awal_val // 100) * 100
            elif stok_awal_val >= 100:
                stok_awal_val = (stok_awal_val // 10) * 10

            master_stok_baseline.append({
                "dana": dana,
                "barang": com["nama"],
                "kat": com["kategori"],
                "awal": stok_awal_val
            })

    # Urutkan berdasarkan Kategori lalu Barang
    master_stok_baseline.sort(key=lambda x: (x["kat"], x["barang"], x["dana"]))

    # -------------------------------------------------------------
    # 7. BUILD EXCEL WORKBOOK
    # -------------------------------------------------------------
    dv_dana = DataValidation(type="list", formula1="'_RefData'!$A$2:$A$4", allow_blank=True)
    dv_kat = DataValidation(type="list", formula1="'_RefData'!$B$2:$B$5", allow_blank=True)
    dv_tujuan = DataValidation(type="list", formula1="'_RefData'!$C$2:$C$29", allow_blank=True)

    # -------------------------------------------------------------
    # SHEET 1: List Data (tblListData - 250 Entri)
    # -------------------------------------------------------------
    ws_list = wb.create_sheet(title="List Data")
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
        "Nomor Dokumen"
    ]
    ws_list.append(headers_list)
    ws_list.row_dimensions[1].height = 26

    for r_idx, item in enumerate(all_transactions, start=2):
        row_vals = [
            item["dana"],
            item["barang"],
            item["kat"],
            item["in"],
            item["out"],
            item["tujuan"],
            item["exp"],
            item["tgl_in"],
            item["tgl_out"],
            item["doc"]
        ]
        ws_list.append(row_vals)
        ws_list.row_dimensions[r_idx].height = 20

    max_row_list = len(all_transactions) + 1
    for col_idx in range(1, 11):
        cell_header = ws_list.cell(row=1, column=col_idx)
        cell_header.font = font_header
        cell_header.fill = fill_header
        cell_header.alignment = align_header
        cell_header.border = border_cell

    for row_idx in range(2, max_row_list + 1):
        is_even = (row_idx % 2 == 0)
        for col_idx in range(1, 11):
            c = ws_list.cell(row=row_idx, column=col_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            
            if col_idx in [1, 3]:
                c.alignment = align_center
            elif col_idx in [4, 5]:
                c.number_format = '#,##0'
                c.alignment = align_right
                if c.value and c.value > 0:
                    c.font = font_bold
            elif col_idx == 6:
                c.alignment = align_center
            elif col_idx in [7, 8, 9]:
                c.alignment = align_center
                if c.value is not None:
                    c.number_format = 'DD/MM/YYYY'
            elif col_idx == 10:
                c.font = font_code
                c.alignment = align_center

    dv_dana.add(f"A2:A{max_row_list}")
    dv_kat.add(f"C2:C{max_row_list}")
    dv_tujuan.add(f"F2:F{max_row_list}")

    tab_list = Table(displayName="tblListData", ref=f"A1:J{max_row_list}")
    tab_list.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_list.add_table(tab_list)

    # -------------------------------------------------------------
    # SHEET 2: Penerimaan (tblPenerimaan - 200 Entri)
    # -------------------------------------------------------------
    ws_penerimaan = wb.create_sheet(title="Penerimaan")
    ws_penerimaan.views.sheetView[0].showGridLines = True
    ws_penerimaan.freeze_panes = "A2"
    ws_penerimaan.add_data_validation(dv_dana)
    ws_penerimaan.add_data_validation(dv_kat)

    headers_penerimaan = [
        "Sumber Pendanaan",
        "Barang",
        "Kategori Barang",
        "Jumlah Masuk",
        "Tanggal Kadaluwarsa",
        "Tanggal barang diterima",
        "Nomor Dokumen"
    ]
    ws_penerimaan.append(headers_penerimaan)
    ws_penerimaan.row_dimensions[1].height = 26

    # Urutkan penerimaan berdasarkan tanggal barang diterima
    penerimaan_sorted = sorted(penerimaan_entries, key=lambda x: x["tgl_in"])
    for r_idx, item in enumerate(penerimaan_sorted, start=2):
        row_vals = [
            item["dana"],
            item["barang"],
            item["kat"],
            item["in"],
            item["exp"],
            item["tgl_in"],
            item["doc"]
        ]
        ws_penerimaan.append(row_vals)
        ws_penerimaan.row_dimensions[r_idx].height = 20

    max_row_penerimaan = len(penerimaan_sorted) + 1
    for col_idx in range(1, 8):
        cell_header = ws_penerimaan.cell(row=1, column=col_idx)
        cell_header.font = font_header
        cell_header.fill = fill_header
        cell_header.alignment = align_header
        cell_header.border = border_cell

    for row_idx in range(2, max_row_penerimaan + 1):
        is_even = (row_idx % 2 == 0)
        for col_idx in range(1, 8):
            c = ws_penerimaan.cell(row=row_idx, column=col_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            if col_idx in [1, 3]:
                c.alignment = align_center
            elif col_idx == 4:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold
            elif col_idx in [5, 6]:
                c.alignment = align_center
                if c.value is not None:
                    c.number_format = 'DD/MM/YYYY'
            elif col_idx == 7:
                c.font = font_code
                c.alignment = align_center

    dv_dana.add(f"A2:A{max_row_penerimaan}")
    dv_kat.add(f"C2:C{max_row_penerimaan}")

    tab_penerimaan = Table(displayName="tblPenerimaan", ref=f"A1:G{max_row_penerimaan}")
    tab_penerimaan.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_penerimaan.add_table(tab_penerimaan)

    # Baris Total di Penerimaan
    tot_row_p = max_row_penerimaan + 1
    ws_penerimaan.cell(row=tot_row_p, column=1, value="TOTAL PENERIMAAN").font = font_bold
    ws_penerimaan.cell(row=tot_row_p, column=1).alignment = align_center
    ws_penerimaan.cell(row=tot_row_p, column=4, value=f"=SUM(D2:D{max_row_penerimaan})").number_format = '#,##0'
    ws_penerimaan.cell(row=tot_row_p, column=4).font = font_bold
    ws_penerimaan.cell(row=tot_row_p, column=4).alignment = align_right
    for c_idx in range(1, 8):
        cell_t = ws_penerimaan.cell(row=tot_row_p, column=c_idx)
        cell_t.fill = fill_total
        cell_t.border = border_total

    # -------------------------------------------------------------
    # SHEET 3: Distribusi (tblDistribusi - 50 Entri)
    # -------------------------------------------------------------
    ws_distribusi = wb.create_sheet(title="Distribusi")
    ws_distribusi.views.sheetView[0].showGridLines = True
    ws_distribusi.freeze_panes = "A2"
    ws_distribusi.add_data_validation(dv_dana)
    ws_distribusi.add_data_validation(dv_kat)
    ws_distribusi.add_data_validation(dv_tujuan)

    headers_distribusi = [
        "Sumber Pendanaan",
        "Barang",
        "Kategori Barang",
        "Jumlah Keluar",
        "Kab/Kota Tujuan",
        "Tanggal barang dikeluarkan",
        "Nomor Dokumen"
    ]
    ws_distribusi.append(headers_distribusi)
    ws_distribusi.row_dimensions[1].height = 26

    distribusi_sorted = sorted(distribusi_entries, key=lambda x: x["tgl_out"])
    for r_idx, item in enumerate(distribusi_sorted, start=2):
        row_vals = [
            item["dana"],
            item["barang"],
            item["kat"],
            item["out"],
            item["tujuan"],
            item["tgl_out"],
            item["doc"]
        ]
        ws_distribusi.append(row_vals)
        ws_distribusi.row_dimensions[r_idx].height = 20

    max_row_distribusi = len(distribusi_sorted) + 1
    for col_idx in range(1, 8):
        cell_header = ws_distribusi.cell(row=1, column=col_idx)
        cell_header.font = font_header
        cell_header.fill = fill_header
        cell_header.alignment = align_header
        cell_header.border = border_cell

    for row_idx in range(2, max_row_distribusi + 1):
        is_even = (row_idx % 2 == 0)
        for col_idx in range(1, 8):
            c = ws_distribusi.cell(row=row_idx, column=col_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            if col_idx in [1, 3, 5]:
                c.alignment = align_center
            elif col_idx == 4:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold
            elif col_idx == 6:
                c.alignment = align_center
                if c.value is not None:
                    c.number_format = 'DD/MM/YYYY'
            elif col_idx == 7:
                c.font = font_code
                c.alignment = align_center

    dv_dana.add(f"A2:A{max_row_distribusi}")
    dv_kat.add(f"C2:C{max_row_distribusi}")
    dv_tujuan.add(f"E2:E{max_row_distribusi}")

    tab_distribusi = Table(displayName="tblDistribusi", ref=f"A1:G{max_row_distribusi}")
    tab_distribusi.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_distribusi.add_table(tab_distribusi)

    tot_row_d = max_row_distribusi + 1
    ws_distribusi.cell(row=tot_row_d, column=1, value="TOTAL DISTRIBUSI").font = font_bold
    ws_distribusi.cell(row=tot_row_d, column=1).alignment = align_center
    ws_distribusi.cell(row=tot_row_d, column=4, value=f"=SUM(D2:D{max_row_distribusi})").number_format = '#,##0'
    ws_distribusi.cell(row=tot_row_d, column=4).font = font_bold
    ws_distribusi.cell(row=tot_row_d, column=4).alignment = align_right
    for c_idx in range(1, 8):
        cell_t = ws_distribusi.cell(row=tot_row_d, column=c_idx)
        cell_t.fill = fill_total
        cell_t.border = border_total

    # -------------------------------------------------------------
    # SHEET 4: Stok Awal (tblStokAwal)
    # -------------------------------------------------------------
    ws_stok_awal = wb.create_sheet(title="Stok Awal")
    ws_stok_awal.views.sheetView[0].showGridLines = True
    ws_stok_awal.freeze_panes = "A2"
    ws_stok_awal.add_data_validation(dv_dana)
    ws_stok_awal.add_data_validation(dv_kat)

    headers_stok_awal = [
        "Kunci",
        "Sumber Pendanaan",
        "Barang",
        "Kategori Barang",
        "Stok Akhir Periode Sebelumnya"
    ]
    ws_stok_awal.append(headers_stok_awal)
    ws_stok_awal.row_dimensions[1].height = 26

    for r_idx, km in enumerate(master_stok_baseline, start=2):
        formula_kunci = f'=B{r_idx}&"|"&C{r_idx}&"|"&D{r_idx}'
        row_vals = [
            formula_kunci,
            km["dana"],
            km["barang"],
            km["kat"],
            km["awal"]
        ]
        ws_stok_awal.append(row_vals)
        ws_stok_awal.row_dimensions[r_idx].height = 20

    max_row_awal = len(master_stok_baseline) + 1
    for col_idx in range(1, 6):
        cell_header = ws_stok_awal.cell(row=1, column=col_idx)
        cell_header.font = font_header
        cell_header.fill = fill_header
        cell_header.alignment = align_header
        cell_header.border = border_cell

    for row_idx in range(2, max_row_awal + 1):
        is_even = (row_idx % 2 == 0)
        for col_idx in range(1, 6):
            c = ws_stok_awal.cell(row=row_idx, column=col_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            if col_idx == 1:
                c.font = font_code
                c.alignment = align_center
            elif col_idx in [2, 4]:
                c.alignment = align_center
            elif col_idx == 5:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold

    dv_dana.add(f"B2:B{max_row_awal}")
    dv_kat.add(f"D2:D{max_row_awal}")

    tab_awal = Table(displayName="tblStokAwal", ref=f"A1:E{max_row_awal}")
    tab_awal.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_stok_awal.add_table(tab_awal)

    # -------------------------------------------------------------
    # SHEET 5: Stok (tblStok - Formula Dinamis)
    # -------------------------------------------------------------
    ws_stok = wb.create_sheet(title="Stok")
    ws_stok.views.sheetView[0].showGridLines = True
    ws_stok.freeze_panes = "A2"
    ws_stok.add_data_validation(dv_dana)
    ws_stok.add_data_validation(dv_kat)

    headers_stok = [
        "Sumber Pendanaan",
        "Barang",
        "Kategori Barang",
        "Stok"
    ]
    ws_stok.append(headers_stok)
    ws_stok.row_dimensions[1].height = 26

    for r_idx, km in enumerate(master_stok_baseline, start=2):
        formula_stok = (
            f'=IFERROR(VLOOKUP(A{r_idx}&"|"&B{r_idx}&"|"&C{r_idx}, \'Stok Awal\'!$A:$E, 5, FALSE), 0)'
            f'+SUMIFS(Penerimaan!$D:$D, Penerimaan!$A:$A, A{r_idx}, Penerimaan!$B:$B, B{r_idx}, Penerimaan!$C:$C, C{r_idx})'
            f'-SUMIFS(Distribusi!$D:$D, Distribusi!$A:$A, A{r_idx}, Distribusi!$B:$B, B{r_idx}, Distribusi!$C:$C, C{r_idx})'
        )
        row_vals = [
            km["dana"],
            km["barang"],
            km["kat"],
            formula_stok
        ]
        ws_stok.append(row_vals)
        ws_stok.row_dimensions[r_idx].height = 20

    max_row_stok = len(master_stok_baseline) + 1
    for col_idx in range(1, 5):
        cell_header = ws_stok.cell(row=1, column=col_idx)
        cell_header.font = font_header
        cell_header.fill = fill_header
        cell_header.alignment = align_header
        cell_header.border = border_cell

    for row_idx in range(2, max_row_stok + 1):
        is_even = (row_idx % 2 == 0)
        for col_idx in range(1, 5):
            c = ws_stok.cell(row=row_idx, column=col_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            if col_idx in [1, 3]:
                c.alignment = align_center
            elif col_idx == 4:
                c.number_format = '#,##0'
                c.alignment = align_right
                c.font = font_bold

    dv_dana.add(f"A2:A{max_row_stok}")
    dv_kat.add(f"C2:C{max_row_stok}")

    tab_stok = Table(displayName="tblStok", ref=f"A1:D{max_row_stok}")
    tab_stok.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_stok.add_table(tab_stok)

    # Baris Total di Stok
    tot_row_s = max_row_stok + 1
    ws_stok.cell(row=tot_row_s, column=1, value="TOTAL KESELURUHAN STOK").font = font_bold
    ws_stok.cell(row=tot_row_s, column=1).alignment = align_center
    ws_stok.cell(row=tot_row_s, column=4, value=f"=SUM(D2:D{max_row_stok})").number_format = '#,##0'
    ws_stok.cell(row=tot_row_s, column=4).font = font_bold
    ws_stok.cell(row=tot_row_s, column=4).alignment = align_right
    for c_idx in range(1, 5):
        cell_t = ws_stok.cell(row=tot_row_s, column=c_idx)
        cell_t.fill = fill_total
        cell_t.border = border_total

    # -------------------------------------------------------------
    # SHEET 6: Kadaluwarsa (tblKadaluwarsa - 5-10% dari Stok Barang)
    # -------------------------------------------------------------
    ws_kadaluwarsa = wb.create_sheet(title="Kadaluwarsa")
    ws_kadaluwarsa.views.sheetView[0].showGridLines = True
    ws_kadaluwarsa.freeze_panes = "A2"
    ws_kadaluwarsa.add_data_validation(dv_dana)
    ws_kadaluwarsa.add_data_validation(dv_kat)

    headers_kadaluwarsa = [
        "Sumber Pendanaan",
        "Barang",
        "Kategori Barang",
        "Nomor Dokumen",
        "Tanggal Kadaluwarsa",
        "Sisa hari"
    ]
    ws_kadaluwarsa.append(headers_kadaluwarsa)
    ws_kadaluwarsa.row_dimensions[1].height = 26

    # Ambil transaksi masuk yang memiliki tanggal kadaluwarsa
    exp_items = [item for item in penerimaan_sorted if item["exp"] is not None]
    
    # Urutkan berdasarkan tanggal kadaluwarsa dari yang paling dekat/lewat
    exp_items.sort(key=lambda x: x["exp"])

    for r_idx, item in enumerate(exp_items, start=2):
        formula_sisa = f'=E{r_idx}-TODAY()'
        row_vals = [
            item["dana"],
            item["barang"],
            item["kat"],
            item["doc"],
            item["exp"],
            formula_sisa
        ]
        ws_kadaluwarsa.append(row_vals)
        ws_kadaluwarsa.row_dimensions[r_idx].height = 20

    max_row_exp = len(exp_items) + 1
    for col_idx in range(1, 7):
        cell_header = ws_kadaluwarsa.cell(row=1, column=col_idx)
        cell_header.font = font_header
        cell_header.fill = fill_header
        cell_header.alignment = align_header
        cell_header.border = border_cell

    for row_idx in range(2, max_row_exp + 1):
        is_even = (row_idx % 2 == 0)
        for col_idx in range(1, 7):
            c = ws_kadaluwarsa.cell(row=row_idx, column=col_idx)
            c.font = font_data
            c.border = border_cell
            if not is_even:
                c.fill = fill_zebra
            if col_idx in [1, 3]:
                c.alignment = align_center
            elif col_idx == 4:
                c.font = font_code
                c.alignment = align_center
            elif col_idx == 5:
                c.alignment = align_center
                if c.value is not None:
                    c.number_format = 'DD/MM/YYYY'
            elif col_idx == 6:
                c.number_format = '#,##0'
                c.alignment = align_center
                c.font = font_bold

    dv_dana.add(f"A2:A{max_row_exp}")
    dv_kat.add(f"C2:C{max_row_exp}")

    tab_exp = Table(displayName="tblKadaluwarsa", ref=f"A1:F{max_row_exp}")
    tab_exp.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws_kadaluwarsa.add_table(tab_exp)

    # Conditional Formatting Kolom F (Sisa Hari)
    # 1. KADALUWARSA (Sisa Hari <= 0) -> Merah
    fill_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    font_red = Font(name=FONT_FAMILY, size=10, bold=True, color="991B1B")
    rule_expired = CellIsRule(operator='lessThanOrEqual', formula=['0'], stopIfTrue=True, fill=fill_red, font=font_red)
    ws_kadaluwarsa.conditional_formatting.add(f"F2:F{max_row_exp}", rule_expired)

    # 2. WARNING (1 <= Sisa Hari <= 30) -> Kuning
    fill_yellow = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    font_yellow = Font(name=FONT_FAMILY, size=10, bold=True, color="92400E")
    rule_warning = CellIsRule(operator='between', formula=['1', '30'], stopIfTrue=True, fill=fill_yellow, font=font_yellow)
    ws_kadaluwarsa.conditional_formatting.add(f"F2:F{max_row_exp}", rule_warning)

    # 3. AMAN (Sisa Hari > 30) -> Hijau
    fill_green = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    font_green = Font(name=FONT_FAMILY, size=10, bold=True, color="166534")
    rule_safe = CellIsRule(operator='greaterThan', formula=['30'], stopIfTrue=True, fill=fill_green, font=font_green)
    ws_kadaluwarsa.conditional_formatting.add(f"F2:F{max_row_exp}", rule_safe)

    # -------------------------------------------------------------
    # 8. URUTAN WORKSHEET & DEFAULT AKTIF
    # -------------------------------------------------------------
    wb._sheets = [ws_list, ws_penerimaan, ws_distribusi, ws_stok_awal, ws_stok, ws_kadaluwarsa, ws_ref]
    wb.active = ws_list

    # -------------------------------------------------------------
    # 9. AUTO-FIT COLUMN WIDTHS
    # -------------------------------------------------------------
    for ws in [ws_list, ws_penerimaan, ws_distribusi, ws_stok_awal, ws_stok, ws_kadaluwarsa, ws_ref]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or '')
                if val.startswith('='):
                    val = "FormulaValue"
                if len(val) > max_len:
                    max_len = len(val)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    ws_list.column_dimensions['A'].width = 20
    ws_list.column_dimensions['B'].width = 25
    ws_list.column_dimensions['C'].width = 18
    ws_list.column_dimensions['D'].width = 16
    ws_list.column_dimensions['E'].width = 16
    ws_list.column_dimensions['F'].width = 26
    ws_list.column_dimensions['G'].width = 20
    ws_list.column_dimensions['H'].width = 22
    ws_list.column_dimensions['I'].width = 24
    ws_list.column_dimensions['J'].width = 24
    
    ws_penerimaan.column_dimensions['A'].width = 20
    ws_penerimaan.column_dimensions['B'].width = 25
    ws_penerimaan.column_dimensions['C'].width = 18
    ws_penerimaan.column_dimensions['D'].width = 16
    ws_penerimaan.column_dimensions['E'].width = 20
    ws_penerimaan.column_dimensions['F'].width = 22
    ws_penerimaan.column_dimensions['G'].width = 24
    
    ws_distribusi.column_dimensions['A'].width = 20
    ws_distribusi.column_dimensions['B'].width = 25
    ws_distribusi.column_dimensions['C'].width = 18
    ws_distribusi.column_dimensions['D'].width = 16
    ws_distribusi.column_dimensions['E'].width = 26
    ws_distribusi.column_dimensions['F'].width = 24
    ws_distribusi.column_dimensions['G'].width = 24

    ws_stok_awal.column_dimensions['A'].width = 38
    ws_stok_awal.column_dimensions['B'].width = 20
    ws_stok_awal.column_dimensions['C'].width = 25
    ws_stok_awal.column_dimensions['D'].width = 18
    ws_stok_awal.column_dimensions['E'].width = 30

    ws_stok.column_dimensions['A'].width = 20
    ws_stok.column_dimensions['B'].width = 25
    ws_stok.column_dimensions['C'].width = 18
    ws_stok.column_dimensions['D'].width = 22

    ws_kadaluwarsa.column_dimensions['A'].width = 20
    ws_kadaluwarsa.column_dimensions['B'].width = 25
    ws_kadaluwarsa.column_dimensions['C'].width = 18
    ws_kadaluwarsa.column_dimensions['D'].width = 24
    ws_kadaluwarsa.column_dimensions['E'].width = 20
    ws_kadaluwarsa.column_dimensions['F'].width = 16

    wb.save(output_path)
    print(f"\n[SUKSES] File Excel 250 Entri Berhasil Dibuat di:\n  -> {output_path}")

if __name__ == "__main__":
    target = r"c:\My Data\01 AKTIF\Semester 7\KP - Capstone - PPI - TA\KP - AI Assisted Project\Dashboard E-LUNA\database_logistik_bpbd_jabar.xlsx"
    generate_expanded_database(target)
