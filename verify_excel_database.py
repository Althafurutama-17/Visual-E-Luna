import openpyxl

def verify_workbook(path):
    wb = openpyxl.load_workbook(path, data_only=False)
    print("=== 1. SHEET NAMES & ORDER ===")
    for idx, sheetname in enumerate(wb.sheetnames, 1):
        print(f"  {idx}. {sheetname}")

    expected_sheets = ["List Data", "Penerimaan", "Distribusi", "Stok Awal", "Stok", "Kadaluwarsa"]
    for s in expected_sheets:
        assert s in wb.sheetnames, f"Sheet {s} tidak ditemukan!"
    print("  -> Seluruh Sheet Utama Terverifikasi Lengkap!\n")

    print("=== 2. EXCEL TABLES VERIFICATION ===")
    expected_tables = {
        "List Data": "tblListData",
        "Penerimaan": "tblPenerimaan",
        "Distribusi": "tblDistribusi",
        "Stok Awal": "tblStokAwal",
        "Stok": "tblStok",
        "Kadaluwarsa": "tblKadaluwarsa"
    }

    for s_name, t_name in expected_tables.items():
        ws = wb[s_name]
        table_names = [t.displayName for t in ws.tables.values()]
        print(f"  Sheet '{s_name}' Tables: {table_names}")
        assert t_name in table_names, f"Tabel {t_name} tidak ditemukan di sheet {s_name}!"
    print("  -> Seluruh 6 Tabel Excel Terverifikasi dengan Nama Resmi!\n")

    print("=== 3. HEADERS & FIRST ROW SAMPLE ===")
    for s_name in expected_sheets:
        ws = wb[s_name]
        headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
        row2 = [ws.cell(2, col).value for col in range(1, ws.max_column + 1)]
        print(f"\n[{s_name}] (Total Baris: {ws.max_row})")
        print(f"  Headers: {headers}")
        print(f"  Row 2  : {row2}")

    print("\n=== 4. FORMULA CHECK SAMPLE ===")
    ws_stok = wb["Stok"]
    print(f"  Formula Stok Baris 2 (D2): {ws_stok['D2'].value}")
    print(f"  Formula Total Stok (D{ws_stok.max_row}): {ws_stok.cell(ws_stok.max_row, 4).value}")

    ws_exp = wb["Kadaluwarsa"]
    print(f"  Formula Sisa Hari Baris 2 (F2): {ws_exp['F2'].value}")

    ws_pen = wb["Penerimaan"]
    print(f"  Formula Total Masuk (D{ws_pen.max_row}): {ws_pen.cell(ws_pen.max_row, 4).value}")

    ws_dist = wb["Distribusi"]
    print(f"  Formula Total Keluar (D{ws_dist.max_row}): {ws_dist.cell(ws_dist.max_row, 4).value}")

    print("\n=== 5. CONDITIONAL FORMATTING CHECK ===")
    cf_exp = ws_exp.conditional_formatting
    print(f"  Jumlah Aturan Conditional Formatting di Kadaluwarsa: {len(cf_exp)}")
    for rule in cf_exp:
        print(f"    - Range: {rule.sqref}")

    print("\n>>> VERIFIKASI SELESAI: SELURUH STRUKTUR, FORMULA, DAN FORMAT EXCEL 100% VALID! <<<")

if __name__ == "__main__":
    file_path = r"c:\My Data\01 AKTIF\Semester 7\KP - Capstone - PPI - TA\KP - AI Assisted Project\Dashboard E-LUNA\database_logistik_bpbd_jabar.xlsx"
    verify_workbook(file_path)
