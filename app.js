/**
 * CONTROLLER & LOGIKA DASHBOARD OPERATOR WEB APPS E-LUNA BARU
 * Badan Penanggulangan Bencana Daerah (BPBD) Provinsi Jawa Barat
 * 
 * INTEGRASI LANGSUNG DATABASE EXCEL (.XLSX):
 * - Membaca & mem-parsing file 'database_logistik_bpbd_jabar.xlsx' langsung di browser via SheetJS
 * - Terintegrasi dengan endpoint server /api/status dan auto-sync
 * - Dukungan unggah file Excel kustom & unduh database .xlsx
 * - Mode Tampilan: 10 Inputan Terakhir (Default) vs Semua Data (250 Transaksi) vs Filter Periode
 */

// ==========================================
// 1. STATE GLOBAL & KONFIGURASI OPERASIONAL
// ==========================================
let currentFilterMode = "default10"; // "default10" | "all" | "filtered"
let activeTransactions = [];
let parameterXHari = 60; // Default ambang batas kadaluwarsa (X Hari) diset ke 60 hari
let activeKadaluwarsaTab = "kritis"; // "kritis" (Kadaluwarsa & Warning) | "expired" | "warning" | "aman" | "semua"

// State Modal Pop-Up Stok Eksisting
let modalKategoriFilter = "all";
let modalSearchKeyword = "";
let expandedBarangSet = new Set();
let modalAggregatedStockData = [];

// Chart.js Instances
let chartPenerimaanInst = null;
let chartDistribusiInst = null;

// Daftar Lengkap 27 Kabupaten / Kota Resmi Jawa Barat
const DAFTAR_27_KAB_KOTA_JABAR = [
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
];

// Inisialisasi saat DOM siap
document.addEventListener("DOMContentLoaded", async () => {
  // 1. Render Lucide Icons
  if (window.lucide) {
    lucide.createIcons();
  }

  // 2. Instant Offline Hydration (Anti-0 ms) dari data.js
  const initialData = window.ELUNA_OPERATOR_DATA || (typeof ELUNA_OPERATOR_DATA !== "undefined" ? ELUNA_OPERATOR_DATA : null);
  if (initialData) {
    window.ELUNA_OPERATOR_DATA = initialData;
    applyOperatorData();
    updateDbStatus("connected", `Excel: database_logistik_bpbd_jabar.xlsx (${initialData.riwayatMutasiSemua.length} Transaksi Luring)`);
  }

  // 3. Muat Data Langsung dari File Excel Biner (Jika berjalan via HTTP/HTTPS)
  await loadExcelDatabaseDirectly();

  console.log("Dashboard E-LUNA Berhasil Diinisialisasi (Murni Visual Dashboard).");
});

// ==========================================
// 2. INTEGRASI LANGSUNG DENGAN FILE EXCEL (.XLSX)
// ==========================================

// Memuat langsung file Excel 'database_logistik_bpbd_jabar.xlsx' dari server
async function loadExcelDatabaseDirectly() {
  const syncIcon = document.getElementById("icon-sync-spin");
  if (syncIcon) syncIcon.classList.add("spin-animation");

  try {
    // Fetch file Excel dengan cache-busting timestamp
    const res = await fetch(`database_logistik_bpbd_jabar.xlsx?t=${Date.now()}`);
    if (!res.ok) {
      throw new Error(`HTTP Error ${res.status}: Gagal mengunduh file Excel dari server.`);
    }

    const arrayBuffer = await res.arrayBuffer();
    processExcelWorkbook(arrayBuffer, "database_logistik_bpbd_jabar.xlsx");
    showToast("Berhasil memuat data langsung dari file Excel!");
  } catch (err) {
    console.warn("Direct Excel fetch tidak dapat diakses (normal jika via file:///):", err);
    // Data lokal sudah aktif melalui Instant Offline Hydration
    const totalTx = (window.ELUNA_OPERATOR_DATA && window.ELUNA_OPERATOR_DATA.riwayatMutasiSemua) ? window.ELUNA_OPERATOR_DATA.riwayatMutasiSemua.length : 250;
    updateDbStatus("fallback", `Excel: database_logistik_bpbd_jabar.xlsx (${totalTx} Transaksi Luring)`);
    
    // Pastikan applyOperatorData terpanggil jika belum
    const localData = window.ELUNA_OPERATOR_DATA || (typeof ELUNA_OPERATOR_DATA !== "undefined" ? ELUNA_OPERATOR_DATA : null);
    if (localData && (!activeTransactions || activeTransactions.length === 0)) {
      window.ELUNA_OPERATOR_DATA = localData;
      applyOperatorData();
    }
  } finally {
    if (syncIcon) syncIcon.classList.remove("spin-animation");
  }
}

// Helper konversi tanggal Excel (mendukung numeric serial, string DD/MM/YYYY, dan Date object tanpa UTC offset)
function parseExcelDateValue(val) {
  if (!val) return { tglStr: "2026-09-15", blnStr: "2026-09", thnStr: "2026" };

  // 1. Jika serial number Excel (contoh: 45658)
  if (typeof val === "number" && typeof XLSX !== "undefined" && XLSX.SSF) {
    const p = XLSX.SSF.parse_date_code(val);
    if (p && p.y) {
      const y = String(p.y);
      const m = String(p.m).padStart(2, "0");
      const d = String(p.d).padStart(2, "0");
      return { tglStr: `${y}-${m}-${d}`, blnStr: `${y}-${m}`, thnStr: y };
    }
  }

  // 2. Jika Date Object (gunakan komponen tanggal lokal agar tidak terkena pergeseran UTC)
  if (val instanceof Date && !isNaN(val)) {
    const y = String(val.getFullYear());
    const m = String(val.getMonth() + 1).padStart(2, "0");
    const d = String(val.getDate()).padStart(2, "0");
    return { tglStr: `${y}-${m}-${d}`, blnStr: `${y}-${m}`, thnStr: y };
  }

  // 3. Jika format string DD/MM/YYYY
  const s = String(val).trim();
  if (s.includes("/")) {
    const p = s.split("/");
    if (p.length === 3) {
      const d = p[0].padStart(2, "0");
      const m = p[1].padStart(2, "0");
      const y = p[2].length === 2 ? `20${p[2]}` : p[2];
      return { tglStr: `${y}-${m}-${d}`, blnStr: `${y}-${m}`, thnStr: y };
    }
  }

  // 4. Jika format string YYYY-MM-DD
  if (s.includes("-")) {
    const p = s.split("-");
    if (p.length === 3) {
      if (p[0].length === 4) {
        return { tglStr: s.slice(0, 10), blnStr: s.slice(0, 7), thnStr: p[0] };
      } else {
        const d = p[0].padStart(2, "0");
        const m = p[1].padStart(2, "0");
        const y = p[2].length === 2 ? `20${p[2]}` : p[2];
        return { tglStr: `${y}-${m}-${d}`, blnStr: `${y}-${m}`, thnStr: y };
      }
    }
  }

  return { tglStr: "2026-09-15", blnStr: "2026-09", thnStr: "2026" };
}

// Memproses ArrayBuffer workbook Excel menggunakan SheetJS
function processExcelWorkbook(arrayBuffer, sourceFilename) {
  if (typeof XLSX === "undefined") {
    console.error("SheetJS (XLSX) tidak terdeteksi!");
    if (window.ELUNA_OPERATOR_DATA) applyOperatorData();
    return;
  }

  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  console.log("Sheet Excel Terdeteksi:", workbook.SheetNames);

  // 1. Baca Sheet 'List Data'
  const wsList = workbook.Sheets["List Data"] || workbook.Sheets[workbook.SheetNames[0]];
  if (!wsList) {
    throw new Error("Sheet 'List Data' tidak ditemukan di workbook!");
  }

  const rawRows = XLSX.utils.sheet_to_json(wsList);
  console.log(`Membaca ${rawRows.length} baris dari sheet List Data...`);

  const transactions = [];
  rawRows.forEach((r, idx) => {
    const dana = (r["Sumber Pendanaan"] || "APBD").toString().trim();
    const barang = (r["Barang"] || "").toString().trim();
    const kat = (r["Kategori Barang"] || "Pangan").toString().trim();
    const masuk = parseInt(r["Jumlah Masuk"] || 0, 10);
    const keluar = parseInt(r["Jumlah Keluar"] || 0, 10);
    const tujuan = (r["Kab/Kota Tujuan"] || "-").toString().trim();
    const doc = (r["Nomor Dokumen"] || `DOC-${idx + 1}`).toString().trim();

    if (!barang) return;

    // Ekstraksi Tanggal dengan parseExcelDateValue
    const tglRaw = (masuk > 0) ? r["Tanggal barang diterima"] : r["Tanggal barang dikeluarkan"];
    const parsedDate = parseExcelDateValue(tglRaw);

    const jenis = (masuk > 0) ? "Masuk" : "Keluar";
    const jumlah = (masuk > 0) ? masuk : keluar;
    const asalTujuan = (masuk > 0) ? `Sumber: ${dana}` : `Tujuan: ${tujuan}`;
    const hargaSatuan = parseInt(r["Harga Satuan (Rp)"] || r["Harga Satuan"] || 0, 10);
    const nilaiTotal = parseInt(r["Nilai Total (Rp)"] || r["Nilai Total"] || (jumlah * hargaSatuan), 10);

    transactions.push({
      id: `TX-${String(idx + 1).padStart(4, "0")}`,
      tanggal: parsedDate.tglStr,
      bulan: parsedDate.blnStr,
      tahun: parsedDate.thnStr,
      noDokumen: doc,
      jenis: jenis,
      kategori: kat,
      uraianBarang: barang,
      jumlah: jumlah,
      satuan: "Unit/Paket",
      sumberDana: dana,
      asalTujuan: asalTujuan,
      kabKotaTujuan: (keluar > 0) ? tujuan : "-",
      hargaSatuan: hargaSatuan,
      nilaiTotal: nilaiTotal
    });
  });

  // Urutkan transaksi dari yang paling baru ke paling lama (Newest first)
  transactions.sort((a, b) => b.tanggal.localeCompare(a.tanggal));

  // 2. Baca Sheet 'Kadaluwarsa'
  const wsExp = workbook.Sheets["Kadaluwarsa"];
  const expItems = [];
  if (wsExp) {
    const rawExp = XLSX.utils.sheet_to_json(wsExp);
    rawExp.forEach((r, idx) => {
      const dana = (r["Sumber Pendanaan"] || "APBD").toString().trim();
      const barang = (r["Barang"] || "").toString().trim();
      const kat = (r["Kategori Barang"] || "Pangan").toString().trim();
      const doc = (r["Nomor Dokumen"] || "").toString().trim();
      const tglExpRaw = r["Tanggal Kadaluwarsa"];
      const harga = parseInt(r["Harga Satuan (Rp)"] || 0, 10);

      if (!barang || !tglExpRaw) return;

      const parsedExp = parseExcelDateValue(tglExpRaw);

      // Cari jumlah unit dari transaksi penerimaan terkait bila ada
      const matchedTx = transactions.find(t => t.noDokumen === doc);
      const qtyStok = matchedTx ? matchedTx.jumlah : 1200;

      expItems.push({
        id: `EXP-${String(idx + 1).padStart(2, "0")}`,
        kodeBarang: `LOG-${kat.slice(0, 3).toUpperCase()}-${String(idx + 1).padStart(3, "0")}`,
        namaBarang: barang,
        kategori: kat,
        jumlahStok: qtyStok,
        satuan: "Unit/Paket",
        tglKadaluwarsaAsli: parsedExp.tglStr,
        lokasiGudang: "Gudang Induk BPBD Jawa Barat",
        sumber: dana,
        noDokumen: doc,
        hargaSatuan: harga,
        nilaiTotal: qtyStok * harga
      });
    });
  }

  // 3. Baca Sheet 'Stok Awal'
  const wsAwal = workbook.Sheets["Stok Awal"];
  const stokPerKat = { Sandang: 0, Pangan: 0, Papan: 0, "Logistik Lainnya": 0 };
  const nilaiAwalPerKat = { Sandang: 0, Pangan: 0, Papan: 0, "Logistik Lainnya": 0 };
  const daftarStokAwalBarang = [];
  if (wsAwal) {
    const rawAwal = XLSX.utils.sheet_to_json(wsAwal);
    rawAwal.forEach(r => {
      const kat = (r["Kategori Barang"] || "").toString().trim();
      const sumber = (r["Sumber Pendanaan"] || "APBD").toString().trim();
      const barang = (r["Barang"] || "").toString().trim();
      const awal = parseInt(r["Stok Akhir Periode Sebelumnya"] || 0, 10);
      const harga = parseInt(r["Harga Satuan (Rp)"] || 0, 10);
      const nilai = parseInt(r["Nilai Total (Rp)"] || (awal * harga), 10);

      if (stokPerKat.hasOwnProperty(kat)) {
        stokPerKat[kat] += awal;
        nilaiAwalPerKat[kat] += nilai;
      }
      if (barang) {
        daftarStokAwalBarang.push({
          sumber: sumber,
          barang: barang,
          kategori: kat,
          stokAwal: awal,
          hargaSatuan: harga,
          nilaiTotal: nilai
        });
      }
    });
  }

  // 4. Baca Sheet 'Permohonan Daerah'
  const wsPerm = workbook.Sheets["Permohonan Daerah"];
  const daftarPermohonanDaerah = [];
  if (wsPerm) {
    const rawPerm = XLSX.utils.sheet_to_json(wsPerm);
    rawPerm.forEach((r, idx) => {
      const noSurat = (r["Nomor Surat Permohonan"] || "").toString().trim();
      const thn = parseInt(r["Tahun Anggaran"] || 2026, 10);
      const tglRaw = r["Tanggal Permohonan"];
      const pemohon = (r["Kab/Kota Pemohon"] || "").toString().trim();
      const kat = (r["Kategori Barang"] || "").toString().trim();
      const barang = (r["Barang"] || "").toString().trim();
      const qty = parseInt(r["Jumlah Diminta (Unit)"] || 0, 10);
      const harga = parseInt(r["Estimasi Harga Satuan (Rp)"] || r["Harga Satuan (Rp)"] || 0, 10);
      const estNilai = parseInt(r["Estimasi Nilai Total (Rp)"] || (qty * harga), 10);

      if (!pemohon || !barang || noSurat.startsWith("TOTAL")) return;

      const parsedTgl = parseExcelDateValue(tglRaw);

      daftarPermohonanDaerah.push({
        id: `REQ-${String(idx + 1).padStart(4, "0")}`,
        noSurat: noSurat,
        tahunAnggaran: thn,
        tanggalPermohonan: parsedTgl.tglStr,
        kabKota: pemohon,
        kategori: kat,
        barang: barang,
        jumlahDiminta: qty,
        hargaSatuan: harga,
        estimasiNilai: estNilai
      });
    });
  }

  // Perbarui state global aplikasi
  const existingRekap = (window.ELUNA_OPERATOR_DATA && window.ELUNA_OPERATOR_DATA.rekapFinansialMakro) || null;
  window.ELUNA_OPERATOR_DATA = {
    metadata: {
      systemName: "E-LUNA LOGISTIK BPBD JAWA BARAT",
      subTitle: "Sistem Informasi & Pemantauan Operasional Logistik Kebencanaan",
      instansi: "BPBD Provinsi Jawa Barat",
      gudangAktif: "Gudang Logistik & Peralatan Provinsi Jawa Barat",
      lokasi: "Jl. Soekarno-Hatta No. 629, Kota Bandung",
      petugas: "Nadhif Althafu Rutama",
      tanggalAcuanSistem: "2026-10-03",
      parameterXDefaultHari: 60,
      sourceDatabase: sourceFilename,
      totalEntries: transactions.length,
      fokusOperasional: "Penguatan Cadangan Logistik Kesiapsiagaan Daerah Tahunan"
    },
    kategoriMaster: [
      { id: "sandang", nama: "Sandang", deskripsi: "Selimut, Matras, Sarung, Pakaian Dewasa, Pakaian Anak", color: "#8B5CF6", bgLight: "#F5F3FF", stokAwal: stokPerKat["Sandang"] || 5340, nilaiStokAwal: nilaiAwalPerKat["Sandang"], satuan: "Unit/Paket" },
      { id: "pangan", nama: "Pangan", deskripsi: "Beras, Makanan Siap Saji, Susu Formula, Biskuit, Lauk Pauk, Air Mineral", color: "#3B82F6", bgLight: "#EFF6FF", stokAwal: stokPerKat["Pangan"] || 21000, nilaiStokAwal: nilaiAwalPerKat["Pangan"], satuan: "Unit/Paket" },
      { id: "papan", nama: "Papan", deskripsi: "Tenda Pengungsi, Tenda Keluarga, Terpal, Seng", color: "#F97316", bgLight: "#FFF7ED", stokAwal: stokPerKat["Papan"] || 2160, nilaiStokAwal: nilaiAwalPerKat["Papan"], satuan: "Unit" },
      { id: "lainnya", nama: "Logistik Lainnya", deskripsi: "Perahu Karet, Genset, Chainsaw, Family Kit, Sanitasi Kit", color: "#10B981", bgLight: "#ECFDF5", stokAwal: stokPerKat["Logistik Lainnya"] || 1689, nilaiStokAwal: nilaiAwalPerKat["Logistik Lainnya"], satuan: "Unit/Paket" }
    ],
    kabKotaJabar: DAFTAR_27_KAB_KOTA_JABAR,
    rekapFinansialMakro: existingRekap || {
      totalNilaiStokAwal: 3950000000,
      totalNilaiPenerimaan: 34497500000,
      totalNilaiDistribusi: 4630500000,
      koreksiTambahRp: 450000000,
      koreksiKurangRp: 285000000,
      totalNilaiStokAkhir: 33990800000
    },
    daftarStokAwalPerBarang: daftarStokAwalBarang,
    daftarBarangKadaluwarsa: expItems,
    daftarPermohonanDaerah: daftarPermohonanDaerah.length > 0 ? daftarPermohonanDaerah : ((window.ELUNA_OPERATOR_DATA && window.ELUNA_OPERATOR_DATA.daftarPermohonanDaerah) || []),
    riwayatMutasiSemua: transactions
  };

  updateDbStatus("connected", `Excel: ${sourceFilename} (${transactions.length} Transaksi)`);
  applyOperatorData();
}

// Menerapkan data yang aktif ke seluruh UI Dashboard
function applyOperatorData() {
  const totalCount = window.ELUNA_OPERATOR_DATA.riwayatMutasiSemua.length;
  const labelAll = document.getElementById("label-mode-all");
  if (labelAll) {
    labelAll.textContent = `Semua Data (${totalCount})`;
  }

  populateFilterDropdowns();
  setFilterMode(currentFilterMode);
  renderKadaluwarsaModul();

  if (currentDashboardView === "kabid") {
    renderKabidView();
  } else if (currentDashboardView === "kalak") {
    renderKalakView();
  }
}

// Update Status Pill di Navbar Atas
function updateDbStatus(type, label) {
  const pill = document.getElementById("db-status-pill");
  const dot = document.getElementById("db-status-dot");
  const txt = document.getElementById("db-status-label");

  if (!pill || !dot || !txt) return;

  dot.className = "db-status-dot";
  if (type === "loading") {
    dot.classList.add("loading");
  } else if (type === "error" || type === "fallback") {
    dot.classList.add("error");
  }

  txt.textContent = label;
}

// ==========================================
// CONTROLLER SWITCH VIEW MULTI-STAKEHOLDER
// ==========================================
let currentDashboardView = "operasional"; // "operasional" | "kabid" | "kalak"

function switchDashboardView(viewKey) {
  currentDashboardView = viewKey;

  // 1. Update status aktif sidebar item
  const navOperasional = document.getElementById("nav-view-operasional");
  const navKabid = document.getElementById("nav-view-kabid");
  const navKalak = document.getElementById("nav-view-kalak");

  if (navOperasional) navOperasional.classList.toggle("active", viewKey === "operasional");
  if (navKabid) navKabid.classList.toggle("active", viewKey === "kabid");
  if (navKalak) navKalak.classList.toggle("active", viewKey === "kalak");

  // 2. Switch container visibility
  const cOperasional = document.getElementById("container-view-operasional");
  const cKabid = document.getElementById("container-view-kabid");
  const cKalak = document.getElementById("container-view-kalak");

  if (cOperasional) cOperasional.style.display = (viewKey === "operasional") ? "block" : "none";
  if (cKabid) cKabid.style.display = (viewKey === "kabid") ? "block" : "none";
  if (cKalak) cKalak.style.display = (viewKey === "kalak") ? "block" : "none";

  // 3. Update judul halaman
  const pageTitle = document.getElementById("page-main-title");
  if (pageTitle) {
    if (viewKey === "operasional") {
      pageTitle.textContent = "Dashboard Operasional Logistik";
    } else if (viewKey === "kabid") {
      pageTitle.textContent = "Executive View: Kepala Bidang Kedaruratan & Logistik";
    } else if (viewKey === "kalak") {
      pageTitle.textContent = "Strategic View: Kepala Pelaksana BPBD Provinsi Jawa Barat";
    }
  }

  // 4. Trigger View Rendering Khusus
  if (viewKey === "kabid") {
    renderKabidView();
  } else if (viewKey === "kalak") {
    renderKalakView();
  }

  // 5. Refresh Lucide Icons
  if (window.lucide) {
    lucide.createIcons();
  }

  showToast(`Beralih ke tampilan: ${pageTitle ? pageTitle.textContent : viewKey}`);
}

// ==========================================
// 3. CONTROLLER FILTER: DEFAULT 10 TERAKHIR / SEMUA / BULAN & TAHUN
// ==========================================

// Switch Mode Filter: 'default10' | 'all'
function setFilterMode(mode) {
  currentFilterMode = mode;

  const btn10 = document.getElementById("btn-mode-default10");
  const btnAll = document.getElementById("btn-mode-all");
  const indicator = document.getElementById("indicator-filter-mode");
  const selBulan = document.getElementById("select-bulan");
  const selTahun = document.getElementById("select-tahun");

  if (!window.ELUNA_OPERATOR_DATA) return;
  const allTx = ELUNA_OPERATOR_DATA.riwayatMutasiSemua;

  if (mode === "default10") {
    if (btn10) btn10.classList.add("active");
    if (btnAll) btnAll.classList.remove("active");
    if (selBulan) selBulan.value = "all";
    if (selTahun) selTahun.value = "all";

    activeTransactions = allTx.slice(0, 10);
    if (indicator) indicator.textContent = "Mode Tampilan: 10 Inputan Pengisian Terakhir (Default)";
  } else if (mode === "all") {
    if (btn10) btn10.classList.remove("active");
    if (btnAll) btnAll.classList.add("active");
    if (selBulan) selBulan.value = "all";
    if (selTahun) selTahun.value = "all";

    activeTransactions = allTx;
    if (indicator) indicator.textContent = `Mode Tampilan: Seluruh Database Excel (${allTx.length} Transaksi Kumulatif)`;
  }

  updateDashboardCore();
}

// Filter Berdasarkan Pilihan Dropdown Tahun & Bulan
function handleFilterChange() {
  const selTahun = document.getElementById("select-tahun").value;
  const selBulan = document.getElementById("select-bulan").value;
  const btn10 = document.getElementById("btn-mode-default10");
  const btnAll = document.getElementById("btn-mode-all");
  const indicator = document.getElementById("indicator-filter-mode");

  if (selTahun === "all" && selBulan === "all") {
    setFilterMode("all");
    return;
  }

  currentFilterMode = "filtered";
  if (btn10) btn10.classList.remove("active");
  if (btnAll) btnAll.classList.remove("active");

  const allTx = ELUNA_OPERATOR_DATA.riwayatMutasiSemua;
  activeTransactions = allTx.filter(tx => {
    const matchTahun = (selTahun === "all") || (tx.tahun === selTahun);
    const matchBulan = (selBulan === "all") || (tx.bulan && tx.bulan.endsWith(`-${selBulan}`));
    return matchTahun && matchBulan;
  });

  const blnNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  let labelFilter = "";
  if (selBulan !== "all" && selTahun !== "all") {
    labelFilter = `${blnNames[parseInt(selBulan, 10) - 1]} ${selTahun}`;
  } else if (selTahun !== "all") {
    labelFilter = `Tahun ${selTahun}`;
  } else {
    labelFilter = `Bulan ${blnNames[parseInt(selBulan, 10) - 1]}`;
  }

  if (indicator) indicator.textContent = `Mode Tampilan: Filter ${labelFilter} (${activeTransactions.length} Transaksi)`;

  updateDashboardCore();
  showToast(`Data logistik difilter berdasarkan ${labelFilter} (${activeTransactions.length} Transaksi)`);
}

// Reset ke Default Mode (10 Inputan Pengisian Terakhir)
function resetToDefault10() {
  setFilterMode("default10");
}

// Mengisi Dropdown Tahun & Bulan Secara Bersih & Dinamis
function populateFilterDropdowns() {
  const selThn = document.getElementById("select-tahun");
  const selBln = document.getElementById("select-bulan");
  if (!window.ELUNA_OPERATOR_DATA) return;

  const allTx = ELUNA_OPERATOR_DATA.riwayatMutasiSemua;

  // 1. Dropdown Tahun
  if (selThn) {
    const tahunSet = new Set();
    allTx.forEach(tx => { if (tx.tahun) tahunSet.add(tx.tahun); });
    const sortedTahun = Array.from(tahunSet).sort().reverse();
    const currThn = selThn.value;
    selThn.innerHTML = '<option value="all">Semua Tahun</option>' + sortedTahun.map(y => `<option value="${y}">${y}</option>`).join("");
    if (currThn && (currThn === "all" || tahunSet.has(currThn))) {
      selThn.value = currThn;
    }
  }

  // 2. Dropdown Bulan (12 Bulan Standar)
  if (selBln) {
    const blnNames = [
      "Januari", "Februari", "Maret", "April", "Mei", "Juni",
      "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ];
    const currBln = selBln.value;
    selBln.innerHTML = '<option value="all">Pilih Bulan (Semua)</option>' + blnNames.map((name, i) => {
      const code = String(i + 1).padStart(2, "0");
      return `<option value="${code}">${name}</option>`;
    }).join("");
    if (currBln) selBln.value = currBln;
  }
}

// ==========================================
// 3.1 STATE & CONTROLLER FILTER INTERAKTIF BAR CHART (CROSS-FILTERING)
// ==========================================
const activeChartFilter = {
  sumberAnggaran: null, // null | "APBD" | "APBN" | "CSR"
  kabKota: null,        // null | "Kota Bandung" | ...
  jenis: "all"          // "all" | "masuk" | "keluar"
};

// Handler saat Bar pada Grafik diklik
function handleChartBarClick(filterType, val) {
  if (filterType === "sumberAnggaran") {
    if (activeChartFilter.sumberAnggaran === val) {
      activeChartFilter.sumberAnggaran = null; // Toggle off
      showToast(`Filter sumber ${val} dinonaktifkan`);
    } else {
      activeChartFilter.sumberAnggaran = val;
      showToast(`Menyaring transaksi Sumber Anggaran: ${val}`);
    }
  } else if (filterType === "kabKota") {
    const cleanVal = cleanKabKotaName(val);
    if (activeChartFilter.kabKota === cleanVal) {
      activeChartFilter.kabKota = null; // Toggle off
      showToast(`Filter wilayah ${cleanVal} dinonaktifkan`);
    } else {
      activeChartFilter.kabKota = cleanVal;
      showToast(`Menyaring distribusi tujuan: ${cleanVal}`);
    }
  }

  // Update visualisasi bar yang dipilih (terang/redup)
  updateChartHighlighting();
  // Render ulang tabel dinamis
  renderDynamicTable();
}

// Reset semua filter grafik (tombol Reset)
function resetChartFilters() {
  activeChartFilter.sumberAnggaran = null;
  activeChartFilter.kabKota = null;
  updateChartHighlighting();
  renderDynamicTable();
  showToast("Filter grafik telah direset ke seluruh data");
}

// Handler filter jenis mutasi pada Tabel Dinamis (Semua / Masuk Saja / Keluar Saja)
function setDynamicTypeFilter(type) {
  activeChartFilter.jenis = type;

  const tabAll = document.getElementById("tab-dyn-all");
  const tabIn = document.getElementById("tab-dyn-masuk");
  const tabOut = document.getElementById("tab-dyn-keluar");

  if (tabAll) tabAll.classList.toggle("active", type === "all");
  if (tabIn) tabIn.classList.toggle("active", type === "masuk");
  if (tabOut) tabOut.classList.toggle("active", type === "keluar");

  renderDynamicTable();
}

// ==========================================
// 4. PEMBAHARUAN INTI DASHBOARD (METRIK, NERACA & GRAFIK)
// ==========================================
function updateDashboardCore() {
  if (!window.ELUNA_OPERATOR_DATA) return;

  let countBast = 0;
  let totalMasukUnit = 0;
  let countSj = 0;
  let totalKeluarUnit = 0;

  // Akumulasi Mutasi Aktif per 4 Kategori Resmi: Sandang, Pangan, Papan, Logistik Lainnya
  const katMutasi = {
    sandang: { masuk: 0, keluar: 0 },
    pangan: { masuk: 0, keluar: 0 },
    papan: { masuk: 0, keluar: 0 },
    lainnya: { masuk: 0, keluar: 0 }
  };

  // Peta Volume Distribusi ke 27 Kab/Kota Jawa Barat
  const kabKotaVolumeMap = {};
  DAFTAR_27_KAB_KOTA_JABAR.forEach(k => { kabKotaVolumeMap[k] = 0; });

  activeTransactions.forEach(tx => {
    const key = normalizeCategory(tx.kategori);

    if (tx.jenis === "Masuk") {
      countBast++;
      totalMasukUnit += tx.jumlah;
      if (katMutasi[key]) katMutasi[key].masuk += tx.jumlah;
    } else if (tx.jenis === "Keluar") {
      countSj++;
      totalKeluarUnit += tx.jumlah;
      if (katMutasi[key]) katMutasi[key].keluar += tx.jumlah;

      const cleanTujuan = cleanKabKotaName(tx.kabKotaTujuan || tx.asalTujuan);
      if (kabKotaVolumeMap.hasOwnProperty(cleanTujuan)) {
        kabKotaVolumeMap[cleanTujuan] += tx.jumlah;
      }
    }
  });

  // 1. Update 4 Kartu Metrik Operasional Utama
  const elMasuk = document.getElementById("kpi-total-masuk");
  const elBast = document.getElementById("kpi-bast-count");
  const elKeluar = document.getElementById("kpi-total-keluar");
  const elSj = document.getElementById("kpi-sj-count");
  const elStok = document.getElementById("kpi-total-stok-akhir");
  const elStokAwal = document.getElementById("kpi-stok-awal-note");

  if (elMasuk) elMasuk.textContent = totalMasukUnit.toLocaleString("id-ID");
  if (elBast) elBast.textContent = `${countBast} Dokumen BAST`;

  if (elKeluar) elKeluar.textContent = totalKeluarUnit.toLocaleString("id-ID");
  if (elSj) elSj.textContent = `${countSj} Dokumen Surat Jalan`;

  // 2. Hitung Neraca Stok 4 Kategori: Stok Akhir = Stok Awal (Konstan) + Masuk - Keluar
  let totalAkumulasiStokAkhir = 0;
  let totalStokAwalBasis = 0;
  const neracaData = ELUNA_OPERATOR_DATA.kategoriMaster.map(km => {
    totalStokAwalBasis += (km.stokAwal || 0);
    const m = katMutasi[km.id] || { masuk: 0, keluar: 0 };
    const stokAkhir = km.stokAwal + m.masuk - m.keluar;
    totalAkumulasiStokAkhir += stokAkhir;
    return {
      ...km,
      masuk: m.masuk,
      keluar: m.keluar,
      stokAkhir: stokAkhir
    };
  });

  if (elStok) elStok.textContent = totalAkumulasiStokAkhir.toLocaleString("id-ID");
  if (elStokAwal) {
    elStokAwal.textContent = `Stok Awal: ${totalStokAwalBasis.toLocaleString("id-ID")} Unit`;
    elStokAwal.title = `Rumus Neraca: Stok Awal (${totalStokAwalBasis.toLocaleString("id-ID")}) + Masuk (${totalMasukUnit.toLocaleString("id-ID")}) - Keluar (${totalKeluarUnit.toLocaleString("id-ID")}) = Stok Akhir (${totalAkumulasiStokAkhir.toLocaleString("id-ID")})`;
  }

  // 3. Render Komponen Visualisasi & Tabel Dinamis
  renderVisualisasiCharts(neracaData, kabKotaVolumeMap, totalMasukUnit, totalKeluarUnit);
  renderDynamicTable();

  // 4. Render Stakeholder Multi-Views (Kabid & Kalak)
  renderKabidView(neracaData, totalAkumulasiStokAkhir, countSj, totalKeluarUnit);
  renderKalakView(neracaData, totalAkumulasiStokAkhir);

  if (window.lucide) {
    lucide.createIcons();
  }
}

// State Cache untuk Memperbarui Sorotan Grafik (Highlighting)
let lastChartDataCache = {
  neracaList: [],
  kabKotaMap: {},
  totalMasuk: 0,
  totalKeluar: 0
};

function updateChartHighlighting() {
  if (lastChartDataCache.neracaList.length > 0 || Object.keys(lastChartDataCache.kabKotaMap).length > 0) {
    renderVisualisasiCharts(
      lastChartDataCache.neracaList,
      lastChartDataCache.kabKotaMap,
      lastChartDataCache.totalMasuk,
      lastChartDataCache.totalKeluar
    );
  }
}

// ==========================================
// 6. VISUALISASI STATISTIK DENGAN CROSS-FILTERING (KLIK BAR CHART)
// ==========================================
function renderVisualisasiCharts(neracaList, kabKotaMap, totalMasuk, totalKeluar) {
  lastChartDataCache = { neracaList, kabKotaMap, totalMasuk, totalKeluar };

  const badgeMasuk = document.getElementById("badge-vol-masuk");
  const badgeKeluar = document.getElementById("badge-vol-keluar");

  if (badgeMasuk) badgeMasuk.textContent = `${totalMasuk.toLocaleString("id-ID")} Unit Masuk`;
  if (badgeKeluar) badgeKeluar.textContent = `${totalKeluar.toLocaleString("id-ID")} Unit Keluar`;

  // --- Chart 1: Penerimaan Berdasarkan Sumber Anggaran (APBD, APBN, CSR) ---
  const canvasPenerimaan = document.getElementById("chartPenerimaanSumberAnggaran");
  if (canvasPenerimaan) {
    const sumberAnggaranMasuk = { APBD: 0, APBN: 0, CSR: 0 };
    activeTransactions.forEach(tx => {
      if (tx.jenis === "Masuk") {
        const s = (tx.sumberDana || "").toUpperCase().trim();
        if (s.includes("APBD")) sumberAnggaranMasuk.APBD += tx.jumlah;
        else if (s.includes("APBN")) sumberAnggaranMasuk.APBN += tx.jumlah;
        else if (s.includes("CSR")) sumberAnggaranMasuk.CSR += tx.jumlah;
        else {
          if (sumberAnggaranMasuk.hasOwnProperty(s)) sumberAnggaranMasuk[s] += tx.jumlah;
          else sumberAnggaranMasuk.APBD += tx.jumlah;
        }
      }
    });

    const baseColorsP = { APBD: "#0E63AA", APBN: "#ED7D18", CSR: "#10B981" };
    const dimmedColorsP = { APBD: "rgba(14, 99, 170, 0.22)", APBN: "rgba(237, 125, 24, 0.22)", CSR: "rgba(16, 185, 129, 0.22)" };
    const bgColorsP = ["APBD", "APBN", "CSR"].map(s => {
      if (!activeChartFilter.sumberAnggaran) return baseColorsP[s];
      return (s === activeChartFilter.sumberAnggaran) ? baseColorsP[s] : dimmedColorsP[s];
    });

    const ctxP = canvasPenerimaan.getContext("2d");
    if (chartPenerimaanInst) chartPenerimaanInst.destroy();

    chartPenerimaanInst = new Chart(ctxP, {
      type: "bar",
      data: {
        labels: ["APBD", "APBN", "CSR"],
        datasets: [{
          label: "Volume Penerimaan (Unit)",
          data: [sumberAnggaranMasuk.APBD, sumberAnggaranMasuk.APBN, sumberAnggaranMasuk.CSR],
          backgroundColor: bgColorsP,
          borderRadius: 4,
          barPercentage: 0.45
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        onClick: (evt, elements, chart) => {
          if (elements && elements.length > 0) {
            const idx = elements[0].index;
            const lbl = chart.data.labels[idx];
            handleChartBarClick("sumberAnggaran", lbl);
          }
        },
        onHover: (event, chartElement) => {
          event.native.target.style.cursor = chartElement[0] ? "pointer" : "default";
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (context) => ` Penerimaan ${context.label}: ${context.parsed.y.toLocaleString("id-ID")} Unit (Klik untuk filter)`
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { font: { family: "'Plus Jakarta Sans'", weight: 700, size: 12 } }
          },
          y: {
            grid: { color: "#F1F5F9" },
            ticks: { callback: v => v >= 1000 ? `${v / 1000}k` : v }
          }
        }
      }
    });
  }

  // --- Chart 2: Distribusi ke 27 Kab / Kota Jawa Barat ---
  const canvasDistribusi = document.getElementById("chartDistribusiKabKota");
  if (canvasDistribusi) {
    const sortedDaerah = Object.entries(kabKotaMap)
      .sort((a, b) => b[1] - a[1])
      .filter(d => d[1] > 0 || currentFilterMode === "all")
      .slice(0, 10); // 10 daerah teratas

    const fullNames = sortedDaerah.map(d => d[0]);
    const bgColorsD = fullNames.map(name => {
      if (!activeChartFilter.kabKota) return "#ED7D18";
      return cleanKabKotaName(name) === cleanKabKotaName(activeChartFilter.kabKota)
        ? "#ED7D18"
        : "rgba(237, 125, 24, 0.22)";
    });

    const ctxD = canvasDistribusi.getContext("2d");
    if (chartDistribusiInst) chartDistribusiInst.destroy();

    chartDistribusiInst = new Chart(ctxD, {
      type: "bar",
      data: {
        labels: sortedDaerah.map(d => d[0].replace("Kabupaten ", "Kab. ")),
        datasets: [{
          label: "Volume Distribusi (Unit)",
          data: sortedDaerah.map(d => d[1]),
          backgroundColor: bgColorsD,
          borderRadius: 4,
          barPercentage: 0.55
        }]
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        onClick: (evt, elements, chart) => {
          if (elements && elements.length > 0) {
            const idx = elements[0].index;
            const targetName = fullNames[idx];
            handleChartBarClick("kabKota", cleanKabKotaName(targetName));
          }
        },
        onHover: (event, chartElement) => {
          event.native.target.style.cursor = chartElement[0] ? "pointer" : "default";
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (context) => ` Distribusi: ${context.parsed.x.toLocaleString("id-ID")} Unit (Klik untuk filter)`
            }
          }
        },
        scales: {
          x: {
            grid: { color: "#F1F5F9" },
            ticks: { callback: v => v >= 1000 ? `${v / 1000}k` : v }
          },
          y: {
            grid: { display: false },
            ticks: { font: { family: "'Plus Jakarta Sans'", size: 10, weight: 600 } }
          }
        }
      }
    });
  }
}

// ==========================================
// 6.1 TABEL DINAMIS HASIL FILTER INTERAKTIF
// ==========================================
function renderDynamicTable() {
  if (!window.ELUNA_OPERATOR_DATA) return;
  const tbody = document.getElementById("tbody-dinamis-transaksi");
  const pillContainer = document.getElementById("active-filter-pills");
  const resetBtn = document.getElementById("btn-reset-filter");
  const counterBadge = document.getElementById("counter-tabel-dinamis");

  if (!tbody) return;

  // Filter transaksi aktif
  let filtered = activeTransactions.filter(tx => {
    // 1. Filter Sumber Anggaran (APBD, APBN, CSR)
    if (activeChartFilter.sumberAnggaran) {
      const dana = (tx.sumberDana || "").toUpperCase();
      if (!dana.includes(activeChartFilter.sumberAnggaran.toUpperCase())) {
        return false;
      }
    }

    // 2. Filter Wilayah Kab/Kota (Distribusi Keluar)
    if (activeChartFilter.kabKota) {
      if (tx.jenis !== "Keluar") return false;
      const cleanTujuan = cleanKabKotaName(tx.kabKotaTujuan || tx.asalTujuan);
      if (cleanTujuan !== activeChartFilter.kabKota) {
        return false;
      }
    }

    // 3. Filter Jenis Mutasi (all / masuk / keluar)
    if (activeChartFilter.jenis === "masuk" && tx.jenis !== "Masuk") return false;
    if (activeChartFilter.jenis === "keluar" && tx.jenis !== "Keluar") return false;

    return true;
  });

  // Update Penghitung Baris
  if (counterBadge) {
    counterBadge.textContent = `${filtered.length} Transaksi`;
  }

  // Update Tag Pills & Tombol Reset
  const hasChartFilter = activeChartFilter.sumberAnggaran || activeChartFilter.kabKota;
  if (resetBtn) {
    resetBtn.style.display = hasChartFilter ? "inline-flex" : "none";
  }

  if (pillContainer) {
    if (!hasChartFilter) {
      pillContainer.innerHTML = '<span class="text-filter-default">Menampilkan seluruh data transaksi (Klik bar chart di atas untuk memfilter data secara spesifik)</span>';
    } else {
      let pillsHtml = '<span style="font-size: 0.775rem; font-weight: 700; color: #475569;">Filter Aktif:</span>';
      if (activeChartFilter.sumberAnggaran) {
        pillsHtml += `
          <span class="filter-pill-tag">
            Sumber: ${activeChartFilter.sumberAnggaran}
            <span class="close-pill" onclick="handleChartBarClick('sumberAnggaran', '${activeChartFilter.sumberAnggaran}')" title="Hapus filter ini">&times;</span>
          </span>
        `;
      }
      if (activeChartFilter.kabKota) {
        pillsHtml += `
          <span class="filter-pill-tag tag-wilayah">
            Tujuan: ${activeChartFilter.kabKota}
            <span class="close-pill" onclick="handleChartBarClick('kabKota', '${activeChartFilter.kabKota}')" title="Hapus filter ini">&times;</span>
          </span>
        `;
      }
      pillContainer.innerHTML = pillsHtml;
    }
  }

  // Render Baris Tabel (9 Kolom)
  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="9" style="text-align: center; padding: 2.5rem 1rem; color: #94A3B8;">
          <i data-lucide="inbox" style="width: 32px; height: 32px; margin-bottom: 0.5rem; opacity: 0.6;"></i>
          <div style="font-weight: 600; font-size: 0.9rem;">Tidak ada data transaksi yang cocok dengan kombinasi filter saat ini.</div>
          <div style="font-size: 0.775rem; margin-top: 0.25rem;">Silakan sesuaikan filter waktu atau klik 'Reset Filter Grafik'.</div>
        </td>
      </tr>
    `;
    if (window.lucide) lucide.createIcons();
    return;
  }

  tbody.innerHTML = filtered.map((tx, idx) => {
    const isMasuk = tx.jenis === "Masuk";
    const badgeJenis = isMasuk
      ? '<span style="font-size: 0.725rem; font-weight: 700; color: #16A34A; background: #DCFCE7; padding: 0.2rem 0.55rem; border-radius: 4px;">Masuk</span>'
      : '<span style="font-size: 0.725rem; font-weight: 700; color: var(--eluna-orange); background: #FFF4EB; padding: 0.2rem 0.55rem; border-radius: 4px;">Keluar</span>';

    const jmlSign = isMasuk ? `+${tx.jumlah.toLocaleString("id-ID")}` : `-${tx.jumlah.toLocaleString("id-ID")}`;
    const jmlColor = isMasuk ? "color: #16A34A;" : "color: var(--eluna-orange);";

    const tujuanText = isMasuk ? "-" : (cleanKabKotaName(tx.kabKotaTujuan || tx.asalTujuan) || "-");

    let badgeDana = "#64748B";
    let bgDana = "#F1F5F9";
    if ((tx.sumberDana || "").includes("APBD")) { badgeDana = "#0E63AA"; bgDana = "#EBF3FA"; }
    else if ((tx.sumberDana || "").includes("APBN")) { badgeDana = "#D97706"; bgDana = "#FEF3C7"; }
    else if ((tx.sumberDana || "").includes("CSR")) { badgeDana = "#16A34A"; bgDana = "#DCFCE7"; }

    return `
      <tr>
        <td class="text-center font-medium">${idx + 1}</td>
        <td style="font-size: 0.8rem; color: #475569;">${formatTgl(tx.tanggal)}</td>
        <td style="font-weight: 700; color: #0E63AA; font-size: 0.8rem;">${tx.noDokumen}</td>
        <td class="text-center">${badgeJenis}</td>
        <td><span style="font-weight: 600; color: #334155;">${tx.kategori}</span></td>
        <td style="font-weight: 600; color: #1E293B;">${tx.uraianBarang}</td>
        <td class="text-right" style="font-weight: 700; ${jmlColor}">${jmlSign} <span style="font-size: 0.7rem; font-weight: 500; color: #64748B;">Unit</span></td>
        <td class="text-center">
          <span style="font-size: 0.725rem; font-weight: 700; color: ${badgeDana}; background: ${bgDana}; padding: 0.2rem 0.55rem; border-radius: 4px;">
            ${tx.sumberDana || "APBD"}
          </span>
        </td>
        <td style="font-size: 0.825rem; color: ${isMasuk ? '#94A3B8' : '#1E293B'}; font-weight: ${isMasuk ? '400' : '600'};">
          ${tujuanText}
        </td>
      </tr>
    `;
  }).join("");

  if (window.lucide) {
    lucide.createIcons();
  }
}

// ==========================================
// 7. MODUL PENGAWASAN KADALUWARSA DINAMIS (PARAMETER X HARI)
// ==========================================
function changeParamX(val) {
  const num = parseInt(val, 10);
  if (isNaN(num) || num < 1) return;
  parameterXHari = num;
  renderKadaluwarsaModul();
  showToast(`Ambang batas kadaluwarsa diset ke ${num} Hari.`);
}

function filterKadaluwarsaTab(tab) {
  activeKadaluwarsaTab = tab;
  renderKadaluwarsaModul();
}

function renderKadaluwarsaModul() {
  if (!window.ELUNA_OPERATOR_DATA) return;

  const refDate = new Date(ELUNA_OPERATOR_DATA.metadata.tanggalAcuanSistem);
  const items = ELUNA_OPERATOR_DATA.daftarBarangKadaluwarsa || [];

  let countExpired = 0;
  let countWarning = 0;

  const analyzed = items.map(item => {
    const expDate = new Date(item.tglKadaluwarsaAsli);
    const diff = Math.ceil((expDate - refDate) / (1000 * 60 * 60 * 24));

    let statusType = "safe";
    let statusText = `Aman (${diff} Hari)`;

    if (diff <= 0) {
      statusType = "expired";
      statusText = `KADALUWARSA (Lewat ${Math.abs(diff)} Hari)`;
      countExpired++;
    } else if (diff <= parameterXHari) {
      statusType = "warning";
      statusText = `WARNING (${diff} Hari Menuju Exp)`;
      countWarning++;
    }

    return {
      ...item,
      sisaHari: diff,
      statusType: statusType,
      statusText: statusText
    };
  });

  // Update Badges & Counters
  const sumExp = document.getElementById("badge-summary-exp");
  const sumWarn = document.getElementById("badge-summary-warn");
  const navKritis = document.getElementById("badge-nav-kritis");
  const kpiKadaluwarsa = document.getElementById("kpi-total-kadaluwarsa");
  const kpiSubExp = document.getElementById("kpi-sub-exp");
  const kpiSubWarn = document.getElementById("kpi-sub-warn");

  if (sumExp) sumExp.textContent = `${countExpired} Kadaluwarsa`;
  if (sumWarn) sumWarn.textContent = `${countWarning} Warning (≤ ${parameterXHari} Hari)`;
  if (navKritis) navKritis.textContent = countExpired + countWarning;

  if (kpiKadaluwarsa) kpiKadaluwarsa.textContent = countExpired + countWarning;
  if (kpiSubExp) kpiSubExp.textContent = `${countExpired} Kadaluwarsa`;
  if (kpiSubWarn) kpiSubWarn.textContent = `${countWarning} Warning (≤ ${parameterXHari} Hari)`;

  // Filter List sesuai Tab Aktif (Default: 'kritis' - Gabungan Kadaluwarsa & Warning)
  let filtered = analyzed;
  if (activeKadaluwarsaTab === "kritis") {
    filtered = analyzed.filter(i => i.statusType === "expired" || i.statusType === "warning");
  } else if (activeKadaluwarsaTab === "expired") {
    filtered = analyzed.filter(i => i.statusType === "expired");
  } else if (activeKadaluwarsaTab === "warning") {
    filtered = analyzed.filter(i => i.statusType === "warning");
  } else if (activeKadaluwarsaTab === "aman") {
    filtered = analyzed.filter(i => i.statusType === "safe");
  }

  // Update Status Aktif Tombol Tab
  const btnKritis = document.getElementById("btn-tab-kritis");
  const btnExp = document.getElementById("btn-tab-expired");
  const btnWarn = document.getElementById("btn-tab-warning");
  const btnAman = document.getElementById("btn-tab-aman");

  if (btnKritis) btnKritis.classList.toggle("active", activeKadaluwarsaTab === "kritis");
  if (btnExp) btnExp.classList.toggle("active", activeKadaluwarsaTab === "expired");
  if (btnWarn) {
    btnWarn.classList.toggle("active", activeKadaluwarsaTab === "warning");
    btnWarn.innerHTML = `🟡 Warning (&le; ${parameterXHari} Hari)`;
  }
  if (btnAman) btnAman.classList.toggle("active", activeKadaluwarsaTab === "aman");

  const tbody = document.getElementById("tbody-kadaluwarsa");
  if (!tbody) return;

  tbody.innerHTML = filtered.map(item => `
    <tr>
      <td style="font-family: monospace; font-weight: 700; color: #64748B;">${item.kodeBarang}</td>
      <td style="font-weight: 600; color: #1E293B;">${item.namaBarang}</td>
      <td style="color: #0E63AA; font-weight: 600;">${item.kategori}</td>
      <td class="text-right" style="font-weight: 700;">${item.jumlahStok.toLocaleString("id-ID")} <span style="font-size: 0.7rem; color: #94A3B8;">${item.satuan}</span></td>
      <td>${formatTgl(item.tglKadaluwarsaAsli)}</td>
      <td class="text-center" style="font-weight: 800; color: ${item.sisaHari <= 0 ? '#DC2626' : item.sisaHari <= parameterXHari ? '#D97706' : '#16A34A'};">
        ${item.sisaHari <= 0 ? `Lewat ${Math.abs(item.sisaHari)} Hari` : `${item.sisaHari} Hari`}
      </td>
      <td>
        <span class="badge-status-exp ${item.statusType}">${item.statusText}</span>
      </td>
      <td style="font-size: 0.75rem; color: #64748B;">${item.lokasiGudang}</td>
    </tr>
  `).join("");
}

// ==========================================
// 8. HELPER & UTILITIES
// ==========================================

// Normalisasi Nama Kategori ke 4 Standar Resmi
function normalizeCategory(kat) {
  const s = (kat || "").toLowerCase();
  if (s.includes("sandang")) return "sandang";
  if (s.includes("pangan")) return "pangan";
  if (s.includes("papan")) return "papan";
  return "lainnya";
}

// Normalisasi Nama Kab/Kota agar cocok dengan DAFTAR_27_KAB_KOTA_JABAR
function cleanKabKotaName(name) {
  if (!name) return "";
  let clean = name.replace("BPBD ", "").trim();
  clean = clean.replace("Kab. ", "Kabupaten ").replace("Kot. ", "Kota ");
  return clean;
}

// Format Tanggal Indonesia (DD Mmm YYYY)
function formatTgl(dateStr) {
  if (!dateStr) return "-";
  const p = dateStr.split("-");
  if (p.length !== 3) return dateStr;
  const bln = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
  return `${p[2]} ${bln[parseInt(p[1], 10) - 1]} ${p[0]}`;
}

// Tampilkan Toast Notifikasi (Maksimal 3 Pop-up, FIFO Auto-Replace)
function showToast(msg) {
  const c = document.getElementById("toast-container");
  if (!c) return;

  // Batasi maksimal 3 pop-up: Hapus toast tertua jika sudah >= 3
  while (c.children.length >= 3) {
    c.firstElementChild.remove();
  }

  const t = document.createElement("div");
  t.className = "toast-item";
  t.textContent = msg;
  c.appendChild(t);

  setTimeout(() => {
    t.style.opacity = "0";
    t.style.transition = "opacity 0.25s ease";
    setTimeout(() => t.remove(), 250);
  }, 2500);
}

function scrollToSection(id) {
  const el = document.getElementById(id);
  if (el) el.scrollIntoView({ behavior: "smooth" });
}

// ==========================================
// PLACEHOLDER HANDLER VIEW KABID & VIEW KALAK
// ==========================================
function renderKabidView(neracaData, totalAkumulasiStokAkhir, countSj, totalKeluarUnit) {
  // Tampilan Kabid saat ini dikosongkan (empty-state placeholder)
}

function renderKalakView(neracaData, totalAkumulasiStokAkhir) {
  // Tampilan Kalak saat ini dikosongkan (empty-state placeholder)
}

// ==========================================
// 9. MODUL POP-UP RINCIAN STOK EKSISTING LOGISTIK
// ==========================================

// Buka Modal Stok Eksisting
function openModalStokEksisting() {
  const modal = document.getElementById("modal-stok-eksisting");
  if (!modal) return;
  modal.classList.add("active");

  // Perbarui Label Periode Aktif pada Header Modal
  const elPeriodLabel = document.getElementById("modal-active-period-label");
  if (elPeriodLabel) {
    if (currentFilterMode === "default10") {
      elPeriodLabel.textContent = "Periode: 10 Inputan Pengisian Terakhir";
    } else if (currentFilterMode === "all") {
      const cnt = (window.ELUNA_OPERATOR_DATA && window.ELUNA_OPERATOR_DATA.riwayatMutasiSemua) ? window.ELUNA_OPERATOR_DATA.riwayatMutasiSemua.length : 250;
      elPeriodLabel.textContent = `Periode: Seluruh Database (${cnt} Transaksi Kumulatif)`;
    } else {
      const ind = document.getElementById("indicator-filter-mode");
      elPeriodLabel.textContent = ind ? ind.textContent.replace("Mode Tampilan: ", "") : "Periode Terfilter";
    }
  }

  buildAggregatedStockData();
  renderModalStokContent();
  if (window.lucide) lucide.createIcons();
}

// Tutup Modal Stok Eksisting
function closeModalStokEksisting() {
  const modal = document.getElementById("modal-stok-eksisting");
  if (!modal) return;
  modal.classList.remove("active");
}

function handleModalBackdropClick(e) {
  if (e.target && e.target.id === "modal-stok-eksisting") {
    closeModalStokEksisting();
  }
}

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    closeModalStokEksisting();
  }
});

// Membangun Agregasi Data Stok per Komoditas & Sumber Dana (Sinkron dengan Periode Terpilih)
function buildAggregatedStockData() {
  if (!window.ELUNA_OPERATOR_DATA) return;

  const daftarStokAwal = window.ELUNA_OPERATOR_DATA.daftarStokAwalPerBarang || [];
  // Sinkronkan arus mutasi masuk dan keluar dengan periode running yang sedang aktif di bilah filter
  const activeTx = (activeTransactions && activeTransactions.length > 0)
    ? activeTransactions
    : ((window.ELUNA_OPERATOR_DATA && window.ELUNA_OPERATOR_DATA.riwayatMutasiSemua) ? window.ELUNA_OPERATOR_DATA.riwayatMutasiSemua : []);
  const daftarKadaluwarsa = window.ELUNA_OPERATOR_DATA.daftarBarangKadaluwarsa || [];
  const refDate = new Date(window.ELUNA_OPERATOR_DATA.metadata.tanggalAcuanSistem);

  const itemMap = new Map();

  // 1. Data Stok Awal per (Barang, Sumber)
  daftarStokAwal.forEach(row => {
    const bName = (row.barang || "").trim();
    if (!bName) return;
    if (!itemMap.has(bName)) {
      itemMap.set(bName, {
        namaBarang: bName,
        kategori: (row.kategori || "Logistik Lainnya").trim(),
        sources: {
          APBD: { awal: 0, masuk: 0, keluar: 0, expired: 0 },
          APBN: { awal: 0, masuk: 0, keluar: 0, expired: 0 },
          CSR:  { awal: 0, masuk: 0, keluar: 0, expired: 0 }
        }
      });
    }
    const item = itemMap.get(bName);
    const s = (row.sumber || "APBD").toUpperCase().trim();
    const sKey = s.includes("APBN") ? "APBN" : s.includes("CSR") ? "CSR" : "APBD";
    item.sources[sKey].awal += (row.stokAwal || 0);
  });

  // 2. Akumulasi Riwayat Transaksi Mutasi Berdasarkan Periode Terpilih
  activeTx.forEach(tx => {
    const bName = (tx.uraianBarang || "").trim();
    if (!bName) return;
    if (!itemMap.has(bName)) {
      itemMap.set(bName, {
        namaBarang: bName,
        kategori: tx.kategori || "Pangan",
        sources: {
          APBD: { awal: 0, masuk: 0, keluar: 0, expired: 0 },
          APBN: { awal: 0, masuk: 0, keluar: 0, expired: 0 },
          CSR:  { awal: 0, masuk: 0, keluar: 0, expired: 0 }
        }
      });
    }
    const item = itemMap.get(bName);
    const s = (tx.sumberDana || "APBD").toUpperCase().trim();
    const sKey = s.includes("APBN") ? "APBN" : s.includes("CSR") ? "CSR" : "APBD";
    if (tx.jenis === "Masuk") {
      item.sources[sKey].masuk += tx.jumlah;
    } else if (tx.jenis === "Keluar") {
      item.sources[sKey].keluar += tx.jumlah;
    }
  });

  // 3. Eksklusi Barang Kadaluwarsa (Lewat Tanggal Kadaluwarsa Asli)
  daftarKadaluwarsa.forEach(exp => {
    const expDate = new Date(exp.tglKadaluwarsaAsli);
    const diff = Math.ceil((expDate - refDate) / (1000 * 60 * 60 * 24));
    if (diff <= 0) {
      const bName = (exp.namaBarang || "").trim();
      if (itemMap.has(bName)) {
        const item = itemMap.get(bName);
        const s = (exp.sumber || "APBD").toUpperCase().trim();
        const sKey = s.includes("APBN") ? "APBN" : s.includes("CSR") ? "CSR" : "APBD";
        item.sources[sKey].expired += (exp.jumlahStok || 0);
      }
    }
  });

  let grandLayak = 0;
  let grandExpired = 0;
  modalAggregatedStockData = [];

  itemMap.forEach(item => {
    let totAwal = 0;
    let totMasuk = 0;
    let totKeluar = 0;
    let totExp = 0;
    let totFisik = 0;
    let totLayak = 0;

    const sourceBreakdowns = ["APBD", "APBN", "CSR"].map(sKey => {
      const sData = item.sources[sKey];
      const fisik = sData.awal + sData.masuk - sData.keluar;
      // Validasi Stok Kadaluwarsa Realistis: Hanya yang nyata-nyata masih sisa di gudang (tidak melebihi stok fisik)
      const actualExpired = Math.max(0, Math.min(sData.expired, fisik));
      const layak = Math.max(0, fisik - actualExpired);

      totAwal += sData.awal;
      totMasuk += sData.masuk;
      totKeluar += sData.keluar;
      totExp += actualExpired;
      totFisik += fisik;
      totLayak += layak;

      return {
        sumber: sKey,
        stokAwal: sData.awal,
        masuk: sData.masuk,
        keluar: sData.keluar,
        stokFisik: fisik,
        expired: actualExpired,
        stokLayak: layak
      };
    });

    grandLayak += totLayak;
    grandExpired += totExp;

    const safeId = "brg-" + item.namaBarang.replace(/[^a-zA-Z0-9]/g, "-").toLowerCase();

    modalAggregatedStockData.push({
      safeId: safeId,
      namaBarang: item.namaBarang,
      kategori: item.kategori,
      stokAwal: totAwal,
      masuk: totMasuk,
      keluar: totKeluar,
      stokFisik: totFisik,
      expired: totExp,
      stokLayak: totLayak,
      sources: sourceBreakdowns
    });
  });

  // Urutkan berdasarkan Kategori, lalu Nama Barang
  modalAggregatedStockData.sort((a, b) => a.kategori.localeCompare(b.kategori) || a.namaBarang.localeCompare(b.namaBarang));

  // Perbarui Badge Total di Header Modal
  const elBadgeLayak = document.getElementById("modal-badge-stok-layak");
  const elBadgeExp = document.getElementById("modal-badge-kadaluwarsa-exclude");

  if (elBadgeLayak) elBadgeLayak.textContent = `${grandLayak.toLocaleString("id-ID")} Unit Layak Pakai`;
  if (elBadgeExp) elBadgeExp.textContent = `${grandExpired.toLocaleString("id-ID")} Unit Kadaluwarsa (Exclude)`;
}

// Render Isi Tabel Modal Stok
function renderModalStokContent() {
  const tbody = document.getElementById("tbody-modal-stok-master");
  if (!tbody) return;

  const kw = modalSearchKeyword.toLowerCase().trim();
  const filtered = modalAggregatedStockData.filter(item => {
    const matchKat = (modalKategoriFilter === "all") || (item.kategori.toLowerCase() === modalKategoriFilter.toLowerCase());
    const matchKw = !kw || item.namaBarang.toLowerCase().includes(kw);
    return matchKat && matchKw;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="9" style="text-align: center; padding: 2.5rem 1rem; color: #94A3B8;">
          <i data-lucide="package-search" style="width: 32px; height: 32px; margin-bottom: 0.5rem; opacity: 0.6;"></i>
          <div style="font-weight: 600;">Tidak ada komoditas logistik yang cocok dengan kriteria pencarian.</div>
        </td>
      </tr>
    `;
    if (window.lucide) lucide.createIcons();
    return;
  }

  tbody.innerHTML = filtered.map((item, idx) => {
    const isExpanded = expandedBarangSet.has(item.safeId);
    const expBadge = item.expired > 0
      ? `<span class="badge-exp-exclude" title="Exclude dari stok layak pakai">-${item.expired.toLocaleString("id-ID")}</span>`
      : `<span style="color: #94A3B8;">-</span>`;

    const subRowsHtml = item.sources.map(s => {
      let colorBg = "#F1F5F9";
      let colorTxt = "#475569";
      if (s.sumber === "APBD") { colorBg = "#EBF3FA"; colorTxt = "#0E63AA"; }
      else if (s.sumber === "APBN") { colorBg = "#FEF3C7"; colorTxt = "#D97706"; }
      else if (s.sumber === "CSR") { colorBg = "#DCFCE7"; colorTxt = "#16A34A"; }

      const sExp = s.expired > 0 ? `<span class="badge-exp-exclude">-${s.expired.toLocaleString("id-ID")}</span>` : `<span style="color: #94A3B8;">-</span>`;

      return `
        <tr>
          <td>
            <span style="font-weight: 700; color: ${colorTxt}; background: ${colorBg}; padding: 0.2rem 0.55rem; border-radius: 4px;">
              ${s.sumber}
            </span>
          </td>
          <td class="text-right">${s.stokAwal.toLocaleString("id-ID")}</td>
          <td class="text-right" style="color: #16A34A; font-weight: 600;">+${s.masuk.toLocaleString("id-ID")}</td>
          <td class="text-right" style="color: var(--eluna-orange); font-weight: 600;">-${s.keluar.toLocaleString("id-ID")}</td>
          <td class="text-right">${sExp}</td>
          <td class="text-right" style="font-weight: 700; color: ${s.stokLayak < 0 ? '#DC2626' : '#0E63AA'};">
            ${s.stokLayak.toLocaleString("id-ID")} <span style="font-size: 0.7rem; font-weight: 500; color: #64748B;">Unit</span>
          </td>
        </tr>
      `;
    }).join("");

    return `
      <!-- Baris Master Komoditas -->
      <tr class="row-item-master ${isExpanded ? 'expanded' : ''}" onclick="toggleDrilldownSumberDana('${item.safeId}')" title="Klik untuk membuka/menutup rincian sumber dana">
        <td class="text-center font-medium">${idx + 1}</td>
        <td>
          <div style="font-weight: 700; color: #1E293B;">
            ${item.namaBarang}
          </div>
        </td>
        <td>
          <span style="font-size: 0.75rem; font-weight: 600; color: #475569; background: #F1F5F9; padding: 0.2rem 0.5rem; border-radius: 4px;">
            ${item.kategori}
          </span>
        </td>
        <td class="text-right" style="color: #475569;">${item.stokAwal.toLocaleString("id-ID")}</td>
        <td class="text-right" style="color: #16A34A; font-weight: 600;">+${item.masuk.toLocaleString("id-ID")}</td>
        <td class="text-right" style="color: var(--eluna-orange); font-weight: 600;">-${item.keluar.toLocaleString("id-ID")}</td>
        <td class="text-right">${expBadge}</td>
        <td class="text-right" style="font-weight: 800; color: #0E63AA; font-size: 0.875rem;">
          ${item.stokLayak.toLocaleString("id-ID")} <span style="font-size: 0.7rem; font-weight: 500; color: #64748B;">Unit</span>
        </td>
        <td class="text-center">
          <button class="btn-arsip" style="height: 28px; padding: 0 0.55rem; font-size: 0.725rem; gap: 0.25rem;">
            <i data-lucide="${isExpanded ? 'chevron-up' : 'chevron-down'}" style="width: 14px; height: 14px;"></i>
            <span>${isExpanded ? 'Tutup' : 'Rincian'}</span>
          </button>
        </td>
      </tr>

      <!-- Baris Detail Drilldown Sumber Dana (Accordion) -->
      <tr class="drilldown-container-row ${isExpanded ? 'active' : ''}" id="drilldown-${item.safeId}">
        <td colspan="9" style="padding: 0; border-top: none;">
          <div class="drilldown-content-box">
            <div style="font-size: 0.75rem; font-weight: 700; color: #334155; margin-bottom: 0.45rem; display: flex; align-items: center; gap: 0.35rem;">
              <i data-lucide="layers" style="width: 13px; height: 13px; color: #0E63AA;"></i>
              <span>Rincian Alokasi Sumber Pendanaan untuk Komoditas: <strong>${item.namaBarang}</strong></span>
            </div>
            <table class="drilldown-subtable">
              <thead>
                <tr>
                  <th style="width: 130px;">Sumber Pendanaan</th>
                  <th style="width: 110px;" class="text-right">Stok Awal</th>
                  <th style="width: 110px;" class="text-right">Masuk (+)</th>
                  <th style="width: 110px;" class="text-right">Keluar (-)</th>
                  <th style="width: 120px;" class="text-right">Kadaluwarsa</th>
                  <th style="width: 150px;" class="text-right">Sisa Stok Layak</th>
                </tr>
              </thead>
              <tbody>
                ${subRowsHtml}
              </tbody>
            </table>
          </div>
        </td>
      </tr>
    `;
  }).join("");

  if (window.lucide) {
    lucide.createIcons();
  }
}

// Toggle Baris Akordeon Drilldown
function toggleDrilldownSumberDana(safeId) {
  if (expandedBarangSet.has(safeId)) {
    expandedBarangSet.delete(safeId);
  } else {
    expandedBarangSet.add(safeId);
  }
  renderModalStokContent();
}

// Filter Kategori di Modal
function filterModalKategori(kat, btnEl) {
  modalKategoriFilter = kat;
  const parent = document.getElementById("modal-category-filter-buttons");
  if (parent) {
    Array.from(parent.children).forEach(b => b.classList.remove("active"));
    if (btnEl) btnEl.classList.add("active");
  }
  renderModalStokContent();
}

// Pencarian Barang di Modal
function handleModalSearch(val) {
  modalSearchKeyword = val;
  renderModalStokContent();
}

// ========================================================
// 8. VIEW KABID (MANAJERIAL KUANTITAS PENGUATAN LOGISTIK)
// ========================================================

let kabidSelectedTahun = "ALL";
let kabidSelectedKategori = "ALL";
let kabidSelectedStatus = "ALL"; // "ALL" | "aman" | "kurang"
let kabidSelectedWilayah = null;
let kabidSearchQuery = "";
let kabidCategoryChartMode = "compare"; // "compare" | "gap"
let chartKabidPemenuhanInstance = null;
let chartKabidGapKategoriInstance = null;

function handleKabidFilterChange() {
  const selThn = document.getElementById("kabid-filter-tahun");
  const selKat = document.getElementById("kabid-filter-kategori");
  if (selThn) kabidSelectedTahun = selThn.value;
  if (selKat) kabidSelectedKategori = selKat.value;
  kabidSelectedWilayah = null;
  renderKabidView();
}

function filterKabidTableStatus(statusKey) {
  kabidSelectedStatus = statusKey;
  const pillIds = {
    "ALL": "btn-status-all",
    "aman": "btn-status-aman",
    "kurang": "btn-status-kurang"
  };
  Object.values(pillIds).forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove("active");
  });
  const activeEl = document.getElementById(pillIds[statusKey]);
  if (activeEl) activeEl.classList.add("active");

  renderKabidTable();
}

function handleKabidSearch(val) {
  kabidSearchQuery = (val || "").toLowerCase().trim();
  renderKabidTable();
}

function resetKabidTableFilter() {
  kabidSelectedStatus = "ALL";
  kabidSelectedWilayah = null;
  kabidSearchQuery = "";
  const sInput = document.getElementById("kabid-search-input");
  if (sInput) sInput.value = "";
  filterKabidTableStatus("ALL");
  renderKabidView();
  showToast("Filter tabel kesenjangan daerah direset");
}

function switchKabidChartKategoriMode(mode) {
  kabidCategoryChartMode = mode;
  const btnCompare = document.getElementById("btn-chart-mode-compare");
  const btnGap = document.getElementById("btn-chart-mode-gap");
  if (btnCompare) btnCompare.classList.toggle("active", mode === "compare");
  if (btnGap) btnGap.classList.toggle("active", mode === "gap");

  const titleEl = document.getElementById("title-chart-kategori");
  const subEl = document.getElementById("subtitle-chart-kategori");
  if (mode === "compare") {
    if (titleEl) titleEl.textContent = "Perbandingan Permintaan vs Realisasi per Kategori";
    if (subEl) subEl.textContent = "Analisis perbandingan kuantitas permohonan vs realisasi penyaluran";
  } else {
    if (titleEl) titleEl.textContent = "Grafik Kesenjangan (Gap) per Kategori Logistik";
    if (subEl) subEl.textContent = "Kuantitas defisit cadangan logistik yang belum terpenuhi";
  }

  const { filteredReq, filteredDist } = computeKabidDataset();
  renderKabidKategoriChart(filteredReq, filteredDist);
}

// Komputasi Dataset Kabid (Kuantitas Permohonan vs Realisasi Distribusi)
function computeKabidDataset() {
  const allReq = (window.ELUNA_OPERATOR_DATA && window.ELUNA_OPERATOR_DATA.daftarPermohonanDaerah) || [];
  const allTx = (window.ELUNA_OPERATOR_DATA && window.ELUNA_OPERATOR_DATA.riwayatMutasiSemua) || [];
  const allDist = allTx.filter(t => t.jenis === "Keluar");
  const masterDaerah = (window.ELUNA_OPERATOR_DATA && window.ELUNA_OPERATOR_DATA.kabKotaJabar) || [];

  // Filter permohonan
  const filteredReq = allReq.filter(r => {
    const matchThn = (kabidSelectedTahun === "ALL") || (r.tahunAnggaran === parseInt(kabidSelectedTahun, 10));
    const matchKat = (kabidSelectedKategori === "ALL") || (r.kategori === kabidSelectedKategori);
    return matchThn && matchKat;
  });

  // Filter distribusi keluar
  const filteredDist = allDist.filter(d => {
    const matchThn = (kabidSelectedTahun === "ALL") || (d.tahun === kabidSelectedTahun);
    const matchKat = (kabidSelectedKategori === "ALL") || (d.kategori === kabidSelectedKategori);
    return matchThn && matchKat;
  });

  // Agregasi per Daerah (27 Kab/Kota Jawa Barat)
  const daerahMap = {};
  masterDaerah.forEach(d => {
    daerahMap[d] = {
      daerah: d,
      minta: 0,
      realisasi: 0,
      gap: 0,
      capaian: 0,
      status: "kurang",
      items: []
    };
  });

  // Rekapitulasi permohonan ke dalam peta daerah
  filteredReq.forEach(r => {
    if (!daerahMap[r.kabKota]) {
      daerahMap[r.kabKota] = {
        daerah: r.kabKota,
        minta: 0,
        realisasi: 0,
        gap: 0,
        capaian: 0,
        status: "kurang",
        items: []
      };
    }
    daerahMap[r.kabKota].minta += r.jumlahDiminta;
    daerahMap[r.kabKota].items.push({
      barang: r.barang,
      kategori: r.kategori,
      diminta: r.jumlahDiminta,
      realisasi: 0,
      gap: r.jumlahDiminta,
      capaian: 0,
      status: "kurang",
      noSurat: r.noSurat
    });
  });

  // Rekapitulasi realisasi distribusi ke dalam peta daerah
  filteredDist.forEach(d => {
    const dTarget = d.kabKotaTujuan;
    if (daerahMap[dTarget]) {
      daerahMap[dTarget].realisasi += d.jumlah;
      const matchedItem = daerahMap[dTarget].items.find(it => it.barang === d.uraianBarang);
      if (matchedItem) {
        matchedItem.realisasi += d.jumlah;
      } else {
        daerahMap[dTarget].items.push({
          barang: d.uraianBarang,
          kategori: d.kategori,
          diminta: 0,
          realisasi: d.jumlah,
          gap: 0,
          capaian: 100,
          status: "aman",
          noSurat: "-"
        });
      }
    }
  });

  // Kalkulasi gap, capaian persentase, dan klasifikasi status (Aman >= 90%, Kurang < 90%)
  Object.values(daerahMap).forEach(d => {
    d.gap = Math.max(0, d.minta - d.realisasi);
    d.capaian = d.minta > 0 ? (d.realisasi / d.minta) * 100 : (d.realisasi > 0 ? 100 : 0);
    d.status = d.capaian >= 90 ? "aman" : "kurang";

    d.items.forEach(it => {
      it.gap = Math.max(0, it.diminta - it.realisasi);
      it.capaian = it.diminta > 0 ? (it.realisasi / it.diminta) * 100 : (it.realisasi > 0 ? 100 : 0);
      it.status = it.capaian >= 90 ? "aman" : "kurang";
    });
  });

  return {
    filteredReq,
    filteredDist,
    daerahMap: Object.values(daerahMap)
  };
}

// Render Utama Tampilan View Kabid
function renderKabidView() {
  const { filteredReq, filteredDist, daerahMap } = computeKabidDataset();

  // 1. Hitung ringkasan akumulatif
  const totalMinta = daerahMap.reduce((acc, d) => acc + d.minta, 0);
  const totalAlokasi = daerahMap.reduce((acc, d) => acc + d.realisasi, 0);
  const totalGap = Math.max(0, totalMinta - totalAlokasi);
  const rasioPemenuhan = totalMinta > 0 ? ((totalAlokasi / totalMinta) * 100).toFixed(1) : "0.0";
  const gapPercent = totalMinta > 0 ? ((totalGap / totalMinta) * 100).toFixed(1) : "0.0";

  const countAman = daerahMap.filter(d => d.status === "aman").length;
  const countKurang = daerahMap.filter(d => d.status === "kurang").length;

  const uniqueSurat = new Set(filteredReq.map(r => r.noSurat)).size;

  // 2. Pasang nilai ke 4 Kartu Metrik Kuantitas
  const elMinta = document.getElementById("kpi-kabid-total-minta");
  if (elMinta) elMinta.textContent = totalMinta.toLocaleString("id-ID");

  const elSurat = document.getElementById("kpi-kabid-surat-count");
  if (elSurat) elSurat.textContent = `${uniqueSurat} Surat Permohonan`;

  const elAlokasi = document.getElementById("kpi-kabid-total-alokasi");
  if (elAlokasi) elAlokasi.textContent = totalAlokasi.toLocaleString("id-ID");

  const elRasio = document.getElementById("kpi-kabid-rasio-pemenuhan");
  if (elRasio) elRasio.textContent = `Capaian: ${rasioPemenuhan}%`;

  const elGap = document.getElementById("kpi-kabid-total-gap");
  if (elGap) elGap.textContent = totalGap.toLocaleString("id-ID");

  const elGapPct = document.getElementById("kpi-kabid-gap-percent");
  if (elGapPct) elGapPct.textContent = `Sisa Kebutuhan: ${gapPercent}%`;

  const elKurang = document.getElementById("kpi-kabid-wilayah-kritis");
  if (elKurang) elKurang.textContent = `${countKurang} Daerah`;

  const elWilSub = document.getElementById("kpi-kabid-wilayah-sub");
  if (elWilSub) elWilSub.textContent = `Prioritas Pemenuhan (${countAman} Aman, ${countKurang} Kurang)`;

  // Perbarui label tombol status
  const bAll = document.getElementById("btn-status-all");
  if (bAll) bAll.textContent = `Semua (${daerahMap.length})`;
  const bAman = document.getElementById("btn-status-aman");
  if (bAman) bAman.textContent = `🟢 Aman (${countAman})`;
  const bKurang = document.getElementById("btn-status-kurang");
  if (bKurang) bKurang.textContent = `🔴 Kurang (${countKurang})`;

  // 3. Render 2 Grafik Chart.js
  renderKabidCharts(daerahMap, filteredReq, filteredDist);

  // 4. Render Tabel 27 Daerah
  renderKabidTable();

  if (window.lucide) {
    lucide.createIcons();
  }
}

// Inisialisasi & Render Grafik Kabid
function renderKabidCharts(daerahList, filteredReq, filteredDist) {
  // --- GRAFIK 1: STATUS PEMENUHAN PER WILAYAH (HORIZONTAL BAR) ---
  const canvas1 = document.getElementById("chartKabidPemenuhanWilayah");
  if (canvas1) {
    if (chartKabidPemenuhanInstance) {
      chartKabidPemenuhanInstance.destroy();
    }

    // Urutkan daerah dari persentase capaian tertinggi ke terendah
    const sortedDaerah = [...daerahList].sort((a, b) => b.capaian - a.capaian);
    const labels = sortedDaerah.map(d => d.daerah.replace("Kabupaten ", "Kab. "));
    const dataValues = sortedDaerah.map(d => parseFloat(d.capaian.toFixed(1)));
    const bgColors = sortedDaerah.map(d => (d.status === "aman" ? "#16A34A" : "#DC2626"));

    const ctx1 = canvas1.getContext("2d");
    chartKabidPemenuhanInstance = new Chart(ctx1, {
      type: "bar",
      data: {
        labels: labels,
        datasets: [{
          label: "Pemenuhan Realisasi (%)",
          data: dataValues,
          backgroundColor: bgColors,
          borderRadius: 4,
          barThickness: 9
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (context) => {
                const idx = context.dataIndex;
                const d = sortedDaerah[idx];
                return [
                  `Capaian: ${context.parsed.x}% (${d.status.toUpperCase()})`,
                  `Permohonan : ${d.minta.toLocaleString('id-ID')} Unit`,
                  `Realisasi  : ${d.realisasi.toLocaleString('id-ID')} Unit`,
                  `Gap        : ${d.gap.toLocaleString('id-ID')} Unit`
                ];
              }
            }
          }
        },
        scales: {
          x: {
            beginAtZero: true,
            max: 100,
            ticks: {
              callback: val => val + "%",
              font: { size: 10 }
            },
            grid: { color: "#F1F5F9" }
          },
          y: {
            ticks: {
              font: { size: 9, weight: "600" },
              autoSkip: false
            },
            grid: { display: false }
          }
        },
        onHover: (evt, elements, chart) => {
          // Pointer kursor aktif saat berada di area teks/label nama daerah di sebelah kiri chartArea
          if (evt.x < chart.chartArea.left && evt.y >= chart.chartArea.top && evt.y <= chart.chartArea.bottom) {
            chart.canvas.style.cursor = "pointer";
          } else {
            chart.canvas.style.cursor = "default";
          }
        },
        onClick: (evt, elements, chart) => {
          // Point 8: Hanya klik pada tulisan/label nama Kab/Kota (area sumbu Y) yang memicu pop-up rincian
          if (evt.x < chart.chartArea.left && evt.y >= chart.chartArea.top && evt.y <= chart.chartArea.bottom) {
            const yScale = chart.scales.y;
            const tickCount = yScale.ticks.length;
            let closestIdx = -1;
            let minDiff = Infinity;
            for (let i = 0; i < tickCount; i++) {
              const pixel = yScale.getPixelForTick(i);
              const diff = Math.abs(pixel - evt.y);
              if (diff < minDiff) {
                minDiff = diff;
                closestIdx = i;
              }
            }
            if (closestIdx >= 0 && closestIdx < sortedDaerah.length) {
              const clickedDaerah = sortedDaerah[closestIdx].daerah;
              openModalKabidRincian(clickedDaerah);
            }
          }
        }
      }
    });
  }

  // --- GRAFIK 2: PERBANDINGAN PER KATEGORI (DENGAN MODE SWITCH TOGGLE) ---
  renderKabidKategoriChart(filteredReq, filteredDist);
}

// Render Grafik Kategori Kabid (Mendukung Switch 'compare' vs 'gap')
function renderKabidKategoriChart(filteredReq, filteredDist) {
  const canvas2 = document.getElementById("chartKabidGapKategori");
  if (!canvas2) return;

  if (chartKabidGapKategoriInstance) {
    chartKabidGapKategoriInstance.destroy();
  }

  const categories = ["Sandang", "Pangan", "Papan", "Logistik Lainnya"];
  const reqPerKat = categories.map(k => {
    return filteredReq.filter(r => r.kategori === k).reduce((acc, r) => acc + r.jumlahDiminta, 0);
  });
  const distPerKat = categories.map(k => {
    return filteredDist.filter(d => d.kategori === k).reduce((acc, d) => acc + d.jumlah, 0);
  });
  const gapPerKat = categories.map((k, i) => Math.max(0, reqPerKat[i] - distPerKat[i]));

  let datasets = [];
  if (kabidCategoryChartMode === "compare") {
    // Mode 1 (Normal): Permintaan vs Realisasi
    datasets = [
      {
        label: "Permohonan (Unit)",
        data: reqPerKat,
        backgroundColor: "#3B82F6",
        borderRadius: 4
      },
      {
        label: "Realisasi Penyaluran (Unit)",
        data: distPerKat,
        backgroundColor: "#10B981",
        borderRadius: 4
      }
    ];
  } else {
    // Mode 2: Pure Grafik Kesenjangan (Gap)
    datasets = [
      {
        label: "Kesenjangan / Gap (Unit)",
        data: gapPerKat,
        backgroundColor: "#F97316",
        borderRadius: 4
      }
    ];
  }

  const ctx2 = canvas2.getContext("2d");
  chartKabidGapKategoriInstance = new Chart(ctx2, {
    type: "bar",
    data: {
      labels: categories,
      datasets: datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "top",
          labels: {
            boxWidth: 12,
            font: { size: 11, weight: "600" }
          }
        },
        tooltip: {
          callbacks: {
            label: context => `${context.dataset.label}: ${context.parsed.y.toLocaleString('id-ID')} Unit`
          }
        }
      },
      scales: {
        x: { grid: { display: false } },
        y: {
          beginAtZero: true,
          ticks: {
            callback: val => val.toLocaleString('id-ID'),
            font: { size: 10 }
          },
          grid: { color: "#F1F5F9" }
        }
      }
    }
  });
}

// Render Tabel Utama Kabid (Tepat 27 Baris Daerah se-Jawa Barat)
function renderKabidTable() {
  const { daerahMap } = computeKabidDataset();
  const tbody = document.getElementById("tbody-kabid-drilldown");
  const counterEl = document.getElementById("kabid-table-counter");
  if (!tbody) return;

  // Urutkan daerah berdasarkan capaian tertinggi ke terendah (7 Aman di atas, 20 Kurang di bawah)
  let sortedDaerah = [...daerahMap].sort((a, b) => b.capaian - a.capaian);

  // Filter pencarian dan status
  let filteredDaerah = sortedDaerah.filter(d => {
    if (kabidSelectedStatus !== "ALL" && d.status !== kabidSelectedStatus) return false;
    if (kabidSearchQuery) {
      return d.daerah.toLowerCase().includes(kabidSearchQuery);
    }
    return true;
  });

  if (filteredDaerah.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="8" style="text-align: center; padding: 2.5rem; color: #64748B;">
          <p style="font-weight: 600;">Tidak ada kabupaten / kota yang sesuai dengan kriteria filter saat ini.</p>
          <button class="status-pill-btn" onclick="resetKabidTableFilter()" style="margin-top: 0.75rem;">Reset Filter</button>
        </td>
      </tr>
    `;
    if (counterEl) counterEl.textContent = "Menampilkan 0 kabupaten / kota";
    if (window.lucide) lucide.createIcons();
    return;
  }

  tbody.innerHTML = filteredDaerah.map((d, idx) => {
    const isAman = d.status === "aman";
    const badgeHtml = isAman
      ? `<span class="badge-status-pemenuhan aman">🟢 Aman (≥90%)</span>`
      : `<span class="badge-status-pemenuhan kurang">🔴 Kurang (&lt;90%)</span>`;

    return `
      <tr>
        <td style="text-align: center; font-weight: 600; color: #64748B;">${idx + 1}</td>
        <td>
          <span class="kabid-daerah-link" onclick="openModalKabidRincian('${d.daerah}')" title="Klik untuk melihat rincian barang">${d.daerah}</span>
        </td>
        <td class="num-cell" style="font-weight: 700;">${d.minta.toLocaleString('id-ID')}</td>
        <td class="num-cell positive" style="font-weight: 700;">${d.realisasi.toLocaleString('id-ID')}</td>
        <td class="num-cell ${d.gap > 0 ? 'negative' : ''}" style="font-weight: 700;">${d.gap.toLocaleString('id-ID')}</td>
        <td style="text-align: center; font-weight: 800; color: ${isAman ? '#16A34A' : '#DC2626'};">
          ${d.capaian.toFixed(1)}%
        </td>
        <td style="text-align: center;">${badgeHtml}</td>
        <td style="text-align: center;">
          <button class="btn-detail-sm" onclick="openModalKabidRincian('${d.daerah}')" title="Rincian per barang">
            <i data-lucide="eye" style="width: 13px; height: 13px;"></i>
            <span>Rincian</span>
          </button>
        </td>
      </tr>
    `;
  }).join("");

  if (counterEl) {
    counterEl.textContent = `Menampilkan ${filteredDaerah.length} dari ${daerahMap.length} Kabupaten / Kota se-Jawa Barat`;
  }

  if (window.lucide) {
    lucide.createIcons();
  }
}

// ========================================================
// MODAL POP-UP RINCIAN PEMENUHAN PENGUATAN DAERAH (VIEW KABID)
// ========================================================

function openModalKabidRincian(daerahName) {
  const { daerahMap } = computeKabidDataset();
  const d = daerahMap.find(item => item.daerah === daerahName || item.daerah.replace("Kabupaten ", "Kab. ") === daerahName.replace("Kabupaten ", "Kab. "));
  if (!d) return;

  const modalEl = document.getElementById("modal-kabid-rincian-daerah");
  if (!modalEl) return;

  // Header Modal
  const titleEl = document.getElementById("modal-kabid-title-daerah");
  if (titleEl) titleEl.textContent = `Rincian Penguatan Logistik: ${d.daerah}`;

  const isAman = d.status === "aman";
  const badgeEl = document.getElementById("modal-kabid-badge-status");
  if (badgeEl) {
    badgeEl.className = `badge-status-pemenuhan ${isAman ? 'aman' : 'kurang'}`;
    badgeEl.textContent = isAman ? `🟢 Aman (${d.capaian.toFixed(1)}%)` : `🔴 Kurang (${d.capaian.toFixed(1)}%)`;
  }

  const firstSurat = d.items.length > 0 && d.items[0].noSurat !== "-" ? d.items[0].noSurat : "PRM/2026/01/" + d.daerah.substring(0, 3).toUpperCase();
  const subInfoEl = document.getElementById("modal-kabid-sub-info");
  if (subInfoEl) {
    subInfoEl.textContent = `Nomor Surat Permohonan: ${firstSurat} • Evaluasi Pagu Cadangan Tahunan`;
  }

  // Ringkasan Badges
  const badgeMinta = document.getElementById("modal-kabid-total-minta-badge");
  if (badgeMinta) badgeMinta.textContent = `Total Permintaan: ${d.minta.toLocaleString('id-ID')} Unit`;

  const badgeSalur = document.getElementById("modal-kabid-total-salur-badge");
  if (badgeSalur) badgeSalur.textContent = `Realisasi Salur: ${d.realisasi.toLocaleString('id-ID')} Unit`;

  const badgeGap = document.getElementById("modal-kabid-total-gap-badge");
  if (badgeGap) badgeGap.textContent = `Kesenjangan (Gap): ${d.gap.toLocaleString('id-ID')} Unit`;

  // Tabel Rincian Barang
  const tbody = document.getElementById("tbody-modal-kabid-rincian");
  if (tbody) {
    const katColors = {
      "Pangan": { bg: "#EFF6FF", text: "#1D4ED8", border: "#BFDBFE" },
      "Sandang": { bg: "#F5F3FF", text: "#6D28D9", border: "#DDD6FE" },
      "Papan": { bg: "#FFF7ED", text: "#C2410C", border: "#FED7AA" },
      "Logistik Lainnya": { bg: "#ECFDF5", text: "#047857", border: "#A7F3D0" }
    };

    if (d.items.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="8" style="text-align: center; padding: 2rem; color: #64748B;">
            Belum ada rincian komoditas barang untuk wilayah ini.
          </td>
        </tr>
      `;
    } else {
      tbody.innerHTML = d.items.map((it, idx) => {
        const kColor = katColors[it.kategori] || { bg: "#F1F5F9", text: "#334155", border: "#CBD5E1" };
        const itAman = it.status === "aman";
        const itBadge = itAman
          ? `<span class="badge-status-pemenuhan aman">Aman (${it.capaian.toFixed(0)}%)</span>`
          : `<span class="badge-status-pemenuhan kurang">Kurang (${it.capaian.toFixed(0)}%)</span>`;

        return `
          <tr>
            <td style="text-align: center; font-weight: 600; color: #64748B;">${idx + 1}</td>
            <td>
              <span class="badge-kat-pill" style="background-color: ${kColor.bg}; color: ${kColor.text}; border: 1px solid ${kColor.border};">
                ${it.kategori}
              </span>
            </td>
            <td style="font-weight: 700; color: #1E293B;">${it.barang}</td>
            <td class="num-cell" style="font-weight: 700;">${it.diminta.toLocaleString('id-ID')}</td>
            <td class="num-cell positive" style="font-weight: 700;">${it.realisasi.toLocaleString('id-ID')}</td>
            <td class="num-cell ${it.gap > 0 ? 'negative' : ''}" style="font-weight: 700;">${it.gap.toLocaleString('id-ID')}</td>
            <td style="text-align: center; font-weight: 800; color: ${itAman ? '#16A34A' : '#D97706'};">
              ${it.capaian.toFixed(1)}%
            </td>
            <td style="text-align: center;">${itBadge}</td>
          </tr>
        `;
      }).join("");
    }
  }

  modalEl.classList.add("active");
  document.body.style.overflow = "hidden";

  if (window.lucide) {
    lucide.createIcons();
  }
}

function closeModalKabidRincian() {
  const modalEl = document.getElementById("modal-kabid-rincian-daerah");
  if (modalEl) {
    modalEl.classList.remove("active");
  }
  document.body.style.overflow = "";
}

function handleModalKabidBackdropClick(event) {
  if (event.target && event.target.id === "modal-kabid-rincian-daerah") {
    closeModalKabidRincian();
  }
}

// ========================================================
// 9. VIEW KALAK (EKSEKUTIF STRATEGIS & FINANSIAL MAKRO RP)
// ========================================================

let kalakSelectedSumberDana = "ALL";
let kalakSelectedTahun = "ALL";
let chartKalakSumberAnggaranInstance = null;
let chartKalakNilaiIndikatifInstance = null;

function handleKalakFilterChange() {
  const selDana = document.getElementById("kalak-filter-sumber-dana");
  const selThn = document.getElementById("kalak-filter-tahun");
  if (selDana) kalakSelectedSumberDana = selDana.value;
  if (selThn) kalakSelectedTahun = selThn.value;
  renderKalakView();
}

function formatRupiahSingkat(num) {
  if (Math.abs(num) >= 1e9) {
    return "Rp " + (num / 1e9).toFixed(2).replace(".", ",") + " Miliar";
  }
  if (Math.abs(num) >= 1e6) {
    return "Rp " + (num / 1e6).toFixed(1).replace(".", ",") + " Juta";
  }
  return "Rp " + Math.round(num).toLocaleString("id-ID");
}

function formatRupiahLengkap(num) {
  return "Rp " + Math.round(num).toLocaleString("id-ID");
}

// Render Utama Tampilan View KALAK
function renderKalakView() {
  const dataStore = window.ELUNA_OPERATOR_DATA;
  if (!dataStore) return;

  const rekap = dataStore.rekapFinansialMakro || {};
  const allTx = dataStore.riwayatMutasiSemua || [];
  const stokAwalList = dataStore.daftarStokAwalPerBarang || [];

  // Filter transaksi berdasarkan pilihan KALAK
  const filteredTx = allTx.filter(t => {
    const matchDana = (kalakSelectedSumberDana === "ALL") || (t.sumberDana === kalakSelectedSumberDana);
    const matchThn = (kalakSelectedTahun === "ALL") || (t.tahun === kalakSelectedTahun);
    return matchDana && matchThn;
  });

  const filteredAwal = stokAwalList.filter(i => {
    return (kalakSelectedSumberDana === "ALL") || (i.sumber === kalakSelectedSumberDana);
  });

  // Kalkulasi nilai finansial dinamis
  const valMasuk = filteredTx.filter(t => t.jenis === "Masuk").reduce((acc, t) => acc + (t.nilaiTotal || 0), 0);
  const valKeluar = filteredTx.filter(t => t.jenis === "Keluar").reduce((acc, t) => acc + (t.nilaiTotal || 0), 0);
  const valAwal = filteredAwal.reduce((acc, i) => acc + (i.nilaiTotal || 0), 0);

  const countBast = filteredTx.filter(t => t.jenis === "Masuk").length;
  const countSj = filteredTx.filter(t => t.jenis === "Keluar").length;

  // Koreksi penyesuaian makro BPBD Jabar
  const baseKorTambah = rekap.koreksiTambahRp || 450000000;
  const baseKorKurang = rekap.koreksiKurangRp || 285000000;
  const korTambah = (kalakSelectedSumberDana === "ALL" && kalakSelectedTahun === "ALL") ? baseKorTambah : Math.round(baseKorTambah * (valMasuk / (rekap.totalNilaiPenerimaan || 1)));
  const korKurang = (kalakSelectedSumberDana === "ALL" && kalakSelectedTahun === "ALL") ? baseKorKurang : Math.round(baseKorKurang * (valMasuk / (rekap.totalNilaiPenerimaan || 1)));

  const valAkhir = valAwal + valMasuk + korTambah - korKurang - valKeluar;

  // 1. Update 6 Kartu Finansial Makro
  const elMasuk = document.getElementById("kpi-kalak-nilai-penerimaan");
  if (elMasuk) elMasuk.textContent = formatRupiahSingkat(valMasuk);

  const elBast = document.getElementById("kpi-kalak-bast-count");
  if (elBast) elBast.textContent = `${countBast} BAST Masuk`;

  const elKeluar = document.getElementById("kpi-kalak-nilai-distribusi");
  if (elKeluar) elKeluar.textContent = formatRupiahSingkat(valKeluar);

  const elSj = document.getElementById("kpi-kalak-sj-count");
  if (elSj) elSj.textContent = `${countSj} Surat Jalan`;

  const elAkhir = document.getElementById("kpi-kalak-nilai-stok-akhir");
  if (elAkhir) elAkhir.textContent = formatRupiahSingkat(valAkhir);

  const elAwalNote = document.getElementById("kpi-kalak-stok-awal-note");
  if (elAwalNote) elAwalNote.textContent = `Awal: ${formatRupiahSingkat(valAwal)}`;

  const elKorTambah = document.getElementById("kpi-kalak-koreksi-tambah");
  if (elKorTambah) elKorTambah.textContent = `+${formatRupiahSingkat(korTambah)}`;

  const elKorKurang = document.getElementById("kpi-kalak-koreksi-kurang");
  if (elKorKurang) elKorKurang.textContent = `-${formatRupiahSingkat(korKurang)}`;

  // SKU Murni Kadaluwarsa (Hanya yang benar-benar telah kedaluwarsa <= CUTOFF_DATE)
  const allExpItems = dataStore.daftarBarangKadaluwarsa || [];
  const CUTOFF_DATE = "2026-10-03";
  const strictlyExpiredItems = allExpItems.filter(x => {
    const expDate = x.tglKadaluwarsaAsli || x.tglKadaluwarsa;
    return expDate && expDate <= CUTOFF_DATE;
  });

  const elExpSku = document.getElementById("kpi-kalak-sku-exp");
  if (elExpSku) elExpSku.textContent = `${strictlyExpiredItems.length} SKU`;

  const totalNilaiExp = strictlyExpiredItems.reduce((acc, x) => acc + (x.nilaiTotal || 0), 0);
  const elExpVal = document.getElementById("kpi-kalak-exp-val-risk");
  if (elExpVal) elExpVal.textContent = `Aset Exp: ${formatRupiahSingkat(totalNilaiExp)}`;

  // 2. Render 2 Grafik Utama Finansial
  renderKalakCharts(dataStore, filteredTx);

  // 3. Render Tabel Neraca Finansial 4 Kategori Barang
  renderKalakNeracaTable(dataStore, valAwal, valMasuk, korTambah, korKurang, valKeluar, valAkhir);

  if (window.lucide) {
    lucide.createIcons();
  }
}

// Inisialisasi & Render Grafik KALAK
function renderKalakCharts(dataStore, filteredTx) {
  // --- GRAFIK 1: NILAI BERDASARKAN SUMBER ANGGARAN (GROUPED BAR) ---
  const canvas1 = document.getElementById("chartKalakSumberAnggaran");
  if (canvas1) {
    if (chartKalakSumberAnggaranInstance) {
      chartKalakSumberAnggaranInstance.destroy();
    }

    const danaLabels = ["APBD Provinsi", "APBN (BNPB)", "CSR / Masyarakat"];
    const danaKeys = ["APBD", "APBN", "CSR"];

    const masukPerDana = danaKeys.map(k => {
      return filteredTx.filter(t => t.jenis === "Masuk" && t.sumberDana === k).reduce((acc, t) => acc + (t.nilaiTotal || 0), 0);
    });

    const keluarPerDana = danaKeys.map(k => {
      return filteredTx.filter(t => t.jenis === "Keluar" && t.sumberDana === k).reduce((acc, t) => acc + (t.nilaiTotal || 0), 0);
    });

    const ctx1 = canvas1.getContext("2d");
    chartKalakSumberAnggaranInstance = new Chart(ctx1, {
      type: "bar",
      data: {
        labels: danaLabels,
        datasets: [
          {
            label: "Nilai Penerimaan Masuk (Rp)",
            data: masukPerDana,
            backgroundColor: "#10B981",
            borderRadius: 4
          },
          {
            label: "Nilai Penyaluran Keluar (Rp)",
            data: keluarPerDana,
            backgroundColor: "#F97316",
            borderRadius: 4
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: "top",
            labels: { font: { size: 11, weight: "600" } }
          },
          tooltip: {
            callbacks: {
              label: context => `${context.dataset.label}: Rp ${context.parsed.y.toLocaleString('id-ID')}`
            }
          }
        },
        scales: {
          x: { grid: { display: false } },
          y: {
            beginAtZero: true,
            ticks: {
              callback: val => "Rp " + (val / 1e9).toFixed(1) + " M",
              font: { size: 10 }
            },
            grid: { color: "#F1F5F9" }
          }
        }
      }
    });
  }

  // --- GRAFIK 2: EVALUASI NILAI INDIKATIF LOGISTIK PER KAB/KOTA (MULTI-BAR) ---
  const canvas2 = document.getElementById("chartKalakNilaiIndikatifDaerah");
  if (canvas2) {
    if (chartKalakNilaiIndikatifInstance) {
      chartKalakNilaiIndikatifInstance.destroy();
    }

    const allReq = dataStore.daftarPermohonanDaerah || [];
    const distTx = filteredTx.filter(t => t.jenis === "Keluar");
    const masterDaerah = dataStore.kabKotaJabar || [];

    const daerahStats = masterDaerah.map(d => {
      const valMinta = allReq.filter(r => r.kabKota === d).reduce((acc, r) => acc + (r.estimasiNilai || 0), 0);
      const valSalur = distTx.filter(t => t.kabKotaTujuan === d).reduce((acc, t) => acc + (t.nilaiTotal || 0), 0);
      const paguAlokasi = Math.round(valMinta * 0.75); // Pagu indikatif provinsi ~75% usulan
      return {
        daerah: d.replace("Kabupaten ", "Kab. "),
        valMinta,
        paguAlokasi,
        valSalur
      };
    });

    daerahStats.sort((a, b) => b.valMinta - a.valMinta);
    const top10 = daerahStats.slice(0, 10);

    const labels = top10.map(d => d.daerah);
    const dataMinta = top10.map(d => d.valMinta);
    const dataPagu = top10.map(d => d.paguAlokasi);
    const dataSalur = top10.map(d => d.valSalur);

    const ctx2 = canvas2.getContext("2d");
    chartKalakNilaiIndikatifInstance = new Chart(ctx2, {
      type: "bar",
      data: {
        labels: labels,
        datasets: [
          {
            label: "Nilai Usulan Permohonan (Rp)",
            data: dataMinta,
            backgroundColor: "#3B82F6",
            borderRadius: 3
          },
          {
            label: "Pagu Alokasi Indikatif (Rp)",
            data: dataPagu,
            backgroundColor: "#8B5CF6",
            borderRadius: 3
          },
          {
            label: "Nilai Distribusi Riil (Rp)",
            data: dataSalur,
            backgroundColor: "#10B981",
            borderRadius: 3
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: "top",
            labels: { font: { size: 10, weight: "600" } }
          },
          tooltip: {
            callbacks: {
              label: context => `${context.dataset.label}: Rp ${context.parsed.y.toLocaleString('id-ID')}`
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { font: { size: 9, weight: "600" } }
          },
          y: {
            beginAtZero: true,
            ticks: {
              callback: val => "Rp " + (val / 1e6).toFixed(0) + " Jt",
              font: { size: 10 }
            },
            grid: { color: "#F1F5F9" }
          }
        }
      }
    });
  }
}

// Render Tabel Neraca Finansial 4 Kategori Barang (Laporan Pertanggungjawaban Aset)
function renderKalakNeracaTable(dataStore, totAwal, totMasuk, korTambah, korKurang, totKeluar, totAkhir) {
  const tbody = document.getElementById("tbody-neraca-kalak");
  const tfoot = document.getElementById("tfoot-neraca-kalak");
  if (!tbody || !tfoot) return;

  const categories = ["Sandang", "Pangan", "Papan", "Logistik Lainnya"];
  const katMaster = dataStore.kategoriMaster || [];
  const allTx = dataStore.riwayatMutasiSemua || [];

  const katColors = {
    "Pangan": { bg: "#EFF6FF", text: "#1D4ED8", border: "#BFDBFE" },
    "Sandang": { bg: "#F5F3FF", text: "#6D28D9", border: "#DDD6FE" },
    "Papan": { bg: "#FFF7ED", text: "#C2410C", border: "#FED7AA" },
    "Logistik Lainnya": { bg: "#ECFDF5", text: "#047857", border: "#A7F3D0" }
  };

  const rowsData = categories.map((kat, idx) => {
    const kObj = katMaster.find(m => m.nama === kat);
    const awalKat = kObj ? (kObj.nilaiStokAwal || 0) : 0;
    const masukKat = allTx.filter(t => t.jenis === "Masuk" && t.kategori === kat).reduce((acc, t) => acc + (t.nilaiTotal || 0), 0);
    const keluarKat = allTx.filter(t => t.jenis === "Keluar" && t.kategori === kat).reduce((acc, t) => acc + (t.nilaiTotal || 0), 0);
    
    // Proporsi pembagian koreksi makro
    const proporsi = totMasuk > 0 ? (masukKat / totMasuk) : 0.25;
    const korTambahKat = Math.round(korTambah * proporsi);
    const korKurangKat = Math.round(korKurang * proporsi);
    const korNettoKat = korTambahKat - korKurangKat;
    
    const akhirKat = awalKat + masukKat + korNettoKat - keluarKat;
    const porsiAset = totAkhir > 0 ? ((akhirKat / totAkhir) * 100).toFixed(1) : "0.0";

    return {
      no: idx + 1,
      kategori: kat,
      awal: awalKat,
      masuk: masukKat,
      korNetto: korNettoKat,
      keluar: keluarKat,
      akhir: akhirKat,
      porsi: porsiAset
    };
  });

  tbody.innerHTML = rowsData.map(r => {
    const kColor = katColors[r.kategori] || { bg: "#F1F5F9", text: "#334155", border: "#CBD5E1" };
    return `
      <tr>
        <td style="text-align: center; font-weight: 700; color: #64748B;">${r.no}</td>
        <td>
          <span class="badge-kat-pill" style="background-color: ${kColor.bg}; color: ${kColor.text}; border: 1px solid ${kColor.border};">
            ${r.kategori}
          </span>
        </td>
        <td class="num-cell">${formatRupiahLengkap(r.awal)}</td>
        <td class="num-cell positive">+${formatRupiahLengkap(r.masuk)}</td>
        <td class="num-cell" style="color: ${r.korNetto >= 0 ? '#15803D' : '#DC2626'}; font-weight: 700;">
          ${r.korNetto >= 0 ? '+' : ''}${formatRupiahLengkap(r.korNetto)}
        </td>
        <td class="num-cell negative">-${formatRupiahLengkap(r.keluar)}</td>
        <td class="num-cell total-highlight">${formatRupiahLengkap(r.akhir)}</td>
        <td style="text-align: center; font-weight: 700; color: #0E63AA;">
          ${r.porsi}%
        </td>
      </tr>
    `;
  }).join("");

  const totalKorNetto = korTambah - korKurang;
  tfoot.innerHTML = `
    <tr>
      <td colspan="2" style="text-align: center; font-weight: 800; color: #0F172A;">
        TOTAL VALUASI NERACA GUDANG
      </td>
      <td class="num-cell" style="font-weight: 800;">${formatRupiahLengkap(totAwal)}</td>
      <td class="num-cell positive" style="font-weight: 800;">+${formatRupiahLengkap(totMasuk)}</td>
      <td class="num-cell" style="font-weight: 800; color: ${totalKorNetto >= 0 ? '#15803D' : '#DC2626'};">
        ${totalKorNetto >= 0 ? '+' : ''}${formatRupiahLengkap(totalKorNetto)}
      </td>
      <td class="num-cell negative" style="font-weight: 800;">-${formatRupiahLengkap(totKeluar)}</td>
      <td class="num-cell total-highlight" style="font-size: 0.95rem; font-weight: 800; color: #0E63AA;">
        ${formatRupiahLengkap(totAkhir)}
      </td>
      <td style="text-align: center; font-weight: 800; color: #0E63AA;">
        100.0%
      </td>
    </tr>
  `;
}

// Ekspor Neraca ke Berkas Excel (.xlsx) dengan SheetJS
function exportNeracaKalakToExcel() {
  if (typeof XLSX === "undefined") {
    showToast("Pustaka SheetJS belum termuat sempurna");
    return;
  }

  const table = document.getElementById("table-neraca-kalak");
  if (!table) return;

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(table);
  XLSX.utils.book_append_sheet(wb, ws, "Neraca Finansial Logistik");
  
  const fileName = `Neraca_Finansial_Logistik_BPBD_Jabar_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, fileName);
  showToast(`Neraca berhasil diekspor ke: ${fileName}`);
}

