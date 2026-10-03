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
let parameterXHari = 30; // Default ambang batas kadaluwarsa (X Hari)
let activeKadaluwarsaTab = "semua"; // "semua" | "expired" | "warning" | "aman"

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

  // 2. Isi Dropdown 27 Kab/Kota pada Modal Distribusi (SJ)
  populateSelect27KabKota();

  // 3. Muat Data Langsung dari File Excel (Direct Integration)
  await loadExcelDatabaseDirectly();

  console.log("Dashboard Operator E-LUNA Berhasil Diinisialisasi.");
});

// ==========================================
// 2. INTEGRASI LANGSUNG DENGAN FILE EXCEL (.XLSX)
// ==========================================

// Memuat langsung file Excel 'database_logistik_bpbd_jabar.xlsx' dari server
async function loadExcelDatabaseDirectly() {
  const syncBtn = document.getElementById("btn-sync-excel");
  const syncIcon = document.getElementById("icon-sync-spin");

  if (syncIcon) syncIcon.classList.add("spin-animation");
  updateDbStatus("loading", "Membaca database_logistik_bpbd_jabar.xlsx...");

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
    console.warn("Direct Excel fetch gagal, mengaktifkan data lokal fallback:", err);
    updateDbStatus("fallback", "Menggunakan Data Lokal (Fallback)");
    
    // Fallback ke data.js jika offline
    if (window.ELUNA_OPERATOR_DATA) {
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
      kabKotaTujuan: (keluar > 0) ? tujuan : "-"
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
        noDokumen: doc
      });
    });
  }

  // 3. Baca Sheet 'Stok Awal'
  const wsAwal = workbook.Sheets["Stok Awal"];
  const stokPerKat = { Sandang: 0, Pangan: 0, Papan: 0, "Logistik Lainnya": 0 };
  if (wsAwal) {
    const rawAwal = XLSX.utils.sheet_to_json(wsAwal);
    rawAwal.forEach(r => {
      const kat = r["Kategori Barang"];
      const awal = parseInt(r["Stok Akhir Periode Sebelumnya"] || 0, 10);
      if (stokPerKat.hasOwnProperty(kat)) {
        stokPerKat[kat] += awal;
      }
    });
  }

  // Perbarui state global aplikasi
  window.ELUNA_OPERATOR_DATA = {
    metadata: {
      systemName: "E-LUNA OPERATOR GUDANG",
      subTitle: "Sistem Pemantauan Operasional Logistik Kebencanaan",
      instansi: "BPBD Provinsi Jawa Barat",
      gudangAktif: "Gudang Logistik & Peralatan Provinsi Jawa Barat",
      lokasi: "Jl. Soekarno-Hatta No. 629, Kota Bandung",
      petugas: "Nadhif Althafu Rutama",
      tanggalAcuanSistem: "2026-10-03",
      parameterXDefaultHari: 30,
      sourceDatabase: sourceFilename,
      totalEntries: transactions.length
    },
    kategoriMaster: [
      { id: "sandang", nama: "Sandang", deskripsi: "Selimut, Matras, Sarung, Pakaian Dewasa, Pakaian Anak", color: "#8B5CF6", bgLight: "#F5F3FF", stokAwal: stokPerKat["Sandang"] || 53400, satuan: "Unit/Paket" },
      { id: "pangan", nama: "Pangan", deskripsi: "Beras, Makanan Siap Saji, Susu Formula, Biskuit, Lauk Pauk, Air Mineral", color: "#3B82F6", bgLight: "#EFF6FF", stokAwal: stokPerKat["Pangan"] || 210000, satuan: "Unit/Paket" },
      { id: "papan", nama: "Papan", deskripsi: "Tenda Pengungsi, Tenda Keluarga, Terpal, Seng", color: "#F97316", bgLight: "#FFF7ED", stokAwal: stokPerKat["Papan"] || 21600, satuan: "Unit" },
      { id: "lainnya", nama: "Logistik Lainnya", deskripsi: "Perahu Karet, Genset, Chainsaw, Family Kit, Sanitasi Kit", color: "#10B981", bgLight: "#ECFDF5", stokAwal: stokPerKat["Logistik Lainnya"] || 16890, satuan: "Unit/Paket" }
    ],
    kabKotaJabar: DAFTAR_27_KAB_KOTA_JABAR,
    daftarBarangKadaluwarsa: expItems,
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

// Trigger Input File Excel
function triggerExcelUpload() {
  const inp = document.getElementById("input-excel-file");
  if (inp) inp.click();
}

// Handler Saat Pengguna Memilih / Mengunggah File Excel dari Komputer
function handleExcelFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      processExcelWorkbook(data, file.name);
      showToast(`File Excel '${file.name}' berhasil dimuat ke Dashboard!`);
    } catch (err) {
      console.error("Gagal membaca file Excel yang diunggah:", err);
      alert(`Gagal membaca file Excel: ${err.message}`);
    }
  };
  reader.readAsArrayBuffer(file);
}

// Unduh File Excel Langsung dari Dashboard
function downloadExcelDatabase() {
  const link = document.createElement("a");
  link.href = "database_logistik_bpbd_jabar.xlsx";
  link.download = "database_logistik_bpbd_jabar.xlsx";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  showToast("Mengunduh file 'database_logistik_bpbd_jabar.xlsx'...");
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

// State dan Handler Filter Mutasi (Semua / Masuk Saja / Keluar Saja)
let currentMutasiTypeFilter = "all";

function setMutasiTypeFilter(type) {
  currentMutasiTypeFilter = type;

  const tabAll = document.getElementById("tab-mutasi-all");
  const tabIn = document.getElementById("tab-mutasi-masuk");
  const tabOut = document.getElementById("tab-mutasi-keluar");
  if (tabAll) tabAll.classList.toggle("active", type === "all");
  if (tabIn) tabIn.classList.toggle("active", type === "masuk");
  if (tabOut) tabOut.classList.toggle("active", type === "keluar");

  renderActiveMutasiTable();
}

function renderActiveMutasiTable() {
  if (!window.ELUNA_OPERATOR_DATA) return;
  const titleMutasi = document.getElementById("title-mutasi-tabel");
  let txToShow = [];

  if (currentFilterMode === "default10") {
    const allTx = ELUNA_OPERATOR_DATA.riwayatMutasiSemua;
    if (currentMutasiTypeFilter === "all") {
      txToShow = allTx.slice(0, 10);
      if (titleMutasi) titleMutasi.textContent = "Daftar Inputan Mutasi (10 Terakhir)";
    } else if (currentMutasiTypeFilter === "masuk") {
      txToShow = allTx.filter(t => t.jenis === "Masuk").slice(0, 10);
      if (titleMutasi) titleMutasi.textContent = "Daftar Inputan Mutasi (10 Masuk Terakhir)";
    } else if (currentMutasiTypeFilter === "keluar") {
      txToShow = allTx.filter(t => t.jenis === "Keluar").slice(0, 10);
      if (titleMutasi) titleMutasi.textContent = "Daftar Inputan Mutasi (10 Keluar Terakhir)";
    }
  } else {
    // Mode 'all' atau 'filtered'
    const baseTx = activeTransactions;
    if (currentMutasiTypeFilter === "masuk") {
      txToShow = baseTx.filter(t => t.jenis === "Masuk");
      if (titleMutasi) titleMutasi.textContent = `Daftar Inputan Mutasi Masuk (${txToShow.length} Baris)`;
    } else if (currentMutasiTypeFilter === "keluar") {
      txToShow = baseTx.filter(t => t.jenis === "Keluar");
      if (titleMutasi) titleMutasi.textContent = `Daftar Inputan Mutasi Keluar (${txToShow.length} Baris)`;
    } else {
      txToShow = baseTx;
      if (titleMutasi) titleMutasi.textContent = `Daftar Inputan Mutasi (${txToShow.length} Baris)`;
    }
  }

  renderMutasiOperatorTable(txToShow);
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

  if (elMasuk) elMasuk.textContent = totalMasukUnit.toLocaleString("id-ID");
  if (elBast) elBast.textContent = `${countBast} Dokumen BAST`;

  if (elKeluar) elKeluar.textContent = totalKeluarUnit.toLocaleString("id-ID");
  if (elSj) elSj.textContent = `${countSj} Dokumen Surat Jalan`;

  // 2. Hitung Neraca Stok 4 Kategori: Stok Akhir = Stok Awal (Konstan) + Masuk - Keluar
  let totalAkumulasiStokAkhir = 0;
  const neracaData = ELUNA_OPERATOR_DATA.kategoriMaster.map(km => {
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

  // 3. Render Komponen & Grafik
  renderNeracaStokCards(neracaData);
  renderVisualisasiCharts(neracaData, kabKotaVolumeMap, totalMasukUnit, totalKeluarUnit);
  renderActiveMutasiTable();

  if (window.lucide) {
    lucide.createIcons();
  }
}

// ==========================================
// 5. NERACA STOK 4 KATEGORI (STOK AWAL + MASUK - KELUAR = STOK AKHIR)
// ==========================================
function renderNeracaStokCards(neracaList) {
  const container = document.getElementById("neraca-grid-container");
  if (!container) return;

  container.innerHTML = neracaList.map(item => `
    <div class="neraca-kategori-box">
      <div class="neraca-kat-title">
        <span>${item.nama}</span>
        <span style="font-size: 0.725rem; color: #64748B; font-weight: 500;">${item.satuan}</span>
      </div>
      <div class="neraca-row">
        <span>Stok Awal:</span>
        <strong style="color: #334155;">${item.stokAwal.toLocaleString("id-ID")}</strong>
      </div>
      <div class="neraca-row" style="color: #10B981;">
        <span>(+) Penerimaan:</span>
        <strong>+${item.masuk.toLocaleString("id-ID")}</strong>
      </div>
      <div class="neraca-row" style="color: var(--eluna-orange);">
        <span>(-) Distribusi:</span>
        <strong>-${item.keluar.toLocaleString("id-ID")}</strong>
      </div>
      <div class="neraca-row result">
        <span>Stok Akhir:</span>
        <span class="val">${item.stokAkhir.toLocaleString("id-ID")}</span>
      </div>
    </div>
  `).join("");
}

// ==========================================
// 6. VISUALISASI STATISTIK (PENERIMAAN SUMBER ANGGARAN & DISTRIBUSI 27 KAB/KOTA)
// ==========================================
function renderVisualisasiCharts(neracaList, kabKotaMap, totalMasuk, totalKeluar) {
  const badgeMasuk = document.getElementById("badge-vol-masuk");
  const badgeKeluar = document.getElementById("badge-vol-keluar");

  if (badgeMasuk) badgeMasuk.textContent = `${totalMasuk.toLocaleString("id-ID")} Unit Masuk`;
  if (badgeKeluar) badgeKeluar.textContent = `${totalKeluar.toLocaleString("id-ID")} Unit Keluar`;

  // --- Chart 1: Penerimaan Berdasarkan Sumber Anggaran (APBD, APBN, CSR) ---
  const canvasPenerimaan = document.getElementById("chartPenerimaanSumberAnggaran") || document.getElementById("chartPenerimaanKategori");
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

    const ctxP = canvasPenerimaan.getContext("2d");
    if (chartPenerimaanInst) chartPenerimaanInst.destroy();

    chartPenerimaanInst = new Chart(ctxP, {
      type: "bar",
      data: {
        labels: ["APBD", "APBN", "CSR"],
        datasets: [{
          label: "Volume Penerimaan (Unit)",
          data: [sumberAnggaranMasuk.APBD, sumberAnggaranMasuk.APBN, sumberAnggaranMasuk.CSR],
          backgroundColor: ["#0E63AA", "#ED7D18", "#10B981"],
          borderRadius: 4,
          barPercentage: 0.45
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (context) => ` Penerimaan ${context.label}: ${context.parsed.y.toLocaleString("id-ID")} Unit`
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
    // Tampilkan seluruh daerah yang memiliki mutasi distribusi (urut terbesar)
    const sortedDaerah = Object.entries(kabKotaMap)
      .sort((a, b) => b[1] - a[1])
      .filter(d => d[1] > 0 || currentFilterMode === "all")
      .slice(0, 10); // 10 daerah teratas

    const ctxD = canvasDistribusi.getContext("2d");
    if (chartDistribusiInst) chartDistribusiInst.destroy();

    chartDistribusiInst = new Chart(ctxD, {
      type: "bar",
      data: {
        labels: sortedDaerah.map(d => d[0].replace("Kabupaten ", "Kab. ")),
        datasets: [{
          label: "Volume Distribusi (Unit)",
          data: sortedDaerah.map(d => d[1]),
          backgroundColor: "#ED7D18",
          borderRadius: 4,
          barPercentage: 0.55
        }]
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (context) => ` Distribusi: ${context.parsed.x.toLocaleString("id-ID")} Unit`
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

  // Filter List sesuai Tab Aktif
  let filtered = analyzed;
  if (activeKadaluwarsaTab === "expired") filtered = analyzed.filter(i => i.statusType === "expired");
  if (activeKadaluwarsaTab === "warning") filtered = analyzed.filter(i => i.statusType === "warning");
  if (activeKadaluwarsaTab === "aman") filtered = analyzed.filter(i => i.statusType === "safe");

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
// 8. TABEL RIWAYAT TRANSAKSI OPERATOR
// ==========================================
function renderMutasiOperatorTable(transactions) {
  const tbody = document.getElementById("tbody-mutasi-operator");
  if (!tbody) return;

  if (transactions.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align: center; padding: 2rem; color: #94A3B8;">Tidak ada data mutasi logistik untuk filter ini.</td></tr>`;
    return;
  }

  tbody.innerHTML = transactions.map((tx, idx) => {
    const isMasuk = tx.jenis === "Masuk";
    return `
      <tr>
        <td class="text-center" style="font-weight: 700; color: #64748B;">${idx + 1}</td>
        <td style="font-size: 0.775rem; color: #64748B;">${tx.tanggal}</td>
        <td style="font-family: monospace; font-weight: 700; color: #0E63AA;">${tx.noDokumen}</td>
        <td>
          <span style="display: inline-block; padding: 0.15rem 0.5rem; border-radius: 4px; font-size: 0.725rem; font-weight: 700; background-color: ${isMasuk ? '#DCFCE7' : '#FFF4EB'}; color: ${isMasuk ? '#16A34A' : '#D97706'};">
            ${tx.jenis}
          </span>
        </td>
        <td style="font-weight: 600;">${tx.kategori}</td>
        <td>${tx.uraianBarang}</td>
        <td class="text-right" style="font-weight: 800; color: ${isMasuk ? '#16A34A' : '#ED7D18'};">
          ${isMasuk ? '+' : '-'}${tx.jumlah.toLocaleString("id-ID")} <span style="font-size: 0.7rem; color: #94A3B8;">${tx.satuan}</span>
        </td>
        <td style="font-size: 0.8rem; font-weight: 600;">${tx.asalTujuan}</td>
      </tr>
    `;
  }).join("");
}

// ==========================================
// 9. HELPER & UTILITIES
// ==========================================

// Isi Dropdown 27 Kabupaten/Kota pada Form Surat Jalan
function populateSelect27KabKota() {
  const sel = document.getElementById("form-sj-tujuan");
  if (!sel) return;
  sel.innerHTML = DAFTAR_27_KAB_KOTA_JABAR.map(kab => `
    <option value="${kab}">${kab}</option>
  `).join("");
}

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

// Tampilkan Toast Notifikasi
function showToast(msg) {
  const c = document.getElementById("toast-container");
  if (!c) return;
  const t = document.createElement("div");
  t.className = "toast-item";
  t.textContent = msg;
  c.appendChild(t);
  setTimeout(() => {
    t.style.opacity = "0";
    t.style.transition = "opacity 0.3s ease";
    setTimeout(() => t.remove(), 300);
  }, 3500);
}

// Modal Controllers
function openModal(id) {
  const m = document.getElementById(id);
  if (m) m.showModal();
}

function closeModal(id) {
  const m = document.getElementById(id);
  if (m) m.close();
}

function scrollToSection(id) {
  const el = document.getElementById(id);
  if (el) el.scrollIntoView({ behavior: "smooth" });
}

// Submit Form Transaksi Manual BAST & SJ
function submitBAST(e) {
  e.preventDefault();
  const no = document.getElementById("form-bast-no").value;
  const kat = document.getElementById("form-bast-kat").value;
  const item = document.getElementById("form-bast-item").value;
  const qty = parseInt(document.getElementById("form-bast-qty").value, 10);
  const asal = document.getElementById("form-bast-asal").value;

  const now = new Date();
  const tglStr = now.toISOString().slice(0, 10);

  const newTx = {
    id: `TX-${Date.now().toString().slice(-4)}`,
    tanggal: tglStr,
    bulan: now.toISOString().slice(0, 7),
    tahun: now.getFullYear().toString(),
    noDokumen: no,
    jenis: "Masuk",
    kategori: kat,
    uraianBarang: item,
    jumlah: qty,
    satuan: "Unit/Paket",
    sumberDana: asal,
    asalTujuan: `Sumber: ${asal}`,
    kabKotaTujuan: "-"
  };

  ELUNA_OPERATOR_DATA.riwayatMutasiSemua.unshift(newTx);
  closeModal("modal-bast");

  if (currentFilterMode === "default10") {
    activeTransactions = ELUNA_OPERATOR_DATA.riwayatMutasiSemua.slice(0, 10);
  } else {
    activeTransactions.unshift(newTx);
  }

  updateDashboardCore();
  showToast(`Penerimaan ${no} (+${qty.toLocaleString("id-ID")} ${kat}) berhasil disimpan!`);
}

function submitSJ(e) {
  e.preventDefault();
  const no = document.getElementById("form-sj-no").value;
  const tujuan = document.getElementById("form-sj-tujuan").value;
  const kat = document.getElementById("form-sj-kat").value;
  const item = document.getElementById("form-sj-item").value;
  const qty = parseInt(document.getElementById("form-sj-qty").value, 10);

  const now = new Date();
  const tglStr = now.toISOString().slice(0, 10);

  const newTx = {
    id: `TX-${Date.now().toString().slice(-4)}`,
    tanggal: tglStr,
    bulan: now.toISOString().slice(0, 7),
    tahun: now.getFullYear().toString(),
    noDokumen: no,
    jenis: "Keluar",
    kategori: kat,
    uraianBarang: item,
    jumlah: qty,
    satuan: "Unit/Paket",
    sumberDana: "APBD",
    asalTujuan: `Tujuan: ${tujuan}`,
    kabKotaTujuan: tujuan
  };

  ELUNA_OPERATOR_DATA.riwayatMutasiSemua.unshift(newTx);
  closeModal("modal-sj");

  if (currentFilterMode === "default10") {
    activeTransactions = ELUNA_OPERATOR_DATA.riwayatMutasiSemua.slice(0, 10);
  } else {
    activeTransactions.unshift(newTx);
  }

  updateDashboardCore();
  showToast(`Surat Jalan ${no} (-${qty.toLocaleString("id-ID")} ke ${tujuan}) berhasil diterbitkan!`);
}
