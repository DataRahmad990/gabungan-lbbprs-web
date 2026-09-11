// Gabungan generik form konven (LBBPRK-XXXX). Port dari core/form_konven.py.
// Label kolom dari merged-cell header (ws['!merges']). Nama via 0016.
import * as S from "./sandiKonven.js";
import { detectBank, discoverBranches } from "./bank.js";
import { buildNameMap, lookupName, nameCount } from "./pihakLawan.js";
import { buildBankNameMap, lookupBankName, bankNameCount } from "./bankNames.js";
import * as H from "./helpers.js";

function readSheet(bytes, XLSX) {
  const wb = XLSX.read(bytes, { type: "array", raw: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const merges = (ws["!merges"] || []).map(m => ({ r: m.s.r, c0: m.s.c, c1: m.e.c }));
  const ncols = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]).e.c + 1 : 0;
  return { aoa, merges, ncols };
}

function isDigits(v, max) {
  let s = String(v == null ? "" : v).trim();
  if (s.endsWith(".0")) s = s.slice(0, -2);
  return new RegExp(`^\\d{1,${max}}$`).test(s);
}

function findDataStart(aoa) {
  for (let r = 14; r < Math.min(60, aoa.length); r++) {
    if (isDigits(H.cell(aoa, r, 4), 3)) return r;
  }
  return 18;
}

function txt(v) { return String(v == null ? "" : v).trim(); }

// Deteksi robust (dipakai syariah): baris header = baris dengan paling banyak
// SEL LABEL TEKS (mengandung huruf, bukan angka murni) di kolom 3..N; data mulai
// setelah blok header. Menangani kasus kolom-4 bukan nomor urut 1-3 digit
// (mis. Tabungan/Hapus Buku di mana kolom-4 = Golongan Nasabah 4 digit).
function wordCells(aoa, r, ncols) {
  let n = 0;
  for (let c = 3; c < ncols; c++) {
    const v = txt(H.cell(aoa, r, c));
    if (v && /[A-Za-z]/.test(v)) n++;
  }
  return n;
}
// Baris DATA = sel bernilai (angka/tanggal/kode tanpa huruf) lebih banyak dari
// sel label (mengandung huruf). Baris header/sub-header didominasi label teks.
function isDataRow(aoa, r, ncols) {
  let words = 0, vals = 0;
  for (let c = 3; c < ncols; c++) {
    const v = txt(H.cell(aoa, r, c));
    if (!v) continue;
    if (/[A-Za-z]/.test(v)) words++;
    else vals++;
  }
  return vals > words;
}
function findDataStartRobust(aoa, ncols) {
  const lim = Math.min(45, aoa.length);
  let hr = -1, best = 1;
  for (let r = 6; r < lim; r++) {
    const w = wordCells(aoa, r, ncols);
    if (w > best) { best = w; hr = r; }
  }
  if (hr < 0) return findDataStart(aoa);
  const end = Math.min(hr + 15, aoa.length);
  for (let r = hr + 1; r < end; r++) if (isDataRow(aoa, r, ncols)) return r;
  return aoa.length; // tidak ada baris data (form kosong) -> jangan salah ambil sub-header
}
function isNumLabel(s) { return s !== "" && /^[\d.]+$/.test(s); }

function headerLabels(aoa, merges, ncols, ds) {
  let bestHr = ds - 1, bestCnt = -1;
  for (let r = Math.max(0, ds - 4); r < ds; r++) {
    let cnt = 0;
    for (let c = 3; c < ncols; c++) { const v = txt(H.cell(aoa, r, c)); if (v && !isNumLabel(v)) cnt++; }
    if (cnt >= bestCnt) { bestCnt = cnt; bestHr = r; }
  }
  const hr = bestHr;
  const parent = {};
  for (const m of merges) {
    if (m.r >= hr && m.r < ds && (m.c1 - m.c0 + 1) < ncols * 0.6) {
      const v = txt(H.cell(aoa, m.r, m.c0));
      if (v) for (let c = m.c0; c <= m.c1; c++) if (!(c in parent)) parent[c] = v;
    }
  }
  const labels = {};
  for (let c = 3; c < ncols; c++) {
    let base = txt(H.cell(aoa, hr, c)) || parent[c] || "";
    if (isNumLabel(base)) base = parent[c] || "";
    const sub = hr + 1 < ds ? txt(H.cell(aoa, hr + 1, c)) : "";
    const label = (sub && sub !== base) ? `${base} ${sub}`.trim() : base;
    if (label) labels[c] = label;
  }
  return labels;
}

// Excel writer merangkai XML jadi string raksasa; form seukuran Daftar Tabungan
// (ratusan ribu baris) bikin prosesnya kehabisan memori. Untuk kasus itu outputnya
// CSV: isinya lengkap, jauh lebih kecil, dan tetap bisa dibuka Excel.
function toCsv(cols, data) {
  const esc = v => {
    if (v === undefined || v === null) return "";
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    const s = String(v);
    return /[",\r\n;]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const parts = ["\ufeff" + cols.map(esc).join(";") + "\r\n"];
  const CHUNK = 5000;
  let buf = [];
  for (const row of data) {
    buf.push(cols.map(c => esc(row[c])).join(";"));
    if (buf.length >= CHUNK) { parts.push(buf.join("\r\n") + "\r\n"); buf = []; }
  }
  if (buf.length) parts.push(buf.join("\r\n") + "\r\n");
  return new TextEncoder().encode(parts.join(""));
}

function sheet(XLSX, cols, data) {
  const aoa = [cols, ...data.map(r => cols.map(c => { const v = r[c]; return v === undefined ? "" : v; }))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const fill = { fill: { fgColor: { rgb: "1F4E79" } }, font: { color: { rgb: "FFFFFF" }, bold: true, sz: 10 } };
  for (let c = 0; c < cols.length; c++) { const a = XLSX.utils.encode_cell({ r: 0, c }); if (ws[a]) ws[a].s = fill; }
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };
  return ws;
}

// cfg (opsional) untuk mendukung report type lain (mis. syariah LBBPRS):
//   { reportPrefix: "LBBPRS", translateMap: <sandi 17>, translate: <fn> }
// Default = konven (LBBPRK, SEOJK 16) supaya pemanggilan lama tetap jalan.
export function makeFormProcessor(formCode, title, namePrefix, cfg = {}) {
  return function (files, period, XLSX) {
    const reportPrefix = cfg.reportPrefix || "LBBPRK";
    const TMAP = cfg.translateMap || S.TRANSLATE_MAP;
    const TR = cfg.translate || S.translate;
    const prefix = `${reportPrefix}-${formCode}-`;
    const branches = discoverBranches(files, prefix);
    if (!branches.length) return null;  // form ga ada -> skip

    const nameMap = buildNameMap(files, XLSX);
    const rawBranch = {};
    const labelsByCol = {};
    const nonempty = new Set();
    for (const code of branches) {
      const name = Object.keys(files).find(n => { const b = n.split("/").pop(); return b.startsWith(prefix) && (b.endsWith(`-${code}.xls`) || b.endsWith(`-${code}_part1.xls`)); });
      if (!name) continue;
      const { aoa, merges, ncols } = readSheet(files[name], XLSX);
      const ds = cfg.robustDataStart ? findDataStartRobust(aoa, ncols) : findDataStart(aoa);
      const labels = headerLabels(aoa, merges, ncols, ds);
      const used = Object.keys(labels).map(Number).sort((a, b) => a - b);
      for (const c of used) if (!(c in labelsByCol)) labelsByCol[c] = labels[c];
      const rows = [];
      for (let r = ds; r < aoa.length; r++) {
        const key = used.length ? txt(H.cell(aoa, r, used[0])) : "";
        const cells = {};
        let any = false;
        for (const c of used) { const v = H.cell(aoa, r, c); cells[c] = v; if (txt(v)) { any = true; nonempty.add(c); } }
        if (key.toUpperCase() === "JUMLAH" || !any) continue;
        rows.push(cells);
      }
      rawBranch[code] = rows;
    }

    // Forward-fill kolom kunci yang hanya terisi di baris pertama grup (mis. Nomor
    // Rekening di Sindikasi: 1 rekening banyak baris peserta, rek cuma di baris awal).
    // Baris lanjutan (blank) mewarisi nilai dari baris sebelumnya dalam cabang yang sama.
    if (cfg.fillDown && cfg.fillDown.length) {
      const fillCols = Object.keys(labelsByCol).map(Number)
        .filter(c => cfg.fillDown.some(s => (labelsByCol[c] || "").includes(s)));
      for (const code of Object.keys(rawBranch)) {
        const last = {};
        for (const cells of rawBranch[code]) {
          for (const c of fillCols) {
            const v = cells[c];
            if (v === undefined || v === null || String(v).trim() === "") {
              if (last[c] !== undefined) { cells[c] = last[c]; nonempty.add(c); }
            } else last[c] = v;
          }
        }
      }
    }

    const kept = Object.keys(labelsByCol).map(Number).sort((a, b) => a - b).filter(c => nonempty.has(c));
    const seen = {}, colLabel = {};
    for (const c of kept) {
      let lab = labelsByCol[c];
      if (seen[lab]) { seen[lab]++; lab = `${labelsByCol[c]} #${seen[labelsByCol[c]]}`; } else seen[lab] = 1;
      colLabel[c] = lab;
    }
    // Kolom yang isinya SANDI BANK (mis. "Sandi Bank Peserta Sindikasi"): dampingi
    // dengan kolom nama bank, supaya pembaca tidak cuma dapat kode angka.
    const bankMap = (cfg.bankNameCols && cfg.bankNameCols.length)
      ? buildBankNameMap(files, period, XLSX) : null;
    const bankNameLabel = {};
    if (bankMap) {
      for (const c of kept) {
        const lab = labelsByCol[c] || "";
        if (!cfg.bankNameCols.some(sub => lab.includes(sub))) continue;
        bankNameLabel[c] = /^sandi\b/i.test(lab) ? lab.replace(/^sandi\b/i, "Nama") : `Nama untuk ${lab}`;
      }
    }

    const idCol = kept.find(c => labelsByCol[c].includes("ID Pihak Lawan"));
    const nikCol = kept.find(c => { const l = labelsByCol[c] || ""; return l.includes("No. Identitas") || l.includes("Nomor Identitas"); });
    const hasName = !!(nameCount(nameMap) && idCol !== undefined);
    const colsOrder = ["Cabang"];
    for (const c of kept) {
      colsOrder.push(colLabel[c]);
      if (bankNameLabel[c]) colsOrder.push(bankNameLabel[c]);
    }
    if (hasName) colsOrder.push("Nama");

    function toRow(code, cells) {
      const row = { "Cabang": code };
      for (const c of kept) {
        let v = cells[c] === undefined ? "" : cells[c];
        const base = labelsByCol[c];
        if (TMAP[base] && v !== null && v !== "" && v !== 0) v = TR(v, TMAP[base]);
        row[colLabel[c]] = v;
        if (bankNameLabel[c]) row[bankNameLabel[c]] = lookupBankName(bankMap, cells[c]);
      }
      if (hasName) row["Nama"] = lookupName(nameMap, cells[idCol], nikCol !== undefined ? cells[nikCol] : undefined);
      return row;
    }
    const perBranch = {};
    let all = [];
    for (const code of Object.keys(rawBranch).sort()) { perBranch[code] = rawBranch[code].map(c => toRow(code, c)); all = all.concat(perBranch[code]); }

    const wb = XLSX.utils.book_new();
    const tag = detectBank(files, XLSX).tag;
    // Form sangat besar (mis. Tabungan puluhan ribu baris): menulis sheet per-cabang
    // yang menduplikasi "SEMUA CABANG" membuat xlsx-js-style gagal ("Invalid array
    // length"). Lewati sheet per-cabang bila data besar; seluruh data tetap ada di
    // "SEMUA CABANG" (dapat difilter lewat kolom "Cabang").
    const PER_BRANCH_MAX = 40000;
    const skipBranch = all.length > PER_BRANCH_MAX;
    // xlsx-js-style merangkai XML tiap sheet jadi SATU string JS. Form besar
    // (mis. Daftar Tabungan ~148 ribu baris) melewati batas panjang string V8 dan
    // bikin "RangeError: Invalid array length" saat XLSX.write. Solusinya: pecah
    // SEMUA CABANG jadi beberapa sheet. Tidak ada baris yang dibuang.
    const SHEET_MAX = 40000;
    const chunkCount = Math.max(1, Math.ceil(all.length / SHEET_MAX));
    // Di atas ambang ini xlsx tidak realistis (80 ribu baris saja sudah 81 MB dan
    // 12 detik; di atas itu proses kehabisan memori). Keluarkan CSV lengkap.
    const CSV_MIN_ROWS = 60000;
    if (all.length >= CSV_MIN_ROWS) {
      return {
        filename: `${namePrefix}_${tag}_${period.periodeLabel}.csv`,
        data: toCsv(colsOrder, all),
        summary: { jumlah_baris: all.length, format: "csv" },
        warning: `${title}: ${all.length.toLocaleString("id-ID")} baris, terlalu besar untuk file Excel. Diturunkan sebagai CSV (pemisah titik koma, semua baris lengkap). Buka langsung di Excel.`,
      };
    }
    // Statistik nama bank dihitung sebelum sheet RINGKASAN ditulis.
    let bankStat = null;
    if (bankMap && Object.keys(bankNameLabel).length) {
      const kodeSet = new Set(), belum = new Set();
      for (const c of Object.keys(bankNameLabel).map(Number)) {
        for (const code of Object.keys(rawBranch)) {
          for (const cells of rawBranch[code]) {
            const sandi = String(cells[c] ?? "").trim().replace(/\.0$/, "");
            if (!sandi) continue;
            kodeSet.add(sandi);
            if (!lookupBankName(bankMap, sandi)) belum.add(sandi);
          }
        }
      }
      if (kodeSet.size) bankStat = { total: kodeSet.size, bernama: kodeSet.size - belum.size, belum: [...belum].sort() };
    }

    const ringkasan = [[`${title} (${reportPrefix}-${formCode}) - ${tag}`, period.periodeLabel], [], ["Total baris", all.length]];
    if (bankStat) {
      ringkasan.push([], ["Nama bank peserta", `${bankStat.bernama} dari ${bankStat.total} sandi berhasil dinamai.`]);
      ringkasan.push(["Sumber nama", `bank pelapor sendiri + form Penempatan pada Bank Lain (KC0500). Total referensi: ${bankNameCount(bankMap)} bank.`]);
      if (bankStat.belum.length) {
        ringkasan.push(["Sandi tanpa nama", bankStat.belum.join(", ")]);
        ringkasan.push(["Catatan", "Sandi di atas dikosongkan, BUKAN ditebak. Nama bank peserta memang tidak tersedia di laporan bulanan ini."]);
      }
    }
    if (skipBranch) ringkasan.push([], ["Catatan", `Data > ${PER_BRANCH_MAX.toLocaleString("id-ID")} baris: sheet per-cabang dilewati. Gunakan filter kolom "Cabang" di sheet SEMUA CABANG.`]);
    if (chunkCount > 1) ringkasan.push([], ["Catatan", `Data ${all.length.toLocaleString("id-ID")} baris dipecah ke ${chunkCount} sheet (SEMUA CABANG 1..${chunkCount}), masing-masing maksimal ${SHEET_MAX.toLocaleString("id-ID")} baris. Ini batas teknis penulis file Excel, bukan data yang terpotong. Total baris di semua sheet = ${all.length.toLocaleString("id-ID")}.`]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ringkasan), "RINGKASAN");
    if (chunkCount === 1) {
      XLSX.utils.book_append_sheet(wb, sheet(XLSX, colsOrder, all), "SEMUA CABANG");
    } else {
      for (let i = 0; i < chunkCount; i++) {
        XLSX.utils.book_append_sheet(wb, sheet(XLSX, colsOrder, all.slice(i * SHEET_MAX, (i + 1) * SHEET_MAX)), `SEMUA CABANG ${i + 1}`);
      }
    }
    if (!skipBranch) for (const code of Object.keys(perBranch).sort()) XLSX.utils.book_append_sheet(wb, sheet(XLSX, colsOrder, perBranch[code]), `Cabang ${code}`);

    const summary = { jumlah_baris: all.length };
    if (bankStat) {
      summary.sandi_bank_unik = bankStat.total;
      summary.sandi_bank_bernama = bankStat.bernama;
      summary.sandi_bank_belum_bernama = bankStat.belum.length;
    }
    for (const col of colsOrder) {
      if (["Baki Debet", "Nominal", "Jumlah"].some(k => col.includes(k))) {
        let tot = 0; for (const r of all) { const v = parseFloat(r[col]); if (!isNaN(v)) tot += v; }
        if (tot) summary[`total[${col}]`] = Math.round(tot * 100) / 100;
      }
    }
    return { filename: `${namePrefix}_${tag}_${period.periodeLabel}.xlsx`, data: XLSX.write(wb, { type: "array", bookType: "xlsx" }), summary };
  };
}
