// Test: SEMUA AKAD & tiap sheet akad harus 1 baris per rekening (tidak dobel),
// TANPA mengubah total Baki Debet.
import fs from "fs"; import path from "path";
import XLSX from "xlsx-js-style";
import JSZip from "jszip";
import { detectPeriod } from "../src/period.js";
import { processPembiayaan } from "../src/pembiayaan.js";

const ZIP = process.argv[2];
// Tidak ada angka bank yang di-hardcode (repo publik). Nilai acuan opsional
// lewat env: EXPECT_BAKI=... node test/dedup.mjs <zip>

async function unzip(buf) {
  const zip = await JSZip.loadAsync(buf); const files = {};
  for (const n of Object.keys(zip.files)) { if (zip.files[n].dir) continue; files[n] = await zip.files[n].async("uint8array"); }
  return files;
}
const files = await unzip(fs.readFileSync(ZIP));
const res = processPembiayaan(files, detectPeriod(path.basename(ZIP)), XLSX);
const wb = XLSX.read(res.data, { type: "array" });

let fail = 0;
const num = v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };

for (const sheetName of wb.SheetNames) {
  if (["RINGKASAN", "RINCIAN AGUNAN", "REF PENGIKATAN", "REF AGUNAN"].includes(sheetName)) continue;
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null });
  const hdr = aoa[0] || [];
  const iRek = hdr.indexOf("Nomor Rekening");
  if (iRek < 0) continue;
  const rows = aoa.slice(1).filter(r => r && String(r[iRek] ?? "").trim());
  const m = {};
  for (const r of rows) { const k = String(r[iRek]).trim(); (m[k] ||= []).push(r); }
  const dup = Object.entries(m).filter(([, v]) => v.length > 1);
  const extra = dup.reduce((s, [, v]) => s + v.length - 1, 0);
  const status = extra === 0 ? "OK  " : "GAGAL";
  if (extra !== 0) fail++;
  console.log(`${status} ${sheetName.padEnd(22)} baris=${String(rows.length).padStart(5)} rek_unik=${String(Object.keys(m).length).padStart(5)} dobel=${String(extra).padStart(4)}`);
}

// RINCIAN AGUNAN sengaja 1 baris per agunan. Yang diuji: jumlahnya konsisten
// dengan kolom "Jumlah Agunan" di SEMUA AKAD, jadi tidak ada agunan yang hilang.
{
  const sa = XLSX.utils.sheet_to_json(wb.Sheets["SEMUA AKAD"], { header: 1, defval: null });
  const h = sa[0]; const iJml = h.indexOf("Jumlah Agunan");
  const klaim = sa.slice(1).reduce((n, r) => n + (r ? (parseInt(r[iJml], 10) || 0) : 0), 0);
  const ag = wb.Sheets["RINCIAN AGUNAN"]
    ? XLSX.utils.sheet_to_json(wb.Sheets["RINCIAN AGUNAN"], { header: 1, defval: null }).slice(1).filter(r => r && r.some(v => v !== null && v !== "")).length
    : 0;
  const ok = klaim === ag;
  if (!ok) fail++;
  console.log(`\n${ok ? "OK   " : "GAGAL"} Agunan: kolom 'Jumlah Agunan' total=${klaim}, baris RINCIAN AGUNAN=${ag} (harus sama)`);
}

// Invarian: total SEMUA AKAD harus sama dengan jumlah seluruh sheet per-akad.
// Ini menangkap baris hilang/dobel tanpa perlu menyimpan angka bank di repo.
const aoa = XLSX.utils.sheet_to_json(wb.Sheets["SEMUA AKAD"], { header: 1, defval: null });
const hdr = aoa[0]; const iBd = hdr.indexOf("Baki Debet");
const total = aoa.slice(1).reduce((s, r) => s + (r ? num(r[iBd]) : 0), 0);
let perAkad = 0;
for (const sn of wb.SheetNames) {
  if (["RINGKASAN", "SEMUA AKAD", "RINCIAN AGUNAN", "REF PENGIKATAN", "REF AGUNAN"].includes(sn)) continue;
  const a = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null });
  const i = (a[0] || []).indexOf("Baki Debet");
  if (i < 0) continue;
  perAkad += a.slice(1).reduce((s, r) => s + (r ? num(r[i]) : 0), 0);
}
const okSum = Math.round(total) === Math.round(perAkad);
console.log(`\n${okSum ? "OK   " : "GAGAL"} Total SEMUA AKAD == jumlah sheet per-akad (selisih ${Math.round(total - perAkad)})`);
if (!okSum) fail++;
if (process.env.EXPECT_BAKI) {
  const ok = Math.round(total) === Number(process.env.EXPECT_BAKI);
  console.log(`${ok ? "OK   " : "GAGAL"} Total cocok dengan EXPECT_BAKI dari env`);
  if (!ok) fail++;
}
console.log(fail === 0 ? "\nSEMUA TEST LULUS" : `\n${fail} TEST GAGAL`);
process.exit(fail === 0 ? 0 : 1);
