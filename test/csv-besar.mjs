// Verifikasi output CSV form raksasa: jumlah baris & total nominal harus sama
// dengan hasil baca langsung dari file mentah.
import fs from "fs"; import path from "path";
import XLSX from "xlsx-js-style";
import JSZip from "jszip";
import { detectPeriod } from "../src/period.js";
import { makeFormProcessor } from "../src/formKonven.js";

const zip = await JSZip.loadAsync(fs.readFileSync(process.argv[2]));
const files = {};
for (const n of Object.keys(zip.files)) { if (zip.files[n].dir) continue; files[n] = await zip.files[n].async("uint8array"); }
const sample = Object.keys(files).find(n => n.split("/").pop().startsWith("LBBPRK-0600-"));
const period = detectPeriod(path.basename(sample));
const res = makeFormProcessor("1100", "Daftar Tabungan", "GABUNGAN_TABUNGAN")(files, period, XLSX);
let fail = 0;
console.log("file        :", res.filename);
console.log("ukuran      :", (res.data.byteLength / 1048576).toFixed(1), "MB");
const text = new TextDecoder().decode(res.data).replace(/^﻿/, "");
const lines = text.split("\r\n").filter(l => l !== "");
const hdr = lines[0].split(";");
const body = lines.slice(1);
console.log("baris CSV   :", body.length.toLocaleString("id-ID"), "(summary:", res.summary.jumlah_baris.toLocaleString("id-ID") + ")");
if (body.length !== res.summary.jumlah_baris) { console.log("GAGAL: jumlah baris CSV != summary"); fail++; }

const iJml = hdr.findIndex(h => h.trim() === "Jumlah");
const iHub = hdr.findIndex(h => h.includes("Hubungan dengan Bank"));
console.log(`kolom Jumlah: ${iJml}  kolom Hubungan: ${iHub}`);
let totCsv = 0; const perHub = {};
for (const l of body) {
  const f = l.split(";");
  const v = parseFloat(f[iJml]); if (!isNaN(v)) totCsv += v;
  const h = (f[iHub] || "").trim(); perHub[h] = (perHub[h] || 0) + (isNaN(v) ? 0 : v);
}
console.log("total Jumlah:", Math.round(totCsv).toLocaleString("id-ID"));
console.log("per hubungan:");
for (const [k, v] of Object.entries(perHub).sort()) console.log(`   ${k || "(kosong)"} = ${Math.round(v).toLocaleString("id-ID")}`);
if (Math.round(totCsv) === 0) { console.log("GAGAL: total nol"); fail++; }
console.log(fail === 0 ? "\nCSV LULUS" : `\n${fail} GAGAL`);
process.exit(fail === 0 ? 0 : 1);
