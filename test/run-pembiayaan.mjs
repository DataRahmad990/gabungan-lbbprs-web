import fs from "fs"; import os from "os"; import path from "path";
import XLSX from "xlsx-js-style";
import JSZip from "jszip";
import { detectPeriod } from "../src/period.js";
import { processPembiayaan } from "../src/pembiayaan.js";

const ZIP = path.join(os.homedir(), "Downloads/LBBPRS-LBBPRS01-K-LBBPRSKI-20260228-620086-corp_ismaya.novita.zip");
const GOLD = path.resolve("../GABUNGAN_PEMBIAYAAN_BPRS_SURIYAH_FEB2026.xlsx");

async function unzip(buf) {
  const zip = await JSZip.loadAsync(buf);
  const files = {};
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir) continue;
    files[name] = await zip.files[name].async("uint8array");
  }
  return files;
}

const period = detectPeriod(path.basename(ZIP));
const files = await unzip(fs.readFileSync(ZIP));
const res = processPembiayaan(files, period, XLSX);
console.log("filename:", res.filename);
console.log("summary:", res.summary);

// baca hasil
const wbNew = XLSX.read(res.data, { type: "array" });
function rows(wb, sheet) {
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: null });
  return aoa;
}
const EXPECT = { "SEMUA AKAD": 3129, "Murabahah": 1607, "Multijasa": 1357, "Musyarakah": 142, "Ijarah": 11, "Mudarabah": 9 };
let ok = true;
for (const [s, n] of Object.entries(EXPECT)) {
  const got = rows(wbNew, s).length - 1;
  const pass = got === n;
  if (!pass) ok = false;
  console.log(`  ${pass?"OK":"XX"} ${s}: ${got} (expect ${n})`);
}

// banding enriched cols vs golden SEMUA AKAD
const wbGold = XLSX.read(fs.readFileSync(GOLD), { type: "array" });
const aN = rows(wbNew, "SEMUA AKAD"), aG = rows(wbGold, "SEMUA AKAD");
const hN = Object.fromEntries(aN[0].map((v,i)=>[v,i])), hG = Object.fromEntries(aG[0].map((v,i)=>[v,i]));
const key = (row,h)=> `${row[h["Cabang"]]}|${row[h["Nomor Rekening"]]}|${row[h["Akad"]]}`;
const goldMap = new Map(); for (let i=1;i<aG.length;i++) goldMap.set(key(aG[i],hG), aG[i]);
const norm = v => (v==null?"":String(v)).replace(/—/g,",").replace(/–/g,"-").replace(/ ,/g,",");
const COLS = ["Lokasi Penggunaan","Jenis Piutang","HT Flag","NPF Flag","Lama Macet (Hari)","Keterangan Pengikatan"];
let checked=0, mism=[];
for (let i=1;i<aN.length;i++){
  const g = goldMap.get(key(aN[i],hN)); if(!g) continue; checked++;
  for(const c of COLS){ if(norm(aN[i][hN[c]])!==norm(g[hG[c]])){ if(mism.length<8) mism.push([c, aN[i][hN[c]], g[hG[c]]]); } }
}
console.log(`enriched check: ${checked} baris match, mismatch sample:`, mism.length, mism.slice(0,4));
console.log(ok && checked>3000 && mism.length===0 ? "\n==== ALL PASS ====" : "\n==== ADA YANG BEDA ====");
