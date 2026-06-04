// UI + orchestrasi. Pakai window.XLSX (xlsx-js-style) & window.JSZip yang di-load lokal.
import { detectPeriod } from "./period.js";
import { processPembiayaan } from "./pembiayaan.js";
import { processAbp, processPenempatan } from "./bankLain.js";
import { processNeraca } from "./neraca.js";
import { processKredit } from "./kredit.js";
import { processPenempatanKonven } from "./bankLainKonven.js";
import { processNeracaKonven } from "./neracaKonven.js";
import { makeFormProcessor } from "./formKonven.js";

// SHA-256 hex dari PIN (lowercase) yang valid. PIN disimpan sbg hash, bukan teks asli.
// ismaya = Suriyah, "nisa alya" = BMP, "devi oktaviani" = Artha Perwira. Case-insensitive.
const PIN_HASHES = [
  "51e88e41cab453ae0c42874b2b9fa6fb152f5a1ae9355c8055a11062143f5bc6", // ismaya
  "c5ae4918bde0b3807702f60783f91af07e0122aed69e10af2c5abc3a6b213a5f", // nisa alya
  "c87b6cbcdf3bf7cb284777a41c200a88be1364ce4864f4c35edee5c54349a87a", // devi oktaviani
];
const XLSX = window.XLSX, JSZip = window.JSZip;
const DASH_KEY = "lbbprs_dashboard_v1";

function esc(s) {
  return String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

async function sha256hex(str) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(str));
  return [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, "0")).join("");
}

// ---------- PIN gate ----------
const gate = document.getElementById("gate");
const appEl = document.getElementById("app");
function unlock() { gate.style.display = "none"; appEl.style.display = "block"; }
if (sessionStorage.getItem("lbbprs_ok") === "1") unlock();

async function tryPin() {
  const pin = document.getElementById("pin").value.trim().toLowerCase();
  const h = await sha256hex(pin);
  if (PIN_HASHES.includes(h)) {
    sessionStorage.setItem("lbbprs_ok", "1");
    unlock();
  } else {
    document.getElementById("pinErr").textContent = "PIN salah.";
    document.getElementById("pin").value = "";
  }
}
document.getElementById("pinBtn").addEventListener("click", tryPin);
document.getElementById("pin").addEventListener("keydown", e => { if (e.key === "Enter") tryPin(); });

// ---------- Upload UI ----------
const drop = document.getElementById("drop");
const fileInput = document.getElementById("file");
const goBtn = document.getElementById("go");
const statusEl = document.getElementById("status");
const useTrend = document.getElementById("useTrend");
const trendFile = document.getElementById("trendFile");
let zipFile = null;

function setZip(f) {
  zipFile = f;
  document.getElementById("fname").textContent = f ? f.name : "";
  goBtn.disabled = !f;
}
drop.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", e => { if (e.target.files[0]) setZip(e.target.files[0]); });
["dragover", "dragenter"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("drag"); }));
["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("drag"); }));
drop.addEventListener("drop", e => { if (e.dataTransfer.files[0]) setZip(e.dataTransfer.files[0]); });
useTrend.addEventListener("change", () => { if (useTrend.checked) trendFile.click(); else { trendFile.value = ""; document.getElementById("trendName").textContent = ""; } });
trendFile.addEventListener("change", e => { document.getElementById("trendName").textContent = e.target.files[0] ? e.target.files[0].name : ""; });

async function readFileBytes(f) { return new Uint8Array(await f.arrayBuffer()); }

async function unzip(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const files = {};
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir) continue;
    files[name] = await zip.files[name].async("uint8array");
  }
  return files;
}

function download(filename, u8) {
  const blob = new Blob([u8], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.className = "dl"; a.textContent = filename;
  document.getElementById("links").appendChild(a);
  // auto-trigger juga
  const tmp = document.createElement("a"); tmp.href = url; tmp.download = filename; tmp.click();
}

goBtn.addEventListener("click", async () => {
  if (!zipFile) return;
  goBtn.disabled = true;
  statusEl.textContent = "Lagi diproses, tunggu sebentar...";
  document.getElementById("links").innerHTML = "";
  document.getElementById("resWarn").innerHTML = "";
  try {
    const period = detectPeriod(zipFile.name);
    const files = await unzip(await readFileBytes(zipFile));
    let priorTrend = null;
    if (useTrend.checked && trendFile.files[0]) priorTrend = await readFileBytes(trendFile.files[0]);

    const warnings = [];
    const summary = {};
    const runners = period.reportType === "konven" ? [
      ["Kredit", () => processKredit(files, period, XLSX)],
      ["Penempatan", () => processPenempatanKonven(files, period, XLSX)],
      ["Agunan", () => makeFormProcessor("0601", "Daftar Agunan", "GABUNGAN_AGUNAN")(files, period, XLSX)],
      ["Hapus Buku", () => makeFormProcessor("1500", "Aset Produktif Dihapus Buku", "GABUNGAN_HAPUS_BUKU")(files, period, XLSX)],
      ["Liabilitas Lainnya", () => makeFormProcessor("1400", "Rincian Liabilitas Lainnya", "GABUNGAN_LIABILITAS_LAINNYA")(files, period, XLSX)],
      ["Sindikasi", () => makeFormProcessor("0602", "Daftar Kredit Sindikasi", "GABUNGAN_SINDIKASI")(files, period, XLSX)],
      ["Tabungan", () => makeFormProcessor("1100", "Daftar Tabungan", "GABUNGAN_TABUNGAN")(files, period, XLSX)],
      ["Deposito", () => makeFormProcessor("1200", "Daftar Deposito", "GABUNGAN_DEPOSITO")(files, period, XLSX)],
      ["Neraca tren", () => processNeracaKonven(files, period, XLSX, priorTrend)],
    ] : [
      ["Pembiayaan", () => processPembiayaan(files, period, XLSX)],
      ["ABP", () => processAbp(files, period, XLSX)],
      ["Penempatan", () => processPenempatan(files, period, XLSX)],
      ["Neraca tren", () => processNeraca(files, period, XLSX, priorTrend)],
    ];
    document.getElementById("resPeriode").textContent = `${period.periodeLabel} (${period.jenisLapor})`;
    document.getElementById("resultCard").style.display = "block";

    for (const [label, fn] of runners) {
      try {
        const res = fn();
        if (!res) continue;  // form ga ada di ZIP -> skip
        download(res.filename, res.data);
        Object.assign(summary, Object.fromEntries(Object.entries(res.summary).map(([k, v]) => [`${label}: ${k}`, v])));
      } catch (err) {
        warnings.push(`${label} gagal: ${err.message}`);
      }
    }
    // #7 Reconcile sindikasi: kredit (Jenis 01) vs form 0602
    if (period.reportType === "konven") {
      const krSind = summary["Kredit: sindikasi_baki_debet"];
      const f0602 = summary["Sindikasi: total[Bagian Pendanaan Baki Debet]"];
      if (krSind !== undefined && f0602 !== undefined) {
        const selisih = Math.round((krSind - f0602) * 100) / 100;
        const st = Math.abs(selisih) < 1 ? "COCOK" : `SELISIH ${selisih.toLocaleString("id-ID")}`;
        warnings.push(`Rekon sindikasi: kredit Jenis-01 Rp ${krSind.toLocaleString("id-ID")} vs form 0602 Rp ${f0602.toLocaleString("id-ID")} -> ${st}`);
      }
    }
    document.getElementById("resWarn").innerHTML = warnings.map(w => `<div>${esc(w)}</div>`).join("");
    document.getElementById("resSummary").innerHTML = Object.entries(summary)
      .map(([k, v]) => `<div>${esc(k)}: <b>${typeof v === "number" ? v.toLocaleString("id-ID") : esc(v)}</b></div>`).join("");

    saveDash(period, summary);
    renderDash();
    statusEl.textContent = "Selesai. File ke-download otomatis.";
  } catch (err) {
    statusEl.textContent = "";
    document.getElementById("resultCard").style.display = "block";
    document.getElementById("resWarn").innerHTML = `<div>Gagal: ${esc(err.message)}</div>`;
  } finally {
    goBtn.disabled = false;
  }
});

// ---------- Dashboard (localStorage) ----------
function loadDash() { try { return JSON.parse(localStorage.getItem(DASH_KEY) || "[]"); } catch { return []; } }
function saveDash(period, summary) {
  const data = loadDash().filter(e => e.periode !== period.periodeLabel);
  data.push({
    periode: period.periodeLabel, jenis: period.jenisLapor,
    tanggal: new Date().toISOString().slice(0, 10),
    rekening: summary["Pembiayaan: jumlah_rekening"] ?? "-",
    npf: summary["Pembiayaan: jumlah_npf"] ?? "-",
  });
  data.sort((a, b) => a.periode.localeCompare(b.periode));
  localStorage.setItem(DASH_KEY, JSON.stringify(data));
}
function renderDash() {
  const data = loadDash();
  const el = document.getElementById("dash");
  if (!data.length) { el.innerHTML = `<p style="color:#888;">Belum ada. Upload ZIP pertama lu di atas.</p>`; return; }
  el.innerHTML = `<table><tr><th>Periode</th><th>Jenis</th><th>Diproses</th>
    <th class="num">Rekening Pembiayaan</th><th class="num">NPF</th></tr>` +
    data.map(e => `<tr><td><span class="pill">${e.periode}</span></td><td>${e.jenis}</td><td>${e.tanggal}</td>
      <td class="num">${typeof e.rekening === "number" ? e.rekening.toLocaleString("id-ID") : e.rekening}</td>
      <td class="num">${e.npf}</td></tr>`).join("") + `</table>`;
}
document.getElementById("clearDash").addEventListener("click", () => {
  if (confirm("Hapus semua riwayat di browser ini?")) { localStorage.removeItem(DASH_KEY); renderDash(); }
});
renderDash();
