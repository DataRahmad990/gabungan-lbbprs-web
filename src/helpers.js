// Helper kalkulasi (port dari core/pembiayaan.py)

// Parse tanggal dari nilai sel xls. Sumber sering simpan tanggal sbg teks "DD-MM-YYYY".
export function parseDate(val) {
  if (val instanceof Date && !isNaN(val)) return val;
  if (typeof val === "number" && isFinite(val)) {
    // serial Excel (basis 1899-12-30)
    const ms = Math.round((val - 25569) * 86400000);
    const d = new Date(ms);
    return isNaN(d) ? null : d;
  }
  if (typeof val === "string") {
    const s = val.trim();
    let m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/); // DD-MM-YYYY
    if (m) return new Date(Date.UTC(+m[3], +m[2]-1, +m[1]));
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/); // YYYY-MM-DD
    if (m) return new Date(Date.UTC(+m[1], +m[2]-1, +m[3]));
  }
  return null;
}

// Jangka waktu pinjaman antara dua tanggal (mulai -> jatuh tempo / akad awal -> akhir).
// Hitung selisih bulan kalender (completed months); kembalikan {bulan, teks "X tahun Y bulan"}.
// bulan=null & teks="" kalau salah satu tanggal tak terbaca atau akhir < mulai.
export function jangkaWaktu(startVal, endVal) {
  const s = parseDate(startVal), e = parseDate(endVal);
  if (!s || !e || e.getTime() < s.getTime()) return { bulan: null, teks: "" };
  let m = (e.getFullYear() - s.getFullYear()) * 12 + (e.getMonth() - s.getMonth());
  if (e.getDate() < s.getDate()) m -= 1;
  if (m < 0) m = 0;
  return { bulan: m, teks: `${Math.floor(m / 12)} tahun ${m % 12} bulan` };
}

export function calcOverdueDays(jtDate, posisi) {
  if (jtDate && jtDate.getTime() < posisi.getTime()) {
    return Math.round((posisi.getTime() - jtDate.getTime()) / 86400000);
  }
  return 0;
}

export function kolLabel(kolVal) {
  const labels = {1:"1-Lancar",2:"2-DPK",3:"3-Kurang Lancar",4:"4-Diragukan",5:"5-Macet"};
  const k = parseInt(parseFloat(kolVal), 10);
  return labels[k] || String(kolVal);
}

export function expectedMinKol(od) {
  if (od <= 30) return 1;
  if (od <= 90) return 2;
  if (od <= 180) return 3;
  if (od <= 360) return 4;
  return 5;
}

export function detectAnomaly(kolActual, od, jtDate) {
  const kol = parseInt(parseFloat(kolActual), 10);
  if (isNaN(kol)) return "";
  const flags = [];
  const minKol = expectedMinKol(od);
  if (kol < minKol) flags.push(`KOL TERLALU TINGGI (harusnya min Kol ${minKol})`);
  if (od > 360 && kol < 5) flags.push(`HARUSNYA MACET (OD ${od} hari)`);
  if (od > 1000 && kol === 1) flags.push("ANOMALI KRITIS: Lancar tapi OD >1000 hari");
  return flags.join(" | ");
}

// cell(aoa, r, c) aman terhadap baris pendek
export function cell(aoa, r, c) {
  const row = aoa[r];
  if (!row) return "";
  const v = row[c];
  return v === undefined || v === null ? "" : v;
}

export function findDataStartRow(aoa) {
  const kws = ["FORM","DAFTAR","APLIKASI","NAMA","POSISI","SANDI","ID PIHAK"];
  const n = Math.min(25, aoa.length);
  for (let r = 10; r < n; r++) {
    const v = cell(aoa, r, 3);
    const s = String(v).trim();
    if (s && s !== "ID Pihak Lawan" && s !== "JUMLAH") {
      const up = s.toUpperCase();
      if (!kws.some(kw => up.includes(kw))) return r;
    }
  }
  return 18;
}

// Format Date -> string buat ditulis ke xlsx (samain tampilan dgn Python datetime)
export function toCellDate(v) {
  return v instanceof Date && !isNaN(v) ? v : v;
}
