// Name lookup dari Form 0016 (Daftar Pihak Lawan). Port dari core/pihak_lawan.py.
// FIX: baca SEMUA file 0016 (part1 + part2; form ini dipecah karena > 65.536 baris
// melebihi limit .xls), dan cocokkan lewat DUA kunci: ID Pihak Lawan (kol 2) DAN
// Nomor Identitas/NIK (kol 6). Sebagian form (mis. kredit) memakai NIK di kolom ID.
import * as H from "./helpers.js";

function normId(v) {
  let s = String(v == null ? "" : v).trim();
  if (s.endsWith(".0")) s = s.slice(0, -2);
  return s;
}

export function buildNameMap(files, XLSX) {
  const names = Object.keys(files).filter(n => {
    const b = n.split("/").pop();
    return b.includes("LBBPRK-0016") && b.endsWith(".xls");
  });
  const byId = {}, byNik = {};
  for (const name of names) {
    const wb = XLSX.read(files[name], { type: "array", raw: true });
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
    for (let r = 14; r < aoa.length; r++) {
      const nm = String(H.cell(aoa, r, 11) || "").trim();
      if (!nm) continue;
      const id = normId(H.cell(aoa, r, 2));
      const nik = normId(H.cell(aoa, r, 6));
      if (id && !(id in byId)) byId[id] = nm;
      if (nik && !(nik in byNik)) byNik[nik] = nm;
    }
  }
  return { byId, byNik };
}

// Cari nama pakai ID Pihak Lawan dulu, kalau kosong fallback ke Nomor Identitas/NIK.
// Backward-compatible: menerima map lama berbentuk {id: nama}.
export function lookupName(map, id, nik) {
  if (!map) return "";
  const m = map.byId ? map : { byId: map, byNik: {} };
  return m.byId[normId(id)] || (nik != null && nik !== "" ? m.byNik[normId(nik)] : "") || "";
}

// Jumlah nama termuat (buat cek apakah kolom Nama layak ditampilkan).
export function nameCount(map) {
  if (!map) return 0;
  return map.byId ? Object.keys(map.byId).length : Object.keys(map).length;
}
