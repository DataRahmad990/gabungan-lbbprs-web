// Deteksi nama bank + cabang dari files map (port dari core/bank.py). Bank-agnostic.
const PREFIXES = [
  "PT Bank Perekonomian Rakyat Syariah",
  "PT Bank Pembiayaan Rakyat Syariah",
  "PT BPRS", "BPRS", "PT BPR Syariah", "PT Bank",
];

function sanitizeTag(nama) {
  let s = nama || "";
  for (const pre of PREFIXES) {
    if (s.trim().startsWith(pre)) { s = s.trim().slice(pre.length); break; }
  }
  s = s.trim();
  const tag = s.replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "").toUpperCase();
  return tag || (nama || "BANK").replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "").toUpperCase();
}

export function detectBank(files, XLSX) {
  for (const token of ["GB0200", "GB0300"]) {
    const name = Object.keys(files).find(n => n.split("/").pop().includes(token) && n.endsWith(".xls"));
    if (!name) continue;
    try {
      const wb = XLSX.read(files[name], { type: "array" });
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null });
      for (let r = 0; r < Math.min(20, aoa.length); r++) {
        const row = (aoa[r] || []).map(v => (v == null ? "" : String(v).trim()));
        if (row.some(v => v.includes("Nama Lembaga"))) {
          const vals = row.filter(v => v && v !== ":" && !v.includes("Nama Lembaga"));
          if (vals.length) { const nama = vals[vals.length - 1]; return { nama, tag: sanitizeTag(nama) }; }
        }
      }
    } catch (e) { /* lanjut */ }
  }
  return { nama: "BANK", tag: "BANK" };
}

// Kode cabang (3 digit) yang beneran ada utk satu form prefix. Dinamis (2/6/N).
export function discoverBranches(files, formPrefix) {
  const codes = new Set();
  const re = /-(\d{3})(?:_part\d+)?\.xls$/;
  for (const n of Object.keys(files)) {
    const b = n.split("/").pop();
    if (b.startsWith(formPrefix)) {
      const m = b.match(re);
      if (m) codes.add(m[1]);
    }
  }
  return [...codes].sort();
}
