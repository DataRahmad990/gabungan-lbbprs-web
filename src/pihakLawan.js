// Name lookup dari Form 0016 (ID Pihak Lawan -> Nama). Port dari core/pihak_lawan.py.
import * as H from "./helpers.js";

function normId(v) {
  let s = String(v == null ? "" : v).trim();
  if (s.endsWith(".0")) s = s.slice(0, -2);
  return s;
}

export function buildNameMap(files, XLSX) {
  const name = Object.keys(files).find(n => n.split("/").pop().includes("LBBPRK-0016") && n.endsWith(".xls"));
  if (!name) return {};
  const wb = XLSX.read(files[name], { type: "array", raw: true });
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
  const map = {};
  for (let r = 14; r < aoa.length; r++) {
    const id = normId(H.cell(aoa, r, 2));
    const nm = String(H.cell(aoa, r, 11) || "").trim();
    if (id && nm) map[id] = nm;
  }
  return map;
}

export function lookupName(map, id) {
  return map[normId(id)] || "";
}
