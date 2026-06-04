// Neraca + Laba Rugi tren BPR Konvensional (LBBPRK-0100 + LBBPRK-0200). Port dari neraca_konven.py.
import * as H from "./helpers.js";
import { detectBank } from "./bank.js";

const FORMS = [
  { prefix: "LBBPRK-0100-", sheet: "Neraca", sandi: 15, label: 5, nilai: 17 },
  { prefix: "LBBPRK-0200-", sheet: "Laba Rugi", sandi: 13, label: 4, nilai: 15 },
];

function findFile(files, prefix) {
  const name = Object.keys(files).find(n => n.split("/").pop().startsWith(prefix) && n.endsWith(".xls"));
  return name ? files[name] : null;
}

function parseForm(bytes, cfg, XLSX) {
  const wb = XLSX.read(bytes, { type: "array", raw: true });
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
  const out = [], seen = new Set();
  for (let r = 0; r < aoa.length; r++) {
    let sandi = String(H.cell(aoa, r, cfg.sandi)).trim();
    if (sandi.endsWith(".0")) sandi = sandi.slice(0, -2);
    if (!/^\d+$/.test(sandi)) continue;
    const nilai = H.cell(aoa, r, cfg.nilai);
    if (typeof nilai !== "number") continue;
    if (seen.has(sandi)) continue;
    seen.add(sandi);
    out.push([sandi, String(H.cell(aoa, r, cfg.label)).trim(), nilai]);
  }
  return out;
}

function upsertAoa(priorAoa, posRows, periodeLabel) {
  let header, body;
  if (priorAoa && priorAoa.length) { header = priorAoa[0].slice(); body = priorAoa.slice(1).map(r => r.slice()); }
  else { header = ["Sandi", "Pos"]; body = []; }
  const sandiRow = new Map();
  body.forEach((r, i) => sandiRow.set(String(r[0]), i));
  let pcol = header.indexOf(periodeLabel);
  if (pcol === -1) { pcol = header.length; header.push(periodeLabel); }
  for (const [sandi, label, nilai] of posRows) {
    let i = sandiRow.get(sandi);
    if (i === undefined) { const row = new Array(header.length).fill(null); row[0] = sandi; row[1] = label; body.push(row); i = body.length - 1; sandiRow.set(sandi, i); }
    while (body[i].length < header.length) body[i].push(null);
    body[i][pcol] = nilai;
  }
  return [header, ...body];
}

export function processNeracaKonven(files, period, XLSX, priorTrendBytes = null) {
  let prior = null;
  if (priorTrendBytes) { try { prior = XLSX.read(priorTrendBytes, { type: "array" }); } catch (e) { prior = null; } }
  const wb = XLSX.utils.book_new();
  const summary = {};
  let any = false;
  for (const cfg of FORMS) {
    const bytes = findFile(files, cfg.prefix);
    if (!bytes) continue;
    any = true;
    const pos = parseForm(bytes, cfg, XLSX);
    let priorAoa = null;
    if (prior && prior.Sheets[cfg.sheet]) priorAoa = XLSX.utils.sheet_to_json(prior.Sheets[cfg.sheet], { header: 1, defval: null });
    const aoa = upsertAoa(priorAoa, pos, period.periodeLabel);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 12 }, { wch: 55 }, ...aoa[0].slice(2).map(() => ({ wch: 18 }))];
    XLSX.utils.book_append_sheet(wb, ws, cfg.sheet);
    summary[`${cfg.sheet.toLowerCase().replace(/ /g, "_")}_jml_pos`] = pos.length;
  }
  if (!any) throw new Error("Nggak nemu LBBPRK-0100/0200 di ZIP.");
  summary.periode_ditambahkan = period.periodeLabel;
  const tag = detectBank(files, XLSX).tag;
  return { filename: `NERACA_TREN_${tag}.xlsx`, data: XLSX.write(wb, { type: "array", bookType: "xlsx" }), summary };
}
