import fs from "fs"; import os from "os"; import path from "path";
import XLSX from "xlsx-js-style";
import JSZip from "jszip";
import { detectPeriod } from "../src/period.js";
import { processAbp, processPenempatan } from "../src/bankLain.js";
import { processNeraca } from "../src/neraca.js";

const FEB = path.join(os.homedir(), "Downloads/LBBPRS-LBBPRS01-K-LBBPRSKI-20260228-620086-corp_ismaya.novita.zip");
const APR = path.join(os.homedir(), "Downloads/LBBPRS-LBBPRS01-R-20260430-620086-corp_ismaya.novita.zip");
async function unzip(p){ const z=await JSZip.loadAsync(fs.readFileSync(p)); const f={}; for(const n of Object.keys(z.files)){ if(z.files[n].dir)continue; f[n]=await z.files[n].async("uint8array"); } return f; }
const rows=(wb,s)=>XLSX.utils.sheet_to_json(wb.Sheets[s],{header:1,defval:null});

const pFeb=detectPeriod(path.basename(FEB)); const feb=await unzip(FEB);

const abp=processAbp(feb,pFeb,XLSX);
console.log("ABP:", abp.summary, "| SEMUA CABANG rows:", rows(XLSX.read(abp.data,{type:"array"}),"SEMUA CABANG").length-1, "(expect 111)");

const pen=processPenempatan(feb,pFeb,XLSX);
console.log("PENEMPATAN:", pen.summary, "| SEMUA CABANG rows:", rows(XLSX.read(pen.data,{type:"array"}),"SEMUA CABANG").length-1, "(expect 93)");

// neraca FEB
const n1=processNeraca(feb,pFeb,XLSX,null);
console.log("NERACA FEB:", n1.summary);
let wbN=XLSX.read(n1.data,{type:"array"});
console.log("  Neraca header:", rows(wbN,"Neraca")[0]);

// neraca APR pakai tren FEB
if (fs.existsSync(APR)) {
  const pApr=detectPeriod(path.basename(APR)); const apr=await unzip(APR);
  const n2=processNeraca(apr,pApr,XLSX,n1.data);
  wbN=XLSX.read(n2.data,{type:"array"});
  const h=rows(wbN,"Neraca")[0];
  console.log("NERACA +APR header:", h);
  const sample=rows(wbN,"Neraca").slice(1,4).map(r=>[r[0],r[1],r[h.indexOf("FEB2026")],r[h.indexOf("APR2026")]]);
  console.log("  tren sample:", JSON.stringify(sample));
}
const ok = abp.summary.jumlah_rekening===111 && pen.summary.jumlah_rekening===93 && n1.summary.neraca_jml_pos===36;
console.log(ok? "\n==== ABP/PENEMPATAN/NERACA PASS ====" : "\n==== ADA BEDA ====");
