import http from "http"; import fs from "fs"; import os from "os"; import path from "path";
import puppeteer from "puppeteer-core";

const ROOT = path.resolve(".");
const MIME = { ".html":"text/html",".js":"text/javascript",".json":"application/json",".css":"text/css" };
const server = http.createServer((req,res)=>{
  let p = decodeURIComponent(req.url.split("?")[0]); if(p==="/")p="/index.html";
  const fp = path.join(ROOT, p);
  if(!fp.startsWith(ROOT)||!fs.existsSync(fp)){res.writeHead(404);return res.end("nf");}
  res.writeHead(200,{"Content-Type":MIME[path.extname(fp)]||"application/octet-stream"});
  fs.createReadStream(fp).pipe(res);
});
await new Promise(r=>server.listen(0,r));
const port = server.address().port;
const CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

const browser = await puppeteer.launch({ executablePath:CHROME, headless:"new", args:["--no-sandbox"] });
const page = await browser.newPage();
const errors=[];
page.on("pageerror",e=>errors.push("PAGEERR: "+e.message));
page.on("console",m=>{ if(m.type()==="error") errors.push("CONSOLE: "+m.text()); });

// intercept downloads
const dlDir = fs.mkdtempSync(path.join(os.tmpdir(),"dl-"));
const client = await page.target().createCDPSession();
await client.send("Page.setDownloadBehavior",{behavior:"allow",downloadPath:dlDir});

await page.goto(`http://127.0.0.1:${port}/index.html`,{waitUntil:"networkidle0"});

// 1) PIN gate ada
const gateVisible = await page.evaluate(()=>getComputedStyle(document.getElementById("gate")).display!=="none");
console.log("gate tampil:", gateVisible);

// 2) PIN salah ditolak
await page.type("#pin","9999"); await page.click("#pinBtn");
await new Promise(r=>setTimeout(r,200));
const errMsg = await page.$eval("#pinErr",e=>e.textContent);
console.log("PIN salah -> pesan:", JSON.stringify(errMsg));

// 3) PIN bener (1234) -> app kebuka
await page.evaluate(()=>document.getElementById("pin").value="");
await page.type("#pin","1234"); await page.click("#pinBtn");
await new Promise(r=>setTimeout(r,300));
const appVisible = await page.evaluate(()=>getComputedStyle(document.getElementById("app")).display!=="none");
console.log("app kebuka stlh PIN bener:", appVisible);

// 4) upload ZIP FEB
const ZIP = path.join(os.homedir(),"Downloads/LBBPRS-LBBPRS01-K-LBBPRSKI-20260228-620086-corp_ismaya.novita.zip");
const input = await page.$("#file");
await input.uploadFile(ZIP);
await page.click("#go");
// tunggu sampe ada link hasil / status selesai
await page.waitForFunction(()=>document.getElementById("status").textContent.includes("Selesai")||document.getElementById("resWarn").textContent.includes("Gagal"),{timeout:60000});
const status = await page.$eval("#status",e=>e.textContent);
const linkCount = await page.$$eval("#links a",as=>as.map(a=>a.textContent));
const summaryText = await page.$eval("#resSummary",e=>e.textContent);
console.log("status:", status);
console.log("download links:", linkCount);
console.log("summary (potong):", summaryText.slice(0,160));

await new Promise(r=>setTimeout(r,1500)); // tunggu file ke-download
const dl = fs.readdirSync(dlDir).filter(f=>!f.endsWith(".crdownload"));
console.log("file ke-download:", dl);

console.log("\nERRORS:", errors.length? errors : "none");
await browser.close(); server.close();
