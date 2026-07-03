/* 轻转 LiteConvert 端到端测试:无头 Chromium 中跑通全部转换 */
import { spawn } from 'child_process';

// 优先用本地安装的 playwright,其次尝试全局路径
let chromium;
try {
  ({ chromium } = await import('playwright'));
} catch (e) {
  ({ chromium } = await import('/home/claude/.npm-global/lib/node_modules/playwright/index.mjs'));
}
import fs from 'fs';
import path from 'path';

const ROOT = path.resolve(new URL('.', import.meta.url).pathname, '..');
const OUT = path.join(ROOT, 'test', 'out');
fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(OUT, { recursive: true });

const server = spawn('python3', [path.join(ROOT, 'serve.py'), '8973'], { stdio: 'ignore' });
// 轮询等待服务就绪(最长 20 秒)
{
  let up = false;
  for (let i = 0; i < 100 && !up; i++) {
    try {
      const res = await fetch('http://127.0.0.1:8973/', { signal: AbortSignal.timeout(500) });
      up = res.ok;
    } catch (e) { await new Promise(r => setTimeout(r, 200)); }
  }
  if (!up) { console.error('本地服务启动失败'); process.exit(1); }
}

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
page.on('console', m => { if (m.type() === 'error') console.log('[page error]', m.text()); });
page.on('pageerror', e => console.log('[pageerror]', e.message));

await page.goto('http://127.0.0.1:8973/', { waitUntil: 'networkidle' });

// 注入测试助手
await page.evaluate(() => {
  window.__test = async (b64, name, mime, action) => {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const file = new File([bytes], name, { type: mime });
    const C = window.LiteConvert;
    let res;
    switch (action) {
      case 'img2jpeg': res = await C.imageToImage(file, 'jpeg'); break;
      case 'img2png': res = await C.imageToImage(file, 'png'); break;
      case 'img2webp': res = await C.imageToImage(file, 'webp'); break;
      case 'img2bmp': res = await C.imageToImage(file, 'bmp'); break;
      case 'img2tiff': res = await C.imageToImage(file, 'tiff'); break;
      case 'img2gif': res = await C.imageToImage(file, 'gif'); break;
      case 'img2pdf': res = { blob: await C.imagesToPdf([file]), ext: 'pdf' }; break;
      case 'pdf2png': res = await C.pdfToImages(file, 'png'); break;
      case 'pdf2jpeg': res = await C.pdfToImages(file, 'jpeg'); break;
      case 'pdf2docx': res = await C.pdfToDocx(file); break;
      case 'pdf2pptx': res = await C.pdfToPptx(file); break;
      case 'pdf2txt': res = await C.pdfToTxt(file); break;
      case 'docx2pdf': res = await C.docxToPdf(file); break;
      case 'docx2txt': res = await C.docxToTxt(file); break;
      case 'pptx2pdf': res = await C.pptxToPdf(file); break;
      case 'pptx2png': res = await C.pptxToImages(file); break;
      default: throw new Error('unknown action ' + action);
    }
    const buf = new Uint8Array(await res.blob.arrayBuffer());
    let s = '';
    const CH = 0x8000;
    for (let i = 0; i < buf.length; i += CH) s += String.fromCharCode.apply(null, buf.subarray(i, i + CH));
    return { b64: btoa(s), ext: res.ext, note: res.note || '' };
  };
});

const read = f => fs.readFileSync(path.join(ROOT, 'test', f)).toString('base64');
const cases = [
  ['sample.png', 'image/png', 'img2jpeg', 'out_png.jpg'],
  ['sample.png', 'image/png', 'img2webp', 'out_png.webp'],
  ['sample.png', 'image/png', 'img2bmp', 'out_png.bmp'],
  ['sample.png', 'image/png', 'img2tiff', 'out_png.tiff'],
  ['sample.png', 'image/png', 'img2gif', 'out_png.gif'],
  ['sample.jpg', 'image/jpeg', 'img2png', 'out_jpg.png'],
  ['sample_src.webp', 'image/webp', 'img2jpeg', 'out_webp.jpg'],
  ['sample.png', 'image/png', 'img2pdf', 'out_img.pdf'],
  ['sample.pdf', 'application/pdf', 'pdf2png', 'out_pdf_png.zip'],
  ['sample.pdf', 'application/pdf', 'pdf2jpeg', 'out_pdf_jpg.zip'],
  ['sample.pdf', 'application/pdf', 'pdf2docx', 'out_pdf.docx'],
  ['sample.pdf', 'application/pdf', 'pdf2pptx', 'out_pdf.pptx'],
  ['sample.pdf', 'application/pdf', 'pdf2txt', 'out_pdf.txt'],
  ['sample.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'docx2pdf', 'out_docx.pdf'],
  ['sample.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'docx2txt', 'out_docx.txt'],
  ['sample.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'pptx2pdf', 'out_pptx.pdf'],
  ['sample.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'pptx2png', 'out_pptx_png.zip'],
];

let pass = 0, fail = 0;
for (const [src, mime, action, outName] of cases) {
  try {
    const r = await page.evaluate(
      ([b64, name, mime, action]) => window.__test(b64, name, mime, action),
      [read(src), src, mime, action]
    );
    const outPath = path.join(OUT, outName);
    fs.writeFileSync(outPath, Buffer.from(r.b64, 'base64'));
    const size = fs.statSync(outPath).size;
    console.log(`PASS ${action.padEnd(10)} ${src} -> ${outName} (${(size / 1024).toFixed(1)} KB) ${r.note}`);
    pass++;
  } catch (e) {
    console.log(`FAIL ${action.padEnd(10)} ${src}: ${e.message.split('\n')[0]}`);
    fail++;
  }
}

// UI 截图(浅色 + 深色)
await page.screenshot({ path: path.join(OUT, 'ui_light.png'), fullPage: true });
await page.emulateMedia({ colorScheme: 'dark' });
await page.waitForTimeout(500); // 等待背景过渡动画结束
await page.screenshot({ path: path.join(OUT, 'ui_dark.png'), fullPage: true });
await page.emulateMedia({ colorScheme: 'light' });

// 带文件的工作台截图
const fileBuf = fs.readFileSync(path.join(ROOT, 'test', 'sample.docx'));
await page.setInputFiles('#fileInput', [
  { name: '季度报告.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: fileBuf },
  { name: '产品照片.png', mimeType: 'image/png', buffer: fs.readFileSync(path.join(ROOT, 'test', 'sample.png')) },
  { name: '合同扫描.pdf', mimeType: 'application/pdf', buffer: fs.readFileSync(path.join(ROOT, 'test', 'sample.pdf')) },
]);
await page.waitForTimeout(400);
await page.screenshot({ path: path.join(OUT, 'ui_workbench.png'), fullPage: false });

await browser.close();
server.kill();
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
