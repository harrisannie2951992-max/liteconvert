/* 轻转 LiteConvert — 转换调度中心
 * 所有转换均在浏览器本地完成。MIT License
 */
(function (global) {
  'use strict';

  /* ---------- 懒加载 ---------- */

  let _pdfjs = null;
  async function getPdfjs() {
    if (_pdfjs) return _pdfjs;
    const mod = await import('../vendor/pdf.min.mjs');
    mod.GlobalWorkerOptions.workerSrc = new URL('vendor/pdf.worker.min.mjs', document.baseURI).href;
    _pdfjs = mod;
    return mod;
  }

  const _scripts = {};
  function loadScript(src, globalName) {
    if (_scripts[src]) return _scripts[src];
    _scripts[src] = new Promise((res, rej) => {
      if (globalName && global[globalName]) return res(global[globalName]);
      const s = document.createElement('script');
      s.src = src;
      s.onload = () => res(global[globalName]);
      s.onerror = () => rej(new Error('加载组件失败: ' + src));
      document.head.appendChild(s);
    });
    return _scripts[src];
  }
  const getDocxLib = () => loadScript('vendor/docx.iife.js', 'docx');
  const getPptxGen = () => loadScript('vendor/pptxgen.bundle.js', 'PptxGenJS');

  /* ---------- 工具 ---------- */

  function baseName(name) { return name.replace(/\.[^.]+$/, ''); }

  async function decodeImage(file) {
    try { return await createImageBitmap(file); }
    catch (e) { /* fallthrough */ }
    return await new Promise((res, rej) => {
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = () => res(img);
      img.onerror = () => {
        URL.revokeObjectURL(url);
        rej(new Error('无法解码此图片。HEIC / TIFF 依赖系统解码,请在 Safari(Mac / iPhone)中使用。'));
      };
      img.src = url;
    });
  }

  function imgToCanvas(img, flattenWhite) {
    const c = document.createElement('canvas');
    c.width = img.width || img.naturalWidth;
    c.height = img.height || img.naturalHeight;
    const ctx = c.getContext('2d');
    if (flattenWhite) { ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height); }
    ctx.drawImage(img, 0, 0);
    return c;
  }

  function canvasToBlob(canvas, type, quality) {
    return new Promise((res, rej) => {
      canvas.toBlob((b) => {
        if (!b) return rej(new Error('图像编码失败'));
        if (type && b.type !== type) return rej(new Error(`当前浏览器不支持导出 ${type.split('/')[1].toUpperCase()} 格式`));
        res(b);
      }, type, quality);
    });
  }

  async function canvasesToPdf(canvases, widthPtOf, heightPtOf, quality) {
    const { PDFDocument } = global.PDFLib;
    const pdf = await PDFDocument.create();
    for (let i = 0; i < canvases.length; i++) {
      const c = canvases[i];
      const jpg = await canvasToBlob(c, 'image/jpeg', quality || 0.92);
      const bytes = new Uint8Array(await jpg.arrayBuffer());
      const img = await pdf.embedJpg(bytes);
      const wPt = typeof widthPtOf === 'function' ? widthPtOf(i, c) : widthPtOf;
      const hPt = typeof heightPtOf === 'function' ? heightPtOf(i, c) : heightPtOf;
      const page = pdf.addPage([wPt, hPt]);
      page.drawImage(img, { x: 0, y: 0, width: wPt, height: hPt });
    }
    const out = await pdf.save();
    return new Blob([out], { type: 'application/pdf' });
  }

  async function zipBlobs(entries) {
    const zip = new JSZip();
    for (const e of entries) zip.file(e.name, e.blob);
    return await zip.generateAsync({ type: 'blob', compression: 'STORE' });
  }

  /* ---------- 图片转换 ---------- */

  async function imageToImage(file, fmt) {
    const img = await decodeImage(file);
    const flat = (fmt === 'jpeg' || fmt === 'bmp');
    const canvas = imgToCanvas(img, flat);
    if (fmt === 'png') return { blob: await canvasToBlob(canvas, 'image/png'), ext: 'png' };
    if (fmt === 'jpeg') return { blob: await canvasToBlob(canvas, 'image/jpeg', 0.92), ext: 'jpg' };
    if (fmt === 'webp') return { blob: await canvasToBlob(canvas, 'image/webp', 0.92), ext: 'webp' };
    const data = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height);
    if (fmt === 'bmp') return { blob: LiteEncoders.encodeBMP(data), ext: 'bmp' };
    if (fmt === 'tiff') return { blob: LiteEncoders.encodeTIFF(data), ext: 'tiff' };
    if (fmt === 'gif') return { blob: LiteEncoders.encodeGIF(data), ext: 'gif' };
    throw new Error('未知目标格式');
  }

  async function imagesToPdf(files) {
    const { PDFDocument } = global.PDFLib;
    const pdf = await PDFDocument.create();
    for (const file of files) {
      const bytes = new Uint8Array(await file.arrayBuffer());
      let embedded = null;
      if (/jpe?g$/i.test(file.name) || file.type === 'image/jpeg') {
        try { embedded = await pdf.embedJpg(bytes); } catch (e) { /* re-encode below */ }
      } else if (/png$/i.test(file.name) || file.type === 'image/png') {
        try { embedded = await pdf.embedPng(bytes); } catch (e) { /* re-encode below */ }
      }
      if (!embedded) {
        const img = await decodeImage(file);
        const canvas = imgToCanvas(img, false);
        const png = await canvasToBlob(canvas, 'image/png');
        embedded = await pdf.embedPng(new Uint8Array(await png.arrayBuffer()));
      }
      const wPt = embedded.width * 72 / 96, hPt = embedded.height * 72 / 96;
      const page = pdf.addPage([wPt, hPt]);
      page.drawImage(embedded, { x: 0, y: 0, width: wPt, height: hPt });
    }
    const out = await pdf.save();
    return new Blob([out], { type: 'application/pdf' });
  }

  /* ---------- PDF 转换 ---------- */

  async function openPdf(file) {
    const pdfjs = await getPdfjs();
    const data = new Uint8Array(await file.arrayBuffer());
    return await pdfjs.getDocument({ data, isEvalSupported: false }).promise;
  }

  async function pdfPageCanvas(doc, n, scale) {
    const page = await doc.getPage(n);
    const viewport = page.getViewport({ scale: scale || 2 });
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(viewport.width);
    canvas.height = Math.round(viewport.height);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport }).promise;
    return { canvas, page };
  }

  async function pdfToImages(file, fmt, onProgress) {
    const doc = await openPdf(file);
    const type = fmt === 'png' ? 'image/png' : 'image/jpeg';
    const ext = fmt === 'png' ? 'png' : 'jpg';
    const blobs = [];
    for (let i = 1; i <= doc.numPages; i++) {
      const { canvas } = await pdfPageCanvas(doc, i, 2);
      blobs.push({ name: `${baseName(file.name)}_第${i}页.${ext}`, blob: await canvasToBlob(canvas, type, 0.92) });
      if (onProgress) onProgress(i, doc.numPages);
    }
    if (blobs.length === 1) return { blob: blobs[0].blob, ext };
    return { blob: await zipBlobs(blobs), ext: 'zip', note: `${blobs.length} 页已打包` };
  }

  /** 提取 PDF 文本行(带字号/粗体信息) */
  async function pdfExtractLines(doc, onProgress) {
    const pages = [];
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i);
      const tc = await page.getTextContent();
      const items = tc.items.filter(it => it.str !== undefined);
      // 按 Y 分组为行
      const rows = [];
      for (const it of items) {
        const y = it.transform[5], x = it.transform[4];
        const h = Math.abs(it.transform[3]) || Math.abs(it.transform[0]) || 10;
        let row = rows.find(r => Math.abs(r.y - y) < h * 0.5);
        if (!row) { row = { y, items: [] }; rows.push(row); }
        row.items.push({ x, str: it.str, h, font: it.fontName });
      }
      rows.sort((a, b) => b.y - a.y);
      const lines = rows.map(r => {
        r.items.sort((a, b) => a.x - b.x);
        let text = '';
        let px = null;
        for (const it of r.items) {
          if (px !== null && it.x - px > it.h * 0.5 && text && !text.endsWith(' ')) text += ' ';
          text += it.str;
          px = it.x + (it.str.length * it.h * 0.5);
        }
        const maxH = Math.max(...r.items.map(i2 => i2.h));
        return { text: text, size: maxH, y: r.y };
      }).filter(l => l.text.trim());
      pages.push({ lines, styles: tc.styles });
      if (onProgress) onProgress(i, doc.numPages);
    }
    return pages;
  }

  async function pdfToTxt(file, onProgress) {
    const doc = await openPdf(file);
    const pages = await pdfExtractLines(doc, onProgress);
    const text = pages.map(p => p.lines.map(l => l.text).join('\n')).join('\n\n────────  分页  ────────\n\n');
    return { blob: new Blob(['﻿' + text], { type: 'text/plain;charset=utf-8' }), ext: 'txt' };
  }

  async function pdfToDocx(file, onProgress) {
    const doc = await openPdf(file);
    const pages = await pdfExtractLines(doc, onProgress);
    const docx = await getDocxLib();
    const { Document, Packer, Paragraph, TextRun, PageBreak } = docx;

    // 估计正文字号(出现最多的字号)
    const sizeCount = {};
    for (const p of pages) for (const l of p.lines) {
      const s = Math.round(l.size);
      sizeCount[s] = (sizeCount[s] || 0) + l.text.length;
    }
    let bodySize = 11;
    let maxCnt = 0;
    for (const s in sizeCount) if (sizeCount[s] > maxCnt) { maxCnt = sizeCount[s]; bodySize = parseInt(s, 10); }

    const children = [];
    pages.forEach((p, pi) => {
      // 相邻行距大 -> 分段
      let paraLines = [];
      const flush = () => {
        if (!paraLines.length) return;
        const sz = Math.max(...paraLines.map(l => l.size));
        const isHeading = sz > bodySize * 1.25 && paraLines.map(l => l.text).join('').length < 60;
        const halfPt = Math.max(12, Math.min(72, Math.round(sz * 2)));
        children.push(new Paragraph({
          children: [new TextRun({ text: paraLines.map(l => l.text).join(' '), bold: isHeading, size: halfPt })],
          spacing: { after: isHeading ? 240 : 120 },
        }));
        paraLines = [];
      };
      for (let i = 0; i < p.lines.length; i++) {
        const l = p.lines[i];
        paraLines.push(l);
        const next = p.lines[i + 1];
        if (!next || (l.y - next.y) > l.size * 1.9 || Math.round(next.size) !== Math.round(l.size)) flush();
      }
      flush();
      if (pi < pages.length - 1) children.push(new Paragraph({ children: [new PageBreak()] }));
    });

    const d = new Document({ sections: [{ children: children.length ? children : [new Paragraph('')] }] });
    const blob = await Packer.toBlob(d);
    return { blob, ext: 'docx', note: '版式为尽力还原,建议人工校对' };
  }

  async function pdfToPptx(file, onProgress) {
    const doc = await openPdf(file);
    const PptxGenJS = await getPptxGen();
    const pptx = new PptxGenJS();
    const first = await doc.getPage(1);
    const vp = first.getViewport({ scale: 1 });
    const wIn = vp.width / 72, hIn = vp.height / 72;
    pptx.defineLayout({ name: 'PDF', width: wIn, height: hIn });
    pptx.layout = 'PDF';
    for (let i = 1; i <= doc.numPages; i++) {
      const { canvas } = await pdfPageCanvas(doc, i, 2);
      const dataUrl = canvas.toDataURL('image/jpeg', 0.9);
      const slide = pptx.addSlide();
      slide.addImage({ data: dataUrl, x: 0, y: 0, w: wIn, h: hIn });
      if (onProgress) onProgress(i, doc.numPages);
    }
    const blob = await pptx.write('blob');
    return { blob, ext: 'pptx', note: '每页 PDF 已转为整页图片幻灯片' };
  }

  /* ---------- Word / PPT 转换 ---------- */

  async function docxToPdf(file, onProgress) {
    const model = await LiteDocx.parseDocx(await file.arrayBuffer());
    if (onProgress) onProgress(1, 2);
    const canvases = await LiteDocx.layoutDocx(model, 2);
    const blob = await canvasesToPdf(canvases, model.page.w, model.page.h, 0.92);
    if (onProgress) onProgress(2, 2);
    return { blob, ext: 'pdf' };
  }

  async function docxToTxt(file) {
    const model = await LiteDocx.parseDocx(await file.arrayBuffer());
    const text = LiteDocx.docxToText(model);
    return { blob: new Blob(['﻿' + text], { type: 'text/plain;charset=utf-8' }), ext: 'txt' };
  }

  async function pptxToPdf(file, onProgress) {
    const { canvases, widthPt, heightPt } = await LitePptx.renderPptx(await file.arrayBuffer(), onProgress);
    const blob = await canvasesToPdf(canvases, widthPt, heightPt, 0.92);
    return { blob, ext: 'pdf' };
  }

  async function pptxToImages(file, onProgress) {
    const { canvases } = await LitePptx.renderPptx(await file.arrayBuffer(), onProgress);
    if (canvases.length === 1) {
      return { blob: await canvasToBlob(canvases[0], 'image/png'), ext: 'png' };
    }
    const entries = [];
    for (let i = 0; i < canvases.length; i++) {
      entries.push({ name: `${baseName(file.name)}_第${i + 1}页.png`, blob: await canvasToBlob(canvases[i], 'image/png') });
    }
    return { blob: await zipBlobs(entries), ext: 'zip', note: `${canvases.length} 页已打包` };
  }

  global.LiteConvert = {
    imageToImage, imagesToPdf,
    pdfToImages, pdfToTxt, pdfToDocx, pdfToPptx,
    docxToPdf, docxToTxt,
    pptxToPdf, pptxToImages,
    zipBlobs, baseName,
  };
})(window);
