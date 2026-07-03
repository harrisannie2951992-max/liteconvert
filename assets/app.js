/* 轻转 LiteConvert — 界面逻辑。MIT License */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const dropzone = $('dropzone'), fileInput = $('fileInput');
  const workbench = $('workbench'), fileList = $('fileList');
  const convertBtn = $('convertBtn'), clearBtn = $('clearBtn');
  const mergeOpt = $('mergeOpt'), mergePdf = $('mergePdf');
  const wbFoot = $('wbFoot'), downloadAllBtn = $('downloadAllBtn');

  const IMG_EXTS = ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'tif', 'tiff', 'heic', 'heif', 'avif'];

  const TARGETS = {
    image: [
      ['png', 'PNG'], ['jpeg', 'JPG'], ['webp', 'WebP'],
      ['bmp', 'BMP'], ['tiff', 'TIFF'], ['gif', 'GIF'], ['pdf', 'PDF'],
    ],
    pdf: [
      ['png', '图片 PNG'], ['jpeg', '图片 JPG'],
      ['docx', 'Word 文档'], ['pptx', 'PPT 演示'], ['txt', '纯文本'],
    ],
    docx: [['pdf', 'PDF'], ['txt', '纯文本']],
    pptx: [['pdf', 'PDF'], ['png', '图片 PNG']],
  };
  const TYPE_LABEL = { image: 'IMG', pdf: 'PDF', docx: 'DOC', pptx: 'PPT' };

  let items = [];   // {id, file, kind, target, status, result:{blob,ext,note}, outName}
  let nextId = 1;

  function detect(file) {
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    if (IMG_EXTS.includes(ext)) return 'image';
    if (ext === 'pdf') return 'pdf';
    if (ext === 'docx') return 'docx';
    if (ext === 'pptx') return 'pptx';
    if (ext === 'doc' || ext === 'ppt') return 'legacy';
    return null;
  }

  function fmtSize(n) {
    if (n < 1024) return n + ' B';
    if (n < 1048576) return (n / 1024).toFixed(1) + ' KB';
    return (n / 1048576).toFixed(1) + ' MB';
  }

  function addFiles(files) {
    for (const f of files) {
      const kind = detect(f);
      if (kind === 'legacy') {
        alert(`「${f.name}」是旧版格式。请先在 Word / WPS / Keynote 中另存为 .docx / .pptx 再转换。`);
        continue;
      }
      if (!kind) { alert(`暂不支持「${f.name}」的格式`); continue; }
      items.push({ id: nextId++, file: f, kind, target: TARGETS[kind][0][0], status: 'ready', result: null });
    }
    render();
  }

  function render() {
    workbench.hidden = items.length === 0;
    fileList.innerHTML = '';
    const imgPdfCount = items.filter(i => i.kind === 'image' && i.target === 'pdf').length;
    mergeOpt.hidden = imgPdfCount < 2;
    const done = items.filter(i => i.status === 'done');
    wbFoot.hidden = done.length < 2;

    for (const it of items) {
      const li = document.createElement('li');
      li.className = 'file-item';

      const icon = document.createElement('div');
      icon.className = 'fi-icon t-' + it.kind;
      icon.textContent = TYPE_LABEL[it.kind];
      li.appendChild(icon);

      const meta = document.createElement('div');
      meta.className = 'fi-meta';
      const nm = document.createElement('div');
      nm.className = 'fi-name';
      nm.textContent = it.file.name;
      const sub = document.createElement('div');
      sub.className = 'fi-sub';
      if (it.status === 'error') { sub.classList.add('err'); sub.textContent = it.error; }
      else if (it.status === 'done') {
        sub.classList.add('ok');
        sub.textContent = `完成 · ${fmtSize(it.result.blob.size)}` + (it.result.note ? ` · ${it.result.note}` : '');
      }
      else if (it.status === 'working') sub.textContent = it.progress || '转换中…';
      else sub.textContent = fmtSize(it.file.size);
      meta.appendChild(nm); meta.appendChild(sub);
      li.appendChild(meta);

      if (it.status === 'working') {
        const sp = document.createElement('div'); sp.className = 'fi-spin'; li.appendChild(sp);
      } else if (it.status === 'done') {
        const dl = document.createElement('button');
        dl.className = 'fi-dl'; dl.type = 'button';
        dl.textContent = '下载 .' + it.result.ext;
        dl.addEventListener('click', () => saveBlob(it.result.blob, it.outName));
        li.appendChild(dl);
      } else {
        const sel = document.createElement('select');
        sel.className = 'fi-target';
        for (const [v, label] of TARGETS[it.kind]) {
          const o = document.createElement('option');
          o.value = v; o.textContent = '→ ' + label;
          if (v === it.target) o.selected = true;
          sel.appendChild(o);
        }
        sel.addEventListener('change', () => { it.target = sel.value; render(); });
        li.appendChild(sel);
      }

      const rm = document.createElement('button');
      rm.className = 'fi-remove'; rm.type = 'button'; rm.title = '移除';
      rm.textContent = '✕';
      rm.addEventListener('click', () => { items = items.filter(x => x.id !== it.id); render(); });
      li.appendChild(rm);

      fileList.appendChild(li);
    }
    convertBtn.disabled = !items.some(i => i.status === 'ready' || i.status === 'error');
  }

  function saveBlob(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }

  const C = () => window.LiteConvert;

  async function convertItem(it, onProgress) {
    const f = it.file, t = it.target;
    if (it.kind === 'image') {
      if (t === 'pdf') return { blob: await C().imagesToPdf([f]), ext: 'pdf' };
      return await C().imageToImage(f, t);
    }
    if (it.kind === 'pdf') {
      if (t === 'png' || t === 'jpeg') return await C().pdfToImages(f, t, onProgress);
      if (t === 'docx') return await C().pdfToDocx(f, onProgress);
      if (t === 'pptx') return await C().pdfToPptx(f, onProgress);
      if (t === 'txt') return await C().pdfToTxt(f, onProgress);
    }
    if (it.kind === 'docx') {
      if (t === 'pdf') return await C().docxToPdf(f, onProgress);
      if (t === 'txt') return await C().docxToTxt(f);
    }
    if (it.kind === 'pptx') {
      if (t === 'pdf') return await C().pptxToPdf(f, onProgress);
      if (t === 'png') return await C().pptxToImages(f, onProgress);
    }
    throw new Error('不支持的转换');
  }

  async function convertAll() {
    convertBtn.disabled = true;

    // 图片合并为一个 PDF 的特殊路径
    const mergeThese = mergePdf.checked
      ? items.filter(i => i.kind === 'image' && i.target === 'pdf' && i.status !== 'done')
      : [];
    if (mergeThese.length >= 2) {
      for (const it of mergeThese) { it.status = 'working'; }
      render();
      try {
        const blob = await C().imagesToPdf(mergeThese.map(i => i.file));
        const name = C().baseName(mergeThese[0].file.name) + `_等${mergeThese.length}张.pdf`;
        mergeThese.forEach((it, idx) => {
          if (idx === 0) {
            it.status = 'done';
            it.result = { blob, ext: 'pdf', note: `${mergeThese.length} 张图片已合并` };
            it.outName = name;
          } else {
            items = items.filter(x => x.id !== it.id);
          }
        });
      } catch (e) {
        mergeThese.forEach(it => { it.status = 'error'; it.error = e.message || '转换失败'; });
      }
      render();
    }

    for (const it of items) {
      if (it.status === 'done' || it.status === 'working') continue;
      it.status = 'working';
      it.progress = '转换中…';
      render();
      try {
        const res = await convertItem(it, (n, total) => {
          it.progress = `转换中… ${n} / ${total}`;
          render();
        });
        it.status = 'done';
        it.result = res;
        it.outName = C().baseName(it.file.name) + '.' + res.ext;
      } catch (e) {
        console.error(e);
        it.status = 'error';
        it.error = (e && e.message) || '转换失败,请重试';
      }
      render();
    }
    render();
  }

  async function downloadAll() {
    const done = items.filter(i => i.status === 'done');
    if (!done.length) return;
    const blob = await C().zipBlobs(done.map(i => ({ name: i.outName, blob: i.result.blob })));
    saveBlob(blob, '轻转结果_' + new Date().toISOString().slice(0, 10) + '.zip');
  }

  /* ---------- 事件 ---------- */

  dropzone.addEventListener('click', () => fileInput.click());
  dropzone.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') fileInput.click(); });
  fileInput.addEventListener('change', () => { addFiles(fileInput.files); fileInput.value = ''; });

  ['dragenter', 'dragover'].forEach(ev => dropzone.addEventListener(ev, (e) => {
    e.preventDefault(); dropzone.classList.add('dragover');
  }));
  ['dragleave', 'drop'].forEach(ev => dropzone.addEventListener(ev, (e) => {
    e.preventDefault(); dropzone.classList.remove('dragover');
  }));
  dropzone.addEventListener('drop', (e) => addFiles(e.dataTransfer.files));
  document.addEventListener('dragover', (e) => e.preventDefault());
  document.addEventListener('drop', (e) => e.preventDefault());

  convertBtn.addEventListener('click', convertAll);
  clearBtn.addEventListener('click', () => { items = []; render(); });
  downloadAllBtn.addEventListener('click', downloadAll);

  // 安装弹窗
  const modal = $('installModal');
  $('installBtn').addEventListener('click', () => { modal.hidden = false; });
  $('installClose').addEventListener('click', () => { modal.hidden = true; });
  modal.addEventListener('click', (e) => { if (e.target === modal) modal.hidden = true; });

  // PWA
  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
})();
