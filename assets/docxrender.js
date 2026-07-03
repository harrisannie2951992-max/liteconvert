/* 轻转 LiteConvert — Word (.docx) 解析与排版引擎
 * 解析 OOXML,将文档逐页排版到 canvas(支持中英文混排、粗斜体、颜色、
 * 对齐、图片、基础表格与列表),用于 Word -> PDF / 图片。
 * MIT License
 */
(function (global) {
  'use strict';

  const TWIP = 20;               // twips per pt
  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  function q(el, name) { return el.getElementsByTagNameNS(NS_W, name); }
  function child(el, name) {
    for (const c of el.children) if (c.localName === name) return c;
    return null;
  }
  function attr(el, name) {
    return el.getAttributeNS(NS_W, name) || el.getAttribute('w:' + name) || el.getAttribute(name);
  }

  /* ---------- 解析 ---------- */

  async function parseDocx(arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const docFile = zip.file('word/document.xml');
    if (!docFile) throw new Error('不是有效的 .docx 文件(未找到 word/document.xml)');
    const xml = new DOMParser().parseFromString(await docFile.async('text'), 'application/xml');

    // 关系表 -> 媒体文件
    const rels = {};
    const relFile = zip.file('word/_rels/document.xml.rels');
    if (relFile) {
      const rx = new DOMParser().parseFromString(await relFile.async('text'), 'application/xml');
      for (const r of rx.getElementsByTagName('Relationship')) {
        rels[r.getAttribute('Id')] = r.getAttribute('Target');
      }
    }

    // 样式表:默认字号 + 各样式的加粗/字号
    const styles = { default: { sz: 11 }, map: {} };
    const styFile = zip.file('word/styles.xml');
    if (styFile) {
      const sx = new DOMParser().parseFromString(await styFile.async('text'), 'application/xml');
      const dd = sx.getElementsByTagNameNS(NS_W, 'docDefaults')[0];
      if (dd) {
        const sz = dd.getElementsByTagNameNS(NS_W, 'sz')[0];
        if (sz) styles.default.sz = parseInt(attr(sz, 'val'), 10) / 2 || 11;
      }
      for (const st of sx.getElementsByTagNameNS(NS_W, 'style')) {
        const id = attr(st, 'styleId');
        if (!id) continue;
        const s = {};
        const rPr = child(st, 'rPr');
        if (rPr) {
          if (child(rPr, 'b') && attr(child(rPr, 'b'), 'val') !== '0') s.b = true;
          const sz = child(rPr, 'sz');
          if (sz) s.sz = parseInt(attr(sz, 'val'), 10) / 2;
          const col = child(rPr, 'color');
          if (col && attr(col, 'val') && attr(col, 'val') !== 'auto') s.color = '#' + attr(col, 'val');
        }
        styles.map[id] = s;
      }
    }

    // 页面设置
    const body = xml.getElementsByTagNameNS(NS_W, 'body')[0];
    let page = { w: 595.3, h: 841.9, mt: 72, mb: 72, ml: 90, mr: 90 };
    const sect = body ? q(body, 'sectPr')[0] : null;
    if (sect) {
      const pgSz = child(sect, 'pgSz'), pgMar = child(sect, 'pgMar');
      if (pgSz) {
        page.w = parseInt(attr(pgSz, 'w'), 10) / TWIP || page.w;
        page.h = parseInt(attr(pgSz, 'h'), 10) / TWIP || page.h;
      }
      if (pgMar) {
        page.mt = parseInt(attr(pgMar, 'top'), 10) / TWIP || page.mt;
        page.mb = parseInt(attr(pgMar, 'bottom'), 10) / TWIP || page.mb;
        page.ml = parseInt(attr(pgMar, 'left'), 10) / TWIP || page.ml;
        page.mr = parseInt(attr(pgMar, 'right'), 10) / TWIP || page.mr;
      }
    }

    const blocks = [];
    for (const el of body.children) {
      if (el.localName === 'p') blocks.push(parseParagraph(el, styles));
      else if (el.localName === 'tbl') blocks.push(parseTable(el, styles));
    }

    return { blocks, page, rels, zip, styles };
  }

  function parseRunProps(rPr, base) {
    const st = Object.assign({}, base);
    if (!rPr) return st;
    const on = (n) => { const e = child(rPr, n); return e && attr(e, 'val') !== '0' && attr(e, 'val') !== 'false'; };
    if (on('b')) st.b = true;
    if (on('i')) st.i = true;
    if (on('u') && attr(child(rPr, 'u'), 'val') !== 'none') st.u = true;
    if (on('strike')) st.strike = true;
    const sz = child(rPr, 'sz');
    if (sz) st.sz = parseInt(attr(sz, 'val'), 10) / 2;
    const col = child(rPr, 'color');
    if (col && attr(col, 'val') && attr(col, 'val') !== 'auto') st.color = '#' + attr(col, 'val');
    const hl = child(rPr, 'highlight');
    if (hl && attr(hl, 'val') !== 'none') st.highlight = attr(hl, 'val');
    return st;
  }

  function parseParagraph(p, styles) {
    const para = { type: 'p', align: 'left', runs: [], spacingAfter: 6, indent: 0, bullet: null, pageBreakBefore: false };
    const pPr = child(p, 'pPr');
    let baseStyle = { sz: styles.default.sz };
    if (pPr) {
      const st = child(pPr, 'pStyle');
      if (st && styles.map[attr(st, 'val')]) baseStyle = Object.assign({}, baseStyle, styles.map[attr(st, 'val')]);
      const jc = child(pPr, 'jc');
      if (jc) {
        const v = attr(jc, 'val');
        if (v === 'center') para.align = 'center';
        else if (v === 'right' || v === 'end') para.align = 'right';
        else if (v === 'both' || v === 'distribute') para.align = 'justify';
      }
      const ind = child(pPr, 'ind');
      if (ind) para.indent = (parseInt(attr(ind, 'left') || attr(ind, 'start'), 10) || 0) / TWIP;
      const numPr = child(pPr, 'numPr');
      if (numPr) {
        const ilvl = child(numPr, 'ilvl');
        para.bullet = { level: ilvl ? parseInt(attr(ilvl, 'val'), 10) || 0 : 0 };
      }
      if (child(pPr, 'pageBreakBefore')) para.pageBreakBefore = true;
      const spacing = child(pPr, 'spacing');
      if (spacing && attr(spacing, 'after') != null) para.spacingAfter = (parseInt(attr(spacing, 'after'), 10) || 0) / TWIP;
      const rPrP = child(pPr, 'rPr');
      baseStyle = parseRunProps(rPrP, baseStyle);
    }
    para.baseStyle = baseStyle;

    const walk = (node) => {
      for (const el of node.children) {
        if (el.localName === 'r') {
          const st = parseRunProps(child(el, 'rPr'), baseStyle);
          for (const c of el.children) {
            if (c.localName === 't') para.runs.push({ text: c.textContent, st });
            else if (c.localName === 'tab') para.runs.push({ text: '    ', st });
            else if (c.localName === 'br') {
              if (attr(c, 'type') === 'page') para.runs.push({ pageBreak: true });
              else para.runs.push({ br: true });
            }
            else if (c.localName === 'drawing' || c.localName === 'pict') {
              const img = parseImage(c);
              if (img) para.runs.push(img);
            }
          }
        } else if (el.localName === 'hyperlink' || el.localName === 'smartTag') {
          walk(el);
        }
      }
    };
    walk(p);
    return para;
  }

  function parseImage(drawing) {
    const blips = drawing.getElementsByTagName('a:blip');
    const blip = blips.length ? blips[0] : drawing.getElementsByTagName('blip')[0];
    if (!blip) return null;
    const rid = blip.getAttribute('r:embed') || blip.getAttribute('embed');
    if (!rid) return null;
    let cx = 0, cy = 0;
    const ext = drawing.getElementsByTagName('wp:extent')[0] || drawing.getElementsByTagName('a:ext')[0];
    if (ext) { cx = parseInt(ext.getAttribute('cx'), 10) || 0; cy = parseInt(ext.getAttribute('cy'), 10) || 0; }
    return { image: rid, wPt: cx / 12700, hPt: cy / 12700 }; // EMU -> pt
  }

  function parseTable(tbl, styles) {
    const t = { type: 'table', rows: [], colWidths: [] };
    const grid = child(tbl, 'tblGrid');
    if (grid) for (const g of grid.children) {
      if (g.localName === 'gridCol') t.colWidths.push((parseInt(attr(g, 'w'), 10) || 1200) / TWIP);
    }
    for (const tr of tbl.children) {
      if (tr.localName !== 'tr') continue;
      const row = [];
      for (const tc of tr.children) {
        if (tc.localName !== 'tc') continue;
        const cell = { paras: [], span: 1 };
        const tcPr = child(tc, 'tcPr');
        if (tcPr) {
          const gs = child(tcPr, 'gridSpan');
          if (gs) cell.span = parseInt(attr(gs, 'val'), 10) || 1;
          const shd = child(tcPr, 'shd');
          if (shd && attr(shd, 'fill') && attr(shd, 'fill') !== 'auto') cell.fill = '#' + attr(shd, 'fill');
        }
        for (const el of tc.children) {
          if (el.localName === 'p') cell.paras.push(parseParagraph(el, styles));
        }
        row.push(cell);
      }
      if (row.length) t.rows.push(row);
    }
    return t;
  }

  /* ---------- 排版渲染 ---------- */

  const FONT_STACK = '-apple-system, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "Noto Sans CJK SC", "Noto Sans SC", "DejaVu Sans", sans-serif';

  function fontOf(st, S) {
    const sz = (st.sz || 11) * S;
    return `${st.i ? 'italic ' : ''}${st.b ? 'bold ' : ''}${sz}px ${FONT_STACK}`;
  }

  /** 将 run 拆成可换行的最小片段:CJK 逐字、拉丁按词 */
  function segments(text) {
    const out = [];
    const re = /([一-鿿　-〿＀-￯぀-ヿ])|(\s+)|([^\s一-鿿　-〿＀-￯぀-ヿ]+)/g;
    let m;
    while ((m = re.exec(text))) out.push(m[0]);
    return out;
  }

  async function loadMedia(model, rid) {
    const target = model.rels[rid];
    if (!target) return null;
    const path = 'word/' + target.replace(/^\//, '').replace(/^word\//, '');
    const f = model.zip.file(path) || model.zip.file(target.replace(/^\//, ''));
    if (!f) return null;
    const blob = await f.async('blob');
    try {
      return await createImageBitmap(blob);
    } catch (e) {
      return await new Promise((res) => {
        const img = new Image();
        img.onload = () => res(img);
        img.onerror = () => res(null);
        img.src = URL.createObjectURL(blob);
      });
    }
  }

  /**
   * 将解析结果排版为多个 canvas 页面
   * @returns {Promise<HTMLCanvasElement[]>}
   */
  async function layoutDocx(model, S) {
    S = S || 2; // pt -> px 缩放
    const { page } = model;
    const W = Math.round(page.w * S), H = Math.round(page.h * S);
    const ml = page.ml * S, mr = page.mr * S, mt = page.mt * S, mb = page.mb * S;
    const contentW = W - ml - mr;

    const pages = [];
    let ctx = null, y = 0;

    const newPage = () => {
      const c = document.createElement('canvas');
      c.width = W; c.height = H;
      ctx = c.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, W, H);
      ctx.textBaseline = 'alphabetic';
      pages.push(c);
      y = mt;
    };
    newPage();

    const ensure = (need) => { if (y + need > H - mb && y > mt + 1) newPage(); };

    // 逐段渲染
    const env = {
      S, x: ml, w: contentW, model, H, mb,
      getCtx: () => ctx,
      getY: () => y,
      setY: (v) => { y = v; },
      newPage: () => { newPage(); },
    };
    for (const block of model.blocks) {
      if (block.type === 'p') {
        y = await renderParagraph(block, ctx, env);
        y += block.spacingAfter * S;
      } else if (block.type === 'table') {
        y = await renderTable(block, env);
        y += 8 * S;
      }
    }
    return pages;
  }

  async function renderParagraph(p, ctx0, env) {
    const { S, x, w, model } = env;
    let ctx = ctx0;
    if (p.pageBreakBefore) { env.newPage(); ctx = env.getCtx ? env.getCtx() : ctx; }
    let y = env.getY ? env.getY() : env.y;

    const indent = (p.indent || 0) * S + (p.bullet ? (14 + p.bullet.level * 14) * S : 0);
    const availW = w - indent;

    // 组行
    const lines = [];
    let line = [], lineW = 0;
    const pushLine = () => { lines.push({ parts: line, width: lineW }); line = []; lineW = 0; };

    for (const run of p.runs) {
      if (run.pageBreak) { line.push({ pageBreak: true }); pushLine(); continue; }
      if (run.br) { pushLine(); continue; }
      if (run.image) {
        const bmp = await loadMedia(model, run.image);
        if (bmp) {
          let iw = (run.wPt || bmp.width * 0.75) * S, ih = (run.hPt || bmp.height * 0.75) * S;
          if (iw > availW) { ih *= availW / iw; iw = availW; }
          if (line.length) pushLine();
          line.push({ img: bmp, w: iw, h: ih });
          lineW = iw;
          pushLine();
        }
        continue;
      }
      const st = run.st || p.baseStyle;
      const canvasCtx = ctx; // measure 用当前上下文
      canvasCtx.font = fontOf(st, S);
      for (const seg of segments(run.text)) {
        const sw = canvasCtx.measureText(seg).width;
        if (lineW + sw > availW && line.length && seg.trim() !== '') pushLine();
        line.push({ text: seg, st, w: sw });
        lineW += sw;
      }
    }
    if (line.length) pushLine();
    if (!lines.length) { env.setY && env.setY(y + (p.baseStyle.sz || 11) * 1.15 * S); return y + (p.baseStyle.sz || 11) * 1.15 * S; }

    // 逐行绘制
    let first = true;
    for (const ln of lines) {
      if (ln.parts.some(pt => pt.pageBreak)) { env.newPage(); ctx = env.getCtx ? env.getCtx() : ctx; y = env.getY(); first = true; continue; }
      const isImg = ln.parts.length === 1 && ln.parts[0].img;
      let maxSz = 0;
      for (const pt of ln.parts) if (pt.st && (pt.st.sz || 11) > maxSz) maxSz = pt.st.sz || 11;
      if (!maxSz) maxSz = p.baseStyle.sz || 11;
      const lineH = isImg ? ln.parts[0].h + 6 * S : maxSz * 1.55 * S;

      if (y + lineH > env.H - env.mb && y > 0) { env.newPage(); ctx = env.getCtx ? env.getCtx() : ctx; y = env.getY(); }

      let dx = x + indent;
      if (p.align === 'center') dx = x + indent + (availW - ln.width) / 2;
      else if (p.align === 'right') dx = x + indent + (availW - ln.width);

      if (first && p.bullet) {
        ctx.font = fontOf(p.baseStyle, S);
        ctx.fillStyle = '#333333';
        const bch = ['•', '◦', '▪'][p.bullet.level % 3];
        ctx.fillText(bch, x + (p.bullet.level * 14 + 2) * S, y + maxSz * 1.15 * S);
        first = false;
      }

      if (isImg) {
        ctx.drawImage(ln.parts[0].img, dx, y + 3 * S, ln.parts[0].w, ln.parts[0].h);
      } else {
        const baseline = y + maxSz * 1.15 * S;
        for (const pt of ln.parts) {
          if (!pt.text) continue;
          const st = pt.st;
          ctx.font = fontOf(st, S);
          if (st.highlight) {
            const hl = { yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', red: '#ff5555', blue: '#5555ff', lightGray: '#dddddd', darkGray: '#999999' }[st.highlight] || '#ffff88';
            ctx.fillStyle = hl;
            ctx.fillRect(dx, y + 2 * S, pt.w, maxSz * 1.4 * S);
          }
          ctx.fillStyle = st.color || '#1a1a1a';
          ctx.fillText(pt.text, dx, baseline);
          if (st.u || st.strike) {
            ctx.strokeStyle = st.color || '#1a1a1a';
            ctx.lineWidth = Math.max(1, 0.06 * (st.sz || 11) * S);
            ctx.beginPath();
            const ly = st.strike ? baseline - (st.sz || 11) * 0.3 * S : baseline + 2 * S;
            ctx.moveTo(dx, ly); ctx.lineTo(dx + pt.w, ly);
            ctx.stroke();
          }
          dx += pt.w;
        }
      }
      y += lineH;
      env.setY && env.setY(y);
    }
    return y;
  }

  async function renderTable(t, env) {
    const { S, x, w, model } = env;
    let y = env.getY();
    const totalTw = t.colWidths.reduce((a, b) => a + b, 0) || 1;
    const scale = w / (totalTw * S) < 1 ? w / (totalTw * S) : 1;
    const colPx = t.colWidths.map(cw => cw * S * scale);
    const pad = 4 * S;

    for (const row of t.rows) {
      // 先离屏排版每个单元格,取行高
      const cellCanvases = [];
      let rowH = 0;
      let cx = 0;
      for (let ci = 0; ci < row.length; ci++) {
        const cell = row[ci];
        let cw = 0;
        let base = 0;
        for (let k = 0, gi = 0; k < row.length && gi < colPx.length; k++) {
          const span = row[k].span || 1;
          if (k === ci) { for (let s = 0; s < span && gi + s < colPx.length; s++) cw += colPx[gi + s]; base = gi; }
          gi += span;
        }
        if (!cw) cw = colPx[ci] || 100 * S;
        const cc = document.createElement('canvas');
        cc.width = Math.max(2, Math.round(cw - pad * 2));
        cc.height = 4000;
        const cctx = cc.getContext('2d');
        cctx.textBaseline = 'alphabetic';
        let cy = 0;
        for (const p of cell.paras) {
          const subenv = {
            S, x: 0, w: cc.width, model, H: 1e9, mb: 0,
            getY: () => cy, setY: (v) => { cy = v; },
            newPage: () => cy, getCtx: () => cctx,
          };
          cy = await renderParagraph(p, cctx, subenv);
          cy += (p.spacingAfter || 4) * S * 0.5;
        }
        cellCanvases.push({ canvas: cc, used: Math.max(cy, 14 * S), w: cw, fill: cell.fill });
        if (cy > rowH) rowH = cy;
      }
      rowH = Math.max(rowH, 16 * S) + pad * 2;

      if (y + rowH > env.H - env.mb) { env.newPage(); y = env.getY(); }
      const ctx = env.getCtx();

      cx = x;
      for (const cc of cellCanvases) {
        if (cc.fill) { ctx.fillStyle = cc.fill; ctx.fillRect(cx, y, cc.w, rowH); }
        ctx.strokeStyle = '#c9ced8';
        ctx.lineWidth = 1;
        ctx.strokeRect(cx + 0.5, y + 0.5, cc.w, rowH);
        ctx.drawImage(cc.canvas, 0, 0, cc.canvas.width, Math.min(cc.used, rowH - pad), cx + pad, y + pad, cc.canvas.width, Math.min(cc.used, rowH - pad));
        cx += cc.w;
      }
      y += rowH;
      env.setY(y);
    }
    return y;
  }

  /** 提取纯文本 */
  function docxToText(model) {
    const out = [];
    for (const b of model.blocks) {
      if (b.type === 'p') {
        out.push(b.runs.map(r => r.text || (r.br ? '\n' : '')).join(''));
      } else if (b.type === 'table') {
        for (const row of b.rows) {
          out.push(row.map(c => c.paras.map(p => p.runs.map(r => r.text || '').join('')).join(' ')).join('\t'));
        }
      }
    }
    return out.join('\n');
  }

  global.LiteDocx = { parseDocx, layoutDocx, docxToText };
})(window);
