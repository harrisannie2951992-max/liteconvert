/* 轻转 LiteConvert — PPT (.pptx) 解析与渲染引擎
 * 解析 OOXML 幻灯片,把每页渲染到 canvas(文本框、图片、基础形状、
 * 纯色/主题色、背景),用于 PPT -> PDF / 图片。
 * MIT License
 */
(function (global) {
  'use strict';

  const EMU_PER_PX = 9525;     // 96dpi
  const SCALE = 2;             // 渲染倍率(清晰度)

  function tag(el, name) {
    for (const c of el.children) if (c.localName === name) return c;
    return null;
  }
  function tags(el, name) {
    const out = [];
    for (const c of el.children) if (c.localName === name) out.push(c);
    return out;
  }
  function findDeep(el, name) {
    if (!el) return null;
    if (el.localName === name) return el;
    for (const c of el.children) {
      const r = findDeep(c, name);
      if (r) return r;
    }
    return null;
  }

  async function xmlOf(zip, path) {
    const f = zip.file(path);
    if (!f) return null;
    return new DOMParser().parseFromString(await f.async('text'), 'application/xml');
  }

  async function relsOf(zip, partPath) {
    const dir = partPath.substring(0, partPath.lastIndexOf('/'));
    const name = partPath.substring(partPath.lastIndexOf('/') + 1);
    const rx = await xmlOf(zip, `${dir}/_rels/${name}.rels`);
    const map = {};
    if (rx) for (const r of rx.getElementsByTagName('Relationship')) {
      let t = r.getAttribute('Target');
      if (!t.startsWith('/')) {
        // 相对路径 -> 绝对
        const parts = (dir + '/' + t).split('/');
        const st = [];
        for (const p of parts) {
          if (p === '..') st.pop();
          else if (p !== '.') st.push(p);
        }
        t = st.join('/');
      } else t = t.slice(1);
      map[r.getAttribute('Id')] = { target: t, type: r.getAttribute('Type') || '' };
    }
    return map;
  }

  /* ---------- 颜色 ---------- */

  function parseTheme(themeXml) {
    const scheme = {};
    if (!themeXml) return scheme;
    const cs = themeXml.getElementsByTagName('a:clrScheme')[0];
    if (!cs) return scheme;
    for (const c of cs.children) {
      const name = c.localName; // dk1 lt1 dk2 lt2 accent1..6 hlink folHlink
      const srgb = findDeep(c, 'srgbClr');
      const sys = findDeep(c, 'sysClr');
      if (srgb) scheme[name] = '#' + srgb.getAttribute('val');
      else if (sys) scheme[name] = '#' + (sys.getAttribute('lastClr') || (name.startsWith('dk') ? '000000' : 'FFFFFF'));
    }
    return scheme;
  }

  function applyLum(hex, mods) {
    let r = parseInt(hex.slice(1, 3), 16), g = parseInt(hex.slice(3, 5), 16), b = parseInt(hex.slice(5, 7), 16);
    const f = (v) => Math.max(0, Math.min(255, Math.round(v)));
    if (mods.lumMod != null) { r *= mods.lumMod; g *= mods.lumMod; b *= mods.lumMod; }
    if (mods.lumOff != null) { r += 255 * mods.lumOff; g += 255 * mods.lumOff; b += 255 * mods.lumOff; }
    if (mods.shade != null) { r *= mods.shade; g *= mods.shade; b *= mods.shade; }
    if (mods.tint != null) { r = 255 - (255 - r) * mods.tint; g = 255 - (255 - g) * mods.tint; b = 255 - (255 - b) * mods.tint; }
    if (mods.alpha != null && mods.alpha < 1) return `rgba(${f(r)},${f(g)},${f(b)},${mods.alpha})`;
    return `rgb(${f(r)},${f(g)},${f(b)})`;
  }

  function colorOf(el, theme) {
    // el: 包含 srgbClr / schemeClr / prstClr 的节点
    if (!el) return null;
    const mods = {};
    const readMods = (c) => {
      for (const m of c.children) {
        const v = parseInt(m.getAttribute('val'), 10);
        if (Number.isNaN(v)) continue;
        if (m.localName === 'lumMod') mods.lumMod = v / 100000;
        else if (m.localName === 'lumOff') mods.lumOff = v / 100000;
        else if (m.localName === 'shade') mods.shade = v / 100000;
        else if (m.localName === 'tint') mods.tint = v / 100000;
        else if (m.localName === 'alpha') mods.alpha = v / 100000;
      }
    };
    let c = findDeep(el, 'srgbClr');
    if (c) { readMods(c); return applyLum('#' + c.getAttribute('val'), mods); }
    c = findDeep(el, 'schemeClr');
    if (c) {
      readMods(c);
      let name = c.getAttribute('val');
      if (name === 'tx1') name = 'dk1';
      if (name === 'tx2') name = 'dk2';
      if (name === 'bg1') name = 'lt1';
      if (name === 'bg2') name = 'lt2';
      const base = theme[name] || '#000000';
      return applyLum(base, mods);
    }
    c = findDeep(el, 'prstClr');
    if (c) { readMods(c); return c.getAttribute('val') || null; }
    return null;
  }

  /* ---------- 解析 ---------- */

  async function parsePptx(arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const pres = await xmlOf(zip, 'ppt/presentation.xml');
    if (!pres) throw new Error('不是有效的 .pptx 文件(未找到 ppt/presentation.xml)');
    const presRels = await relsOf(zip, 'ppt/presentation.xml');

    const sldSz = pres.getElementsByTagName('p:sldSz')[0];
    const cx = sldSz ? parseInt(sldSz.getAttribute('cx'), 10) : 12192000;
    const cy = sldSz ? parseInt(sldSz.getAttribute('cy'), 10) : 6858000;

    const slidePaths = [];
    const lst = pres.getElementsByTagName('p:sldIdLst')[0];
    if (lst) for (const s of lst.getElementsByTagName('p:sldId')) {
      const rid = s.getAttribute('r:id');
      if (presRels[rid]) slidePaths.push(presRels[rid].target);
    }

    // 主题
    let theme = {};
    for (const k in presRels) {
      if (presRels[k].type.includes('theme')) {
        theme = parseTheme(await xmlOf(zip, presRels[k].target));
        break;
      }
    }

    return { zip, slidePaths, cx, cy, theme };
  }

  async function mediaBitmap(zip, path) {
    const f = zip.file(path);
    if (!f) return null;
    const blob = await f.async('blob');
    try { return await createImageBitmap(blob); }
    catch (e) {
      return await new Promise((res) => {
        const img = new Image();
        img.onload = () => res(img);
        img.onerror = () => res(null);
        img.src = URL.createObjectURL(blob);
      });
    }
  }

  function xfrmOf(sp) {
    const x = findDeep(sp, 'xfrm');
    if (!x) return null;
    const off = tag(x, 'off'), ext = tag(x, 'ext');
    if (!off || !ext) return null;
    return {
      x: parseInt(off.getAttribute('x'), 10) || 0,
      y: parseInt(off.getAttribute('y'), 10) || 0,
      w: parseInt(ext.getAttribute('cx'), 10) || 0,
      h: parseInt(ext.getAttribute('cy'), 10) || 0,
      rot: (parseInt(x.getAttribute('rot'), 10) || 0) / 60000,
    };
  }

  /* ---------- 渲染 ---------- */

  const FONT_STACK = '-apple-system, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "Noto Sans CJK SC", "Noto Sans SC", "DejaVu Sans", sans-serif';

  function phType(sp) {
    const ph = findDeep(sp, 'ph');
    return ph ? (ph.getAttribute('type') || 'body') : null;
  }

  function defaultSizeFor(ph) {
    if (ph === 'title' || ph === 'ctrTitle') return 36;
    if (ph === 'subTitle') return 22;
    return 18;
  }

  async function renderSlide(doc, slidePath, model, canvas) {
    const { zip, theme } = model;
    const rels = await relsOf(zip, slidePath);
    const px = (emu) => emu / EMU_PER_PX * SCALE;
    const W = canvas.width, H = canvas.height;
    const ctx = canvas.getContext('2d');

    // 背景:slide -> layout -> master
    let bgColor = null, bgImage = null;
    let layoutPath = null, masterPath = null;
    for (const k in rels) if (rels[k].type.includes('slideLayout')) layoutPath = rels[k].target;
    let layoutRels = null, masterRels = null, layoutXml = null, masterXml = null;
    if (layoutPath) {
      layoutRels = await relsOf(zip, layoutPath);
      layoutXml = await xmlOf(zip, layoutPath);
      for (const k in layoutRels) if (layoutRels[k].type.includes('slideMaster')) masterPath = layoutRels[k].target;
    }
    if (masterPath) { masterRels = await relsOf(zip, masterPath); masterXml = await xmlOf(zip, masterPath); }

    const bgOf = async (xml, r) => {
      if (!xml) return null;
      const bg = xml.getElementsByTagName('p:bg')[0];
      if (!bg) return null;
      const blip = findDeep(bg, 'blip');
      if (blip && r) {
        const rid = blip.getAttribute('r:embed');
        if (rid && r[rid]) {
          const bmp = await mediaBitmap(zip, r[rid].target);
          if (bmp) return { image: bmp };
        }
      }
      const col = colorOf(findDeep(bg, 'solidFill') || bg, theme);
      if (col) return { color: col };
      return null;
    };

    let bg = await bgOf(doc, rels);
    if (!bg) bg = await bgOf(layoutXml, layoutRels);
    if (!bg) bg = await bgOf(masterXml, masterRels);
    if (bg && bg.image) bgImage = bg.image;
    else if (bg && bg.color) bgColor = bg.color;

    ctx.fillStyle = bgColor || '#ffffff';
    ctx.fillRect(0, 0, W, H);
    if (bgImage) ctx.drawImage(bgImage, 0, 0, W, H);

    // 背景亮度 -> 默认文字颜色
    let bgLum = 1;
    if (bgColor) {
      const m = bgColor.match(/\d+/g);
      if (m) bgLum = (0.299 * m[0] + 0.587 * m[1] + 0.114 * m[2]) / 255;
      else if (bgColor[0] === '#') {
        bgLum = (0.299 * parseInt(bgColor.slice(1, 3), 16) + 0.587 * parseInt(bgColor.slice(3, 5), 16) + 0.114 * parseInt(bgColor.slice(5, 7), 16)) / 255;
      }
    } else if (bgImage) bgLum = 0.6;
    const defaultText = bgLum > 0.45 ? (theme.dk1 || '#111111') : (theme.lt1 || '#ffffff');

    const spTree = doc.getElementsByTagName('p:spTree')[0];
    if (spTree) await renderTree(spTree, { zip, rels, theme, ctx, px, defaultText, W, H }, { dx: 0, dy: 0 });
  }

  async function renderTree(tree, env, off) {
    for (const el of tree.children) {
      if (el.localName === 'sp') await renderShape(el, env, off);
      else if (el.localName === 'pic') await renderPic(el, env, off);
      else if (el.localName === 'grpSp') {
        const x = xfrmOf(el);
        let d = { ...off };
        if (x) {
          const chOff = findDeep(el, 'chOff');
          const cox = chOff ? parseInt(chOff.getAttribute('x'), 10) || 0 : 0;
          const coy = chOff ? parseInt(chOff.getAttribute('y'), 10) || 0 : 0;
          d = { dx: off.dx + x.x - cox, dy: off.dy + x.y - coy };
        }
        await renderTree(el, env, d);
      } else if (el.localName === 'graphicFrame') {
        await renderGraphicFrame(el, env, off);
      }
    }
  }

  async function renderPic(pic, env, off) {
    const { zip, rels, ctx, px } = env;
    const x = xfrmOf(pic);
    if (!x) return;
    const blip = findDeep(pic, 'blip');
    if (!blip) return;
    const rid = blip.getAttribute('r:embed');
    if (!rid || !rels[rid]) return;
    const bmp = await mediaBitmap(zip, rels[rid].target);
    if (!bmp) return;
    ctx.save();
    if (x.rot) {
      ctx.translate(px(off.dx + x.x + x.w / 2), px(off.dy + x.y + x.h / 2));
      ctx.rotate(x.rot * Math.PI / 180);
      ctx.drawImage(bmp, -px(x.w) / 2, -px(x.h) / 2, px(x.w), px(x.h));
    } else {
      ctx.drawImage(bmp, px(off.dx + x.x), px(off.dy + x.y), px(x.w), px(x.h));
    }
    ctx.restore();
  }

  function roundRectPath(ctx, x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y, x + w, y, r);
    ctx.closePath();
  }

  async function renderShape(sp, env, off) {
    const { ctx, px, theme } = env;
    const x = xfrmOf(sp);
    if (!x) { await renderText(sp, env, off, null); return; }

    const spPr = findDeep(sp, 'spPr');
    const geom = spPr ? findDeep(spPr, 'prstGeom') : null;
    const shape = geom ? geom.getAttribute('prst') : null;

    let fill = null, line = null, lineW = 1;
    if (spPr) {
      for (const c of spPr.children) {
        if (c.localName === 'solidFill') fill = colorOf(c, theme);
        if (c.localName === 'noFill') fill = null;
        if (c.localName === 'ln') {
          const lf = tag(c, 'solidFill');
          if (lf) line = colorOf(lf, theme);
          const w = parseInt(c.getAttribute('w'), 10);
          if (w) lineW = Math.max(1, px(w));
        }
        if (c.localName === 'blipFill') {
          // 形状图片填充
          const blip = findDeep(c, 'blip');
          if (blip) {
            const rid = blip.getAttribute('r:embed');
            if (rid && env.rels[rid]) {
              const bmp = await mediaBitmap(env.zip, env.rels[rid].target);
              if (bmp) {
                ctx.drawImage(bmp, px(off.dx + x.x), px(off.dy + x.y), px(x.w), px(x.h));
              }
            }
          }
        }
      }
    }

    const X = px(off.dx + x.x), Y = px(off.dy + x.y), Wd = px(x.w), Ht = px(x.h);
    ctx.save();
    if (x.rot) {
      ctx.translate(X + Wd / 2, Y + Ht / 2);
      ctx.rotate(x.rot * Math.PI / 180);
      ctx.translate(-(X + Wd / 2), -(Y + Ht / 2));
    }
    if (fill || line) {
      if (shape === 'ellipse') {
        ctx.beginPath();
        ctx.ellipse(X + Wd / 2, Y + Ht / 2, Wd / 2, Ht / 2, 0, 0, Math.PI * 2);
      } else if (shape === 'roundRect') {
        roundRectPath(ctx, X, Y, Wd, Ht, Math.min(Wd, Ht) * 0.12);
      } else {
        ctx.beginPath();
        ctx.rect(X, Y, Wd, Ht);
      }
      if (fill) { ctx.fillStyle = fill; ctx.fill(); }
      if (line) { ctx.strokeStyle = line; ctx.lineWidth = lineW; ctx.stroke(); }
    }
    ctx.restore();

    await renderText(sp, env, off, x, fill);
  }

  async function renderText(sp, env, off, x, shapeFill) {
    const { ctx, px, theme, defaultText } = env;
    const tx = findDeep(sp, 'txBody');
    if (!tx || !x) return;

    const bodyPr = tag(tx, 'bodyPr') || findDeep(tx, 'bodyPr');
    let anchor = 't';
    let insL = 91440, insR = 91440, insT = 45720, insB = 45720;
    if (bodyPr) {
      anchor = bodyPr.getAttribute('anchor') || 't';
      const gi = (n, d) => { const v = parseInt(bodyPr.getAttribute(n), 10); return Number.isNaN(v) ? d : v; };
      insL = gi('lIns', insL); insR = gi('rIns', insR); insT = gi('tIns', insT); insB = gi('bIns', insB);
    }

    const ph = phType(sp);
    const defSz = defaultSizeFor(ph);

    // 若形状有填充色,依据填充亮度决定默认文字色
    let defColor = defaultText;
    if (shapeFill) {
      const m = shapeFill.match(/\d+/g);
      if (m && m.length >= 3) {
        const lum = (0.299 * m[0] + 0.587 * m[1] + 0.114 * m[2]) / 255;
        defColor = lum > 0.45 ? '#111111' : '#ffffff';
      }
    }

    const boxX = px(off.dx + x.x + insL), boxW = px(x.w - insL - insR);
    const boxY = px(off.dy + x.y + insT), boxH = px(x.h - insT - insB);
    if (boxW <= 4) return;

    // 收集段落 -> 行
    const paras = [];
    for (const p of tags(tx, 'p')) {
      const pPr = tag(p, 'pPr');
      let algn = 'l', bullet = null, lvl = 0;
      let pDefSz = null, pDefB = false;
      if (pPr) {
        algn = pPr.getAttribute('algn') || 'l';
        lvl = parseInt(pPr.getAttribute('lvl'), 10) || 0;
        if (tag(pPr, 'buChar')) bullet = tag(pPr, 'buChar').getAttribute('char') || '•';
        else if (tag(pPr, 'buAutoNum')) bullet = '#';
        else if (!tag(pPr, 'buNone') && lvl > 0) bullet = '•';
        const dr = tag(pPr, 'defRPr');
        if (dr) {
          const s = parseInt(dr.getAttribute('sz'), 10);
          if (s) pDefSz = s / 100;
          if (dr.getAttribute('b') === '1') pDefB = true;
        }
      }
      const runs = [];
      for (const c of p.children) {
        if (c.localName === 'r') {
          const rPr = tag(c, 'rPr');
          const st = { sz: pDefSz || defSz, b: pDefB, i: false, u: false, color: null };
          if (rPr) {
            const s = parseInt(rPr.getAttribute('sz'), 10);
            if (s) st.sz = s / 100;
            if (rPr.getAttribute('b') === '1') st.b = true;
            if (rPr.getAttribute('i') === '1') st.i = true;
            if (rPr.getAttribute('u') && rPr.getAttribute('u') !== 'none') st.u = true;
            const sf = tag(rPr, 'solidFill');
            if (sf) st.color = colorOf(sf, theme);
          }
          const t = tag(c, 't');
          runs.push({ text: t ? t.textContent : '', st });
        } else if (c.localName === 'br') {
          runs.push({ br: true });
        } else if (c.localName === 'fld') {
          const t = tag(c, 't');
          if (t) runs.push({ text: t.textContent, st: { sz: pDefSz || defSz } });
        }
      }
      paras.push({ algn, bullet, lvl, runs, defSz: pDefSz || defSz });
    }

    // 排版
    const seg = (text) => {
      const out = [];
      const re = /([一-鿿　-〿＀-￯぀-ヿ])|(\s+)|([^\s一-鿿　-〿＀-￯぀-ヿ]+)/g;
      let m;
      while ((m = re.exec(text))) out.push(m[0]);
      return out;
    };
    const fontOf = (st) => `${st.i ? 'italic ' : ''}${st.b ? 'bold ' : ''}${st.sz * SCALE * (96 / 72)}px ${FONT_STACK}`;

    const lines = [];
    let autoNum = 0;
    for (const p of paras) {
      autoNum = p.bullet === '#' ? autoNum + 1 : 0;
      const indent = px(p.lvl * 342900) + (p.bullet ? 18 * SCALE : 0);
      const avail = boxW - indent;
      let cur = [], curW = 0, maxSz = p.defSz;
      const flush = (last) => {
        lines.push({ parts: cur, width: curW, maxSz, algn: p.algn, indent, bullet: null });
        cur = []; curW = 0;
      };
      let firstLine = true;
      const pushPart = (part, w) => {
        if (curW + w > avail && cur.length) {
          lines.push({ parts: cur, width: curW, maxSz, algn: p.algn, indent, bullet: firstLine ? bulletChar() : null });
          firstLine = false;
          cur = []; curW = 0;
        }
        cur.push(part); curW += w;
      };
      const bulletChar = () => {
        if (!p.bullet) return null;
        return p.bullet === '#' ? `${autoNum}.` : p.bullet;
      };
      if (!p.runs.length) {
        lines.push({ parts: [], width: 0, maxSz: p.defSz, algn: p.algn, indent, bullet: null, empty: true });
        continue;
      }
      for (const r of p.runs) {
        if (r.br) {
          lines.push({ parts: cur, width: curW, maxSz, algn: p.algn, indent, bullet: firstLine ? bulletChar() : null });
          firstLine = false; cur = []; curW = 0;
          continue;
        }
        if (r.st.sz > maxSz) maxSz = r.st.sz;
        ctx.font = fontOf(r.st);
        for (const s of seg(r.text)) {
          const w = ctx.measureText(s).width;
          pushPart({ text: s, st: r.st, w }, w);
        }
      }
      if (cur.length || firstLine) {
        lines.push({ parts: cur, width: curW, maxSz, algn: p.algn, indent, bullet: firstLine ? bulletChar() : null });
      }
    }

    const lineHeights = lines.map(l => l.maxSz * SCALE * (96 / 72) * (l.empty ? 0.6 : 1.35));
    const totalH = lineHeights.reduce((a, b) => a + b, 0);
    let ty = boxY;
    if (anchor === 'ctr') ty = boxY + Math.max(0, (boxH - totalH) / 2);
    else if (anchor === 'b') ty = boxY + Math.max(0, boxH - totalH);

    for (let i = 0; i < lines.length; i++) {
      const ln = lines[i];
      const lh = lineHeights[i];
      const fontPx = ln.maxSz * SCALE * (96 / 72);
      const baseline = ty + fontPx * 1.0;
      let dx = boxX + ln.indent;
      if (ln.algn === 'ctr') dx = boxX + ln.indent + (boxW - ln.indent - ln.width) / 2;
      else if (ln.algn === 'r') dx = boxX + (boxW - ln.width);

      if (ln.bullet) {
        ctx.font = `${ln.maxSz * SCALE * (96 / 72) * 0.9}px ${FONT_STACK}`;
        ctx.fillStyle = defColor;
        ctx.fillText(ln.bullet, boxX + ln.indent - 16 * SCALE, baseline);
      }
      for (const part of ln.parts) {
        ctx.font = fontOf(part.st);
        ctx.fillStyle = part.st.color || defColor;
        ctx.fillText(part.text, dx, baseline);
        if (part.st.u) {
          ctx.strokeStyle = part.st.color || defColor;
          ctx.lineWidth = Math.max(1, part.st.sz * SCALE * 0.05);
          ctx.beginPath();
          ctx.moveTo(dx, baseline + 2 * SCALE);
          ctx.lineTo(dx + part.w, baseline + 2 * SCALE);
          ctx.stroke();
        }
        dx += part.w;
      }
      ty += lh;
    }
  }

  async function renderGraphicFrame(gf, env, off) {
    // 表格(a:tbl)的简易渲染
    const { ctx, px, theme, defaultText } = env;
    const x = xfrmOf(gf);
    const tbl = findDeep(gf, 'tbl');
    if (!x || !tbl) return;
    const rows = [];
    for (const tr of tbl.children) {
      if (tr.localName !== 'tr') continue;
      const cells = [];
      for (const tc of tr.children) {
        if (tc.localName !== 'tc') continue;
        let text = '';
        const txb = findDeep(tc, 'txBody');
        if (txb) for (const t of txb.getElementsByTagName('a:t')) text += t.textContent + ' ';
        let fill = null;
        const tcPr = tag(tc, 'tcPr');
        if (tcPr) {
          const sf = tag(tcPr, 'solidFill');
          if (sf) fill = colorOf(sf, theme);
        }
        cells.push({ text: text.trim(), fill });
      }
      if (cells.length) rows.push(cells);
    }
    if (!rows.length) return;
    const X = px(off.dx + x.x), Y = px(off.dy + x.y), W = px(x.w), H = px(x.h);
    const rh = H / rows.length;
    const fontPx = Math.min(16 * SCALE, rh * 0.5);
    for (let r = 0; r < rows.length; r++) {
      const cw = W / rows[r].length;
      for (let c = 0; c < rows[r].length; c++) {
        const cell = rows[r][c];
        if (cell.fill) { ctx.fillStyle = cell.fill; ctx.fillRect(X + c * cw, Y + r * rh, cw, rh); }
        ctx.strokeStyle = 'rgba(128,128,128,.5)';
        ctx.strokeRect(X + c * cw + 0.5, Y + r * rh + 0.5, cw, rh);
        ctx.fillStyle = defaultText;
        ctx.font = `${fontPx}px ${FONT_STACK}`;
        ctx.save();
        ctx.beginPath();
        ctx.rect(X + c * cw, Y + r * rh, cw, rh);
        ctx.clip();
        ctx.fillText(cell.text, X + c * cw + 6 * SCALE, Y + r * rh + rh / 2 + fontPx * 0.35);
        ctx.restore();
      }
    }
  }

  /** pptx -> canvas 数组 */
  async function renderPptx(arrayBuffer, onProgress) {
    const model = await parsePptx(arrayBuffer);
    const W = Math.round(model.cx / EMU_PER_PX * SCALE);
    const H = Math.round(model.cy / EMU_PER_PX * SCALE);
    const canvases = [];
    for (let i = 0; i < model.slidePaths.length; i++) {
      const doc = await xmlOf(model.zip, model.slidePaths[i]);
      const canvas = document.createElement('canvas');
      canvas.width = W; canvas.height = H;
      if (doc) await renderSlide(doc, model.slidePaths[i], model, canvas);
      canvases.push(canvas);
      if (onProgress) onProgress(i + 1, model.slidePaths.length);
    }
    return { canvases, widthPt: model.cx / 914400 * 72, heightPt: model.cy / 914400 * 72 };
  }

  global.LitePptx = { renderPptx };
})(window);
