/* 轻转 LiteConvert — 自研图片编码器 (BMP / TIFF / GIF)
 * 浏览器 canvas 原生只能导出 PNG/JPEG/WebP,其余格式在这里手工编码。
 * MIT License
 */
(function (global) {
  'use strict';

  /** ImageData -> BMP (24-bit, alpha 以白底拍平) */
  function encodeBMP(imageData) {
    const { width: w, height: h, data } = imageData;
    const rowSize = Math.ceil((w * 3) / 4) * 4;
    const pixelBytes = rowSize * h;
    const fileSize = 54 + pixelBytes;
    const buf = new ArrayBuffer(fileSize);
    const v = new DataView(buf);
    // BITMAPFILEHEADER
    v.setUint8(0, 0x42); v.setUint8(1, 0x4d);        // 'BM'
    v.setUint32(2, fileSize, true);
    v.setUint32(10, 54, true);                        // pixel data offset
    // BITMAPINFOHEADER
    v.setUint32(14, 40, true);
    v.setInt32(18, w, true);
    v.setInt32(22, h, true);                          // bottom-up
    v.setUint16(26, 1, true);                         // planes
    v.setUint16(28, 24, true);                        // bpp
    v.setUint32(30, 0, true);                         // BI_RGB
    v.setUint32(34, pixelBytes, true);
    v.setInt32(38, 2835, true);                       // 72 dpi
    v.setInt32(42, 2835, true);
    const out = new Uint8Array(buf);
    for (let y = 0; y < h; y++) {
      let o = 54 + (h - 1 - y) * rowSize;
      let i = y * w * 4;
      for (let x = 0; x < w; x++, i += 4) {
        const a = data[i + 3] / 255;
        out[o++] = Math.round(data[i + 2] * a + 255 * (1 - a)); // B
        out[o++] = Math.round(data[i + 1] * a + 255 * (1 - a)); // G
        out[o++] = Math.round(data[i]     * a + 255 * (1 - a)); // R
      }
    }
    return new Blob([buf], { type: 'image/bmp' });
  }

  /** ImageData -> TIFF (little-endian, RGBA 无压缩单条带) */
  function encodeTIFF(imageData) {
    const { width: w, height: h, data } = imageData;
    const nEntries = 13;
    const ifdSize = 2 + nEntries * 12 + 4;
    const bitsOff = 8 + ifdSize;          // 4 SHORTs
    const xResOff = bitsOff + 8;          // RATIONAL
    const yResOff = xResOff + 8;
    const dataOff = yResOff + 8;
    const byteCount = w * h * 4;
    const buf = new ArrayBuffer(dataOff + byteCount);
    const v = new DataView(buf);
    v.setUint8(0, 0x49); v.setUint8(1, 0x49);   // 'II' little-endian
    v.setUint16(2, 42, true);
    v.setUint32(4, 8, true);                     // IFD offset
    let p = 8;
    v.setUint16(p, nEntries, true); p += 2;
    const entry = (tag, type, count, value) => {
      v.setUint16(p, tag, true);
      v.setUint16(p + 2, type, true);
      v.setUint32(p + 4, count, true);
      if (type === 3 && count === 1) { v.setUint16(p + 8, value, true); v.setUint16(p + 10, 0, true); }
      else v.setUint32(p + 8, value, true);
      p += 12;
    };
    entry(256, 4, 1, w);            // ImageWidth
    entry(257, 4, 1, h);            // ImageLength
    entry(258, 3, 4, bitsOff);      // BitsPerSample 8,8,8,8
    entry(259, 3, 1, 1);            // no compression
    entry(262, 3, 1, 2);            // RGB
    entry(273, 4, 1, dataOff);      // StripOffsets
    entry(277, 3, 1, 4);            // SamplesPerPixel
    entry(278, 4, 1, h);            // RowsPerStrip
    entry(279, 4, 1, byteCount);    // StripByteCounts
    entry(282, 5, 1, xResOff);      // XResolution
    entry(283, 5, 1, yResOff);      // YResolution
    entry(296, 3, 1, 2);            // inch
    entry(338, 3, 1, 2);            // ExtraSamples: unassociated alpha
    v.setUint32(p, 0, true);        // next IFD = none
    for (let i = 0; i < 4; i++) v.setUint16(bitsOff + i * 2, 8, true);
    v.setUint32(xResOff, 72, true); v.setUint32(xResOff + 4, 1, true);
    v.setUint32(yResOff, 72, true); v.setUint32(yResOff + 4, 1, true);
    new Uint8Array(buf, dataOff).set(data);
    return new Blob([buf], { type: 'image/tiff' });
  }

  /* ---------- GIF ---------- */

  /** 中位切分法调色板 (<=256 色) */
  function buildPalette(data, maxColors) {
    const step = Math.max(1, Math.floor(data.length / 4 / 30000)) * 4;
    let px = [];
    for (let i = 0; i < data.length; i += step) {
      if (data[i + 3] >= 128) px.push([data[i], data[i + 1], data[i + 2]]);
    }
    if (!px.length) px.push([255, 255, 255]);
    let boxes = [px];
    while (boxes.length < maxColors) {
      let bi = -1, br = -1;
      for (let i = 0; i < boxes.length; i++) {
        const b = boxes[i];
        if (b.length < 2) continue;
        let mn = [255, 255, 255], mx = [0, 0, 0];
        for (const c of b) for (let k = 0; k < 3; k++) { if (c[k] < mn[k]) mn[k] = c[k]; if (c[k] > mx[k]) mx[k] = c[k]; }
        const range = Math.max(mx[0] - mn[0], mx[1] - mn[1], mx[2] - mn[2]);
        if (range > br) { br = range; bi = i; }
      }
      if (bi < 0 || br === 0) break;
      const box = boxes[bi];
      let mn = [255, 255, 255], mx = [0, 0, 0];
      for (const c of box) for (let k = 0; k < 3; k++) { if (c[k] < mn[k]) mn[k] = c[k]; if (c[k] > mx[k]) mx[k] = c[k]; }
      let ch = 0;
      if (mx[1] - mn[1] >= mx[0] - mn[0] && mx[1] - mn[1] >= mx[2] - mn[2]) ch = 1;
      else if (mx[2] - mn[2] >= mx[0] - mn[0] && mx[2] - mn[2] >= mx[1] - mn[1]) ch = 2;
      box.sort((a, b) => a[ch] - b[ch]);
      const mid = box.length >> 1;
      boxes.splice(bi, 1, box.slice(0, mid), box.slice(mid));
    }
    return boxes.map(b => {
      let r = 0, g = 0, bl = 0;
      for (const c of b) { r += c[0]; g += c[1]; bl += c[2]; }
      const n = b.length || 1;
      return [Math.round(r / n), Math.round(g / n), Math.round(bl / n)];
    });
  }

  function nearestIndex(pal, r, g, b, cache) {
    const key = ((r >> 3) << 10) | ((g >> 3) << 5) | (b >> 3);
    let idx = cache[key];
    if (idx !== undefined) return idx;
    let best = 0, bd = Infinity;
    for (let i = 0; i < pal.length; i++) {
      const dr = pal[i][0] - r, dg = pal[i][1] - g, db = pal[i][2] - b;
      const d = dr * dr + dg * dg + db * db;
      if (d < bd) { bd = d; best = i; }
    }
    cache[key] = best;
    return best;
  }

  /** GIF LZW 压缩(码宽切换时机与 giflib / PIL 解码器约定一致:
   *  发码之后检查 next > maxcode 才加宽,即比"表刚满"晚一个码字) */
  function lzwEncode(minCodeSize, indices) {
    const clear = 1 << minCodeSize, eoi = clear + 1;
    let codeSize = minCodeSize + 1;
    let maxcode = (1 << codeSize) - 1;
    let next = eoi + 1;
    let dict = new Map();
    const bytes = [];
    let cur = 0, curBits = 0;
    const emit = (c) => {
      cur |= c << curBits; curBits += codeSize;
      while (curBits >= 8) { bytes.push(cur & 255); cur >>= 8; curBits -= 8; }
      if (next > maxcode && codeSize < 12) {
        codeSize++;
        maxcode = (1 << codeSize) - 1;
      }
    };
    emit(clear);
    let prefix = indices[0];
    for (let i = 1; i < indices.length; i++) {
      const k = indices[i];
      const key = (prefix << 8) | k;
      if (dict.has(key)) { prefix = dict.get(key); continue; }
      emit(prefix);
      if (next < 4096) {
        dict.set(key, next);
        next++;
      } else {
        emit(clear);
        codeSize = minCodeSize + 1;
        maxcode = (1 << codeSize) - 1;
        next = eoi + 1;
        dict = new Map();
      }
      prefix = k;
    }
    emit(prefix);
    emit(eoi);
    if (curBits > 0) bytes.push(cur & 255);
    return Uint8Array.from(bytes);
  }

  /** ImageData -> GIF (静态单帧, <=256 色, 支持透明) */
  function encodeGIF(imageData) {
    const { width: w, height: h, data } = imageData;
    let hasAlpha = false;
    for (let i = 3; i < data.length; i += 4) if (data[i] < 128) { hasAlpha = true; break; }
    const maxColors = hasAlpha ? 255 : 256;
    const pal = buildPalette(data, maxColors);
    const transIndex = hasAlpha ? pal.length : -1;
    const palSize = hasAlpha ? pal.length + 1 : pal.length;
    let gctBits = 1;
    while ((1 << gctBits) < palSize) gctBits++;
    if (gctBits < 2) gctBits = 2;
    const gctLen = 1 << gctBits;

    const cache = {};
    const indices = new Uint8Array(w * h);
    for (let i = 0, j = 0; i < data.length; i += 4, j++) {
      indices[j] = (hasAlpha && data[i + 3] < 128)
        ? transIndex
        : nearestIndex(pal, data[i], data[i + 1], data[i + 2], cache);
    }

    const minCodeSize = Math.max(2, gctBits);
    const lzw = lzwEncode(minCodeSize, indices);

    const head = new Uint8Array(13 + gctLen * 3 + (transIndex >= 0 ? 8 : 0) + 10 + 1);
    let p = 0;
    const put = (...bs) => { for (const b of bs) head[p++] = b; };
    put(0x47, 0x49, 0x46, 0x38, 0x39, 0x61);                    // GIF89a
    put(w & 255, w >> 8, h & 255, h >> 8);
    put(0x80 | 0x70 | (gctBits - 1), 0, 0);                      // GCT flag + size
    for (let i = 0; i < gctLen; i++) {
      const c = i < pal.length ? pal[i] : [0, 0, 0];
      put(c[0], c[1], c[2]);
    }
    if (transIndex >= 0) put(0x21, 0xF9, 0x04, 0x01, 0, 0, transIndex, 0x00);
    put(0x2C, 0, 0, 0, 0, w & 255, w >> 8, h & 255, h >> 8, 0);  // image descriptor
    put(minCodeSize);

    const parts = [head.subarray(0, p)];
    for (let i = 0; i < lzw.length; i += 255) {
      const chunk = lzw.subarray(i, Math.min(i + 255, lzw.length));
      const blk = new Uint8Array(chunk.length + 1);
      blk[0] = chunk.length;
      blk.set(chunk, 1);
      parts.push(blk);
    }
    parts.push(Uint8Array.from([0x00, 0x3B]));
    return new Blob(parts, { type: 'image/gif' });
  }

  global.LiteEncoders = { encodeBMP, encodeTIFF, encodeGIF };
})(window);
