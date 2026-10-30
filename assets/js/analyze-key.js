/*
 * The Ultraspeaker Check Lab — analyze-key.js
 * Reads an Apple Keynote .key file (single-file format: a ZIP archive) in the browser and
 * collects the same facts as analyze-pptx.js, slide by slide.
 *  - Index/*.iwa   Snappy-compressed protobuf objects: slide order, skipped slides, notes, the text
 *                  of every object with its font, pictures and movies with their size on the slide
 *  - Index/Metadata.iwa  which file in Data/ belongs to which picture or movie
 *  - Data/         the real image, video and audio files
 * Apple does not document the format: if it cannot be read, the old estimate is used (res.limited).
 * Field numbers follow Keynote's own message definitions (KN, TSD, TSWP, TSS, TSP).
 * © Nicholas D. Charles — The Ultraspeaker
 */
(function (g) {
  'use strict';
  var CL = g.CheckLab = g.CheckLab || {};

  /* .iwa = chunks of [0x00, 3-byte length, raw Snappy data]; returns the decompressed stream */
  function iwa(src) {
    var parts = [], o = 0, total = 0;
    while (o + 4 <= src.length && src[o] === 0) {
      var end = o + 4 + (src[o + 1] | (src[o + 2] << 8) | (src[o + 3] << 16)), p = o + 4;
      var len = 0, shift = 0, c;
      do { c = src[p++]; len += (c & 0x7F) * Math.pow(2, shift); shift += 7; } while (c & 0x80);
      var out = new Uint8Array(len), op = 0;
      while (p < end && op < len) {
        var tag = src[p++], type = tag & 3, l, off;
        if (type === 0) {
          l = tag >>> 2;
          if (l >= 60) { var nb = l - 59; l = 0; for (var i = 0; i < nb; i++) l += src[p++] * Math.pow(256, i); }
          out.set(src.subarray(p, p + l + 1), op); p += l + 1; op += l + 1;
          continue;
        }
        if (type === 1) { l = ((tag >>> 2) & 7) + 4; off = ((tag >>> 5) << 8) | src[p++]; }
        else if (type === 2) { l = (tag >>> 2) + 1; off = src[p] | (src[p + 1] << 8); p += 2; }
        else { l = (tag >>> 2) + 1; off = (src[p] | (src[p + 1] << 8) | (src[p + 2] << 16)) + src[p + 3] * 16777216; p += 4; }
        if (!off || off > op) throw new Error('snappy');
        for (var k = 0; k < l; k++, op++) out[op] = out[op - off];
      }
      parts.push(out); total += out.length; o = end;
    }
    var all = new Uint8Array(total), at = 0;
    parts.forEach(function (x) { all.set(x, at); at += x.length; });
    return all;
  }

  /* ---------- minimal protobuf reader ---------- */
  var utf8 = new TextDecoder('utf-8');
  /* A message as { field number: [values] }: numbers for varints and floats, {d, s, e} for bytes */
  function msg(d, s, e) {
    var f = {}, p = s, dv = null;
    while (p < e) {
      var key = 0, mul = 1, c;
      do { c = d[p++]; key += (c & 0x7F) * mul; mul *= 128; } while (c & 0x80 && p < e);
      var n = Math.floor(key / 8), w = key % 8, v;
      if (w === 0) { v = 0; mul = 1; do { c = d[p++]; v += (c & 0x7F) * mul; mul *= 128; } while (c & 0x80); }
      else if (w === 2) { var l = 0; mul = 1; do { c = d[p++]; l += (c & 0x7F) * mul; mul *= 128; } while (c & 0x80); v = { d: d, s: p, e: p + l }; p += l; }
      else if (w === 5) { dv = dv || new DataView(d.buffer, d.byteOffset); v = dv.getFloat32(p, true); p += 4; }
      else if (w === 1) { dv = dv || new DataView(d.buffer, d.byteOffset); v = dv.getFloat64(p, true); p += 8; }
      else throw new Error('protobuf');
      (f[n] = f[n] || []).push(v);
    }
    return f;
  }
  function sub(m, n) { var v = m && m[n] && m[n][0]; return v && v.d ? msg(v.d, v.s, v.e) : null; }
  function subs(m, n) { return ((m && m[n]) || []).filter(function (v) { return v.d; }).map(function (v) { return msg(v.d, v.s, v.e); }); }
  function num(m, n) { var v = m && m[n] && m[n][0]; return typeof v === 'number' ? v : undefined; }
  function str(m, n) { var v = m && m[n] && m[n][0]; return v && v.d ? utf8.decode(v.d.subarray(v.s, v.e)) : ''; }
  function strs(m, n) { return ((m && m[n]) || []).map(function (v) { return utf8.decode(v.d.subarray(v.s, v.e)); }); }
  function ref(m, n) { return num(sub(m, n), 1); }                 // TSP.Reference / DataReference: identifier = 1
  function refs(m, n) { return subs(m, n).map(function (r) { return num(r, 1); }); }

  /* Every object of an .iwa stream: id -> { type, its first message, the Data/ files it uses } */
  function objects(d, into) {
    var p = 0;
    while (p < d.length) {
      var l = 0, mul = 1, c;
      do { c = d[p++]; l += (c & 0x7F) * mul; mul *= 128; } while (c & 0x80);
      var info = msg(d, p, p + l), id = num(info, 1), first = null, data = [];
      p += l;
      subs(info, 2).forEach(function (mi) {                       // MessageInfo: type 1, length 3, data references 6
        var len = num(mi, 3) || 0;
        if (!first) first = { type: num(mi, 1), d: d, s: p, e: p + len, data: data, m: null };
        data.push.apply(data, mi[6] || []);
        p += len;
      });
      if (id !== undefined && first) into[id] = first;
    }
  }

  /* Protobuf strings that look like PostScript font names ("HelveticaNeue-Bold"): the old estimate */
  var HINT = /(-|^)(Regular|Roman|Book|Light|Thin|Medium|SemiBold|Semibold|DemiBold|Bold|ExtraBold|Heavy|Black|Italic|Oblique|BoldItalic|LightItalic|Condensed|Std|Pro|LT)$|MT$/;
  var NOT_FONT = /^(Regular|Bold|Italic|Light|Medium|Black|Title|Body|Heading|Caption|Bullet|Subtitle|Quote|Attribution|Label|Default|None|Normal)$/i;
  function fontStrings(d, found) {
    for (var i = 1; i < d.length - 3; i++) {
      var L = d[i];
      if (L < 3 || L > 63 || (d[i - 1] & 7) !== 2 || i + 1 + L > d.length) continue;
      var s = String.fromCharCode.apply(null, d.subarray(i + 1, i + 1 + L));
      if (!/^[A-Za-z][A-Za-z0-9]*(-[A-Za-z0-9]+)?$/.test(s) || NOT_FONT.test(s)) continue;
      if (HINT.test(s) || CL.fonts.classify(s) !== 'custom') found[s.split('-')[0]] = 1;
    }
  }
  /* "HelveticaNeue-Bold" -> "Helvetica Neue"; Apple's hidden system font names -> "SF Pro" */
  function family(ps) {
    if (/^\./.test(ps) || /^SFNS|^SFPro/.test(ps)) return 'SF Pro';
    return ps.split('-')[0].replace(/([a-z])([A-Z])/g, '$1 $2').replace(/MT$/, ' MT').trim();
  }

  // Message type numbers (Keynote's registry)
  var T = { DOC: 1, PH: 7, PH2: 12, STORAGE: 2001, STORAGE2: 2005, SHAPEINFO: 2011, CHARSTYLE: 2021, PARASTYLE: 2022, LINK: 2032,
    IMAGE: 3005, MOVIE: 3007, GROUP: 3008, META: 11006 };

  function readStructure(objs, res) {
    function get(id) { var o = objs[id]; if (o && !o.m) o.m = msg(o.d, o.s, o.e); return o; }
    function find(type) { for (var id in objs) if (objs[id].type === type) return get(id); return null; }

    // Data/ files by data identifier
    var meta = find(T.META), files = {};
    if (!meta) throw new Error('no metadata');
    subs(meta.m, 4).forEach(function (di) {
      var file = str(di, 4) || str(di, 3);
      if (file) files[num(di, 1)] = { path: 'Data/' + file, name: str(di, 3) || file };
    });

    // Slide size and order: the slide tree, depth first; skipped = hidden
    var doc = find(T.DOC), show = doc && get(ref(doc.m, 2));
    if (!show) throw new Error('no show');
    var size = sub(show.m, 4), tree = sub(show.m, 3);
    res.slideW = num(size, 1) || 0; res.slideH = num(size, 2) || 0;
    var order = [], seen = {};
    (function walk(nodes) {                                       // SlideNode: slide 2, skipped 4, indented children 1
      nodes.forEach(function (id) {
        var node = get(id), slideId = node && ref(node.m, 2);
        if (!node || seen[id]) return;
        seen[id] = 1;
        if (slideId) order.push({ slide: slideId, hidden: num(node.m, 4) === 1 });
        walk(refs(node.m, 1));
      });
    })(tree[2] ? refs(tree, 2) : [ref(tree, 1)]);                // top-level slides, or the root node of older files
    if (!order.length) throw new Error('no slides');

    // Font of a style: its own font or the nearest parent's (TSS.StyleArchive.parent)
    var fontCache = {};
    function styleFont(id) {
      if (id in fontCache) return fontCache[id];
      var f = null, seen = 0, cur = id;
      while (cur && seen++ < 30) {
        var st = get(cur);
        if (!st || (st.type !== T.CHARSTYLE && st.type !== T.PARASTYLE)) break;
        var name = str(sub(st.m, 11), 5);                       // char_properties.font_name
        if (name) { f = name; break; }
        cur = ref(sub(st.m, 1), 3);                             // super.parent
      }
      return (fontCache[id] = f);
    }
    function table(storage, n) {                                  // ObjectAttributeTable: [{at, id}]
      return subs(sub(storage, n), 1).map(function (e) { return { at: num(e, 1) || 0, id: ref(e, 2) }; }).sort(function (a, b) { return a.at - b.at; });
    }
    function clean(s) { return s.replace(/[\u0000-\u001F\u2028\u2029\uFFFC]/g, ' ').replace(/\s+/g, ' ').trim(); }
    /* Text of a storage, cut by font: { font: text } */
    function textByFont(storageId) {
      var st = get(storageId), out = {};
      if (!st || (st.type !== T.STORAGE && st.type !== T.STORAGE2)) return out;
      var text = strs(st.m, 3).join(''), paras = table(st.m, 5), chars = table(st.m, 8);
      var cuts = [0, text.length].concat(paras.map(function (x) { return x.at; }), chars.map(function (x) { return x.at; }))
        .filter(function (v, i, a) { return v <= text.length && a.indexOf(v) === i; }).sort(function (a, b) { return a - b; });
      function at(list, i) { var r = null; list.forEach(function (x) { if (x.at <= i) r = x; }); return r && r.id; }
      for (var i = 0; i < cuts.length - 1; i++) {
        var piece = text.slice(cuts[i], cuts[i + 1]), cs = at(chars, cuts[i]), ps = at(paras, cuts[i]);
        var font = (cs && styleFont(cs)) || (ps && styleFont(ps));
        if (font && clean(piece)) out[font] = (out[font] || '') + piece;
      }
      return out;
    }
    function storageText(id) { var st = get(id); return st && (st.type === T.STORAGE || st.type === T.STORAGE2) ? clean(strs(st.m, 3).join('')) : ''; }
    function links(storageId, slide) {
      var st = get(storageId);
      if (st) table(st.m, 11).forEach(function (x) {
        var f = get(x.id), url = f && f.type === T.LINK ? str(f.m, 2) : '';
        if (/^(https?|file):/i.test(url)) res.links.push({ slide: slide, kind: /^https?:/i.test(url) ? 'web' : 'file', target: url });
      });
    }
    // DrawableArchive (the "super" chain differs by type): geometry, name for the reader
    function drawable(o) {
      var t = o.type, m = o.m;
      if (t === T.PH || t === T.PH2) m = sub(sub(sub(m, 1), 1), 1);        // Placeholder > ShapeInfo > Shape > Drawable
      else if (t === T.SHAPEINFO) m = sub(sub(m, 1), 1);                   // ShapeInfo > Shape > Drawable
      else m = sub(m, 1);                                                  // Shape, Image, Movie, Group, Mask > Drawable
      var geo = sub(m, 1), sz = sub(geo, 2), pos = sub(geo, 1);
      return { w: num(sz, 1) || 0, h: num(sz, 2) || 0, x: num(pos, 1) || 0, y: num(pos, 2) || 0, label: str(m, 8), url: str(m, 4) };
    }

    var fonts = {}, mediaById = {}, imageUses = {};
    function picture(file) { return (imageUses[file.path] = imageUses[file.path] || { path: file.path, name: file.name, uses: [] }).uses; }
    function addFont(ps, use) {
      var name = family(ps), k = name.toLowerCase();
      (fonts[k] = fonts[k] || { name: name, uses: [] }).uses.push(use);
    }

    /* One object shown on a slide: returns true when it puts something on the screen */
    function visit(id, sn, kindHint, fromMaster, depth) {
      var o = get(id);
      if (!o || depth > 20) return false;
      var t = o.type, d, content = false;
      if (t === T.GROUP) {
        refs(o.m, 2).forEach(function (c) { if (visit(c, sn, null, fromMaster, depth + 1)) content = true; });
        return content;
      }
      if (t === T.PH || t === T.PH2 || t === T.SHAPEINFO) {
        if (fromMaster && t !== T.SHAPEINFO) return false;               // the slide has its own placeholders
        d = drawable(o);
        var info = t === T.SHAPEINFO ? o.m : sub(o.m, 1), storage = ref(info, 4);
        var kind = fromMaster ? 'master' : (kindHint || (num(info, 6) === 1 ? 'textbox' : 'shape'));
        var byFont = textByFont(storage);
        Object.keys(byFont).forEach(function (f) {
          var said = clean(byFont[f]);
          addFont(f, { slide: sn, kind: kind, shape: d.label, text: said.length > 80 ? said.slice(0, 79) + '…' : said });
        });
        links(storage, sn);
        if (d.url && /^https?:/i.test(d.url)) res.links.push({ slide: sn, kind: 'web', target: d.url });
        return t === T.SHAPEINFO || !!storageText(storage);              // an empty placeholder shows nothing
      }
      if (t === T.IMAGE) {
        d = drawable(o);
        var file = files[ref(o.m, 11)] || files[ref(o.m, 13)];
        if (!file) return true;
        var mask = get(ref(o.m, 5)), crop = { l: 0, t: 0, r: 0, b: 0 }, shown = d.w;
        if (mask && d.w && d.h) {                                       // the mask is the visible window, in the picture's frame
          var mk = drawable(mask);
          if (mk.w && mk.h) {
            crop = { l: Math.max(0, mk.x / d.w), t: Math.max(0, mk.y / d.h), r: Math.max(0, 1 - (mk.x + mk.w) / d.w), b: Math.max(0, 1 - (mk.y + mk.h) / d.h) };
            shown = Math.min(d.w, mk.w);
          }
        }
        picture(file).push({ slide: sn, cx: shown, crop: crop, kind: fromMaster ? 'master' : 'image', shape: d.label });
        if (d.url && /^https?:/i.test(d.url)) res.links.push({ slide: sn, kind: 'web', target: d.url });
        return true;
      }
      if (t === T.MOVIE) {
        var mf = files[ref(o.m, 14)];
        if (!mf) return true;
        var m = mediaById[mf.path] = mediaById[mf.path] || { path: mf.path, name: mf.name, slides: [] };
        if (m.slides.indexOf(sn) < 0) m.slides.push(sn);
        return true;
      }
      return true;                                                       // shapes, lines, tables, charts
    }
    function background(styleId, sn) {                                  // pictures used as slide background
      for (var k = 0, cur = styleId; cur && k < 6; k++) {
        var st = get(cur);
        if (!st) return;
        st.data.forEach(function (did) {
          var file = files[did];
          if (file) picture(file).push({ slide: sn, cx: res.slideW, crop: { l: 0, t: 0, r: 0, b: 0 }, kind: 'master' });
        });
        cur = ref(sub(st.m, 1), 3);
      }
    }

    // Objects shown on a slide, in stacking order (field 42). Files saved before Keynote wrote that list only have the owned objects.
    var zOrder = order.some(function (e) { var s = get(e.slide); return s && s.m[42]; });
    function shown(m) { return zOrder ? refs(m, 42) : refs(m, 7); }
    order.forEach(function (entry, i) {
      var sn = i + 1, sl = get(entry.slide), slide = { index: sn, hidden: entry.hidden, notes: false, empty: true };
      res.slides.push(slide);
      if (!sl) return;
      var m = sl.m, visible = shown(m);
      var kinds = {}; kinds[ref(m, 5)] = 'title'; kinds[ref(m, 6)] = 'body'; kinds[ref(m, 20)] = 'slideNumber';
      visible.forEach(function (id) { if (visit(id, sn, kinds[id], false, 0) && kinds[id] !== 'slideNumber') slide.empty = false; });
      var note = get(ref(m, 27));
      slide.notes = !!(note && storageText(ref(note.m, 1)));
      // What the master slide puts on every slide: its own objects (not placeholders) and background
      var tpl = get(ref(m, 17));
      if (tpl) {
        shown(tpl.m).forEach(function (id) { visit(id, sn, null, true, 0); });
        background(ref(tpl.m, 1), sn);
      }
      background(ref(m, 1), sn);
    });

    function values(o) { return Object.keys(o).map(function (k) { return o[k]; }); }
    res.fonts = values(fonts);
    return { media: values(mediaById), images: values(imageUses) };
  }

  /* When the structure cannot be read: the estimate of earlier versions (slide count, fonts named in the styles, every picture in Data/) */
  async function estimate(zip, res, iwas, streams, data) {
    res.limited = true;
    res.slides = []; res.fonts = []; res.links = [];
    iwas.forEach(function (n) { if (/\/Slide[-\d]*\.iwa$/.test(n)) res.slides.push({ index: res.slides.length + 1, hidden: null, empty: null, notes: null }); });
    var prev = zip.file('preview.jpg');
    if (prev) { var p = CL.probe.image(await prev.async('uint8array'), 'preview.jpg'); res.slideW = p.width; res.slideH = p.height; }
    var est = {};
    streams.forEach(function (s) { fontStrings(s, est); });
    res.fontsUnreadable = !streams.length;
    Object.keys(est).forEach(function (ps) { res.fonts.push({ name: family(ps), uses: [] }); });
    return { media: [], images: data.filter(isImage).map(function (n) { return { path: n, name: n.slice(5), uses: [] }; }) };
  }
  function isImage(n) { return CL.probe.kindOf(CL.probe.sniff([], n)) === 'image'; }

  /* ---------- the analysis ---------- */
  async function analyze(zip, opts) {
    function check() { if (opts.token && opts.token.cancelled) throw new Error('cancelled'); }
    var res = { type: 'key', fileName: opts.fileName, fileBytes: opts.fileBytes, slideW: 0, slideH: 0, slides: [], fonts: [], embeddedFonts: [], media: [], images: [], links: [] };
    var names = Object.keys(zip.files), iwas = names.filter(function (n) { return /^Index\/.*\.iwa$/.test(n); });
    if (!iwas.length) throw new Error('notKey');
    // Keynote's own thumbnails and theme previews are skipped
    var data = names.filter(function (n) { return /^Data\/[^/]+$/.test(n) && !/(^|-)small(-|\.)|thumbnail|^Data\/(st|mt)-|preview/i.test(n); });

    // Read every object of the document
    var objs = {}, streams = [];
    for (var i = 0; i < iwas.length; i++) {
      check();
      opts.onProgress('fonts', 0.8 * i / iwas.length);
      try { var d = iwa(await zip.file(iwas[i]).async('uint8array')); streams.push(d); objects(d, objs); } catch (e) { /* skipped */ }
    }
    var found;
    try { found = readStructure(objs, res); } catch (e) { found = await estimate(zip, res, iwas, streams, data); }
    opts.onProgress('fonts', 1);

    // Video and audio: the movies on the slides, plus any other sound or video file in Data/ (soundtrack)
    var media = found.media.concat(data.filter(function (n) { return !isImage(n) && !found.media.some(function (m) { return m.path === n; }); })
      .map(function (n) { return { path: n, name: n.slice(5), slides: [] }; }));
    var mediaDone = CL.probe.byBytes(zip, media);
    for (var j = 0; j < media.length; j++) {
      check();
      opts.onProgress('media', mediaDone(j));
      var mm = media[j], mfile = zip.file(mm.path);
      if (!mfile) continue;
      var bytes = await mfile.async('uint8array'), kind = CL.probe.kindOf(CL.probe.sniff(bytes, mm.path));
      if (kind === 'image' || kind === 'other') continue;
      var info = CL.probe.media(bytes, mm.path);
      res.media.push({ path: mm.path, name: mm.name, slides: mm.slides.sort(function (a, b) { return a - b; }), bytes: bytes.length, info: info, kind: info.kind });
    }

    // Images: those on the slides, with their size and crop
    var imageDone = CL.probe.byBytes(zip, found.images);
    for (var q = 0; q < found.images.length; q++) {
      check();
      opts.onProgress('images', imageDone(q));
      var pic = found.images[q], f = zip.file(pic.path);
      if (!f) continue;
      var ib = await f.async('uint8array');
      if (CL.probe.kindOf(CL.probe.sniff(ib, pic.path)) !== 'image') continue;
      res.images.push({ path: pic.path, name: pic.name, info: CL.probe.image(ib, pic.path), uses: pic.uses });
    }
    return res;
  }

  CL.analyzeKey = analyze;
})(window);
