/*
 * The Ultraspeaker Check Lab — analyze-key.js
 * Reads an Apple Keynote .key file (single-file format: a ZIP archive) in the browser.
 * Apple does not document the format. What is read:
 *  - Data/      every image, video and audio file (real files: reliable)
 *  - Index/*.iwa Snappy-compressed protobuf, scanned for font names and slide count (estimate)
 *  - preview.jpg slide aspect ratio
 * © Nicholas D. Charles — The Ultraspeaker
 */
(function (g) {
  'use strict';
  var CL = g.CheckLab = g.CheckLab || {};

  /* .iwa = chunks of [0x00, 3-byte length, raw Snappy data]; returns the decompressed bytes */
  function iwa(src) {
    var parts = [], o = 0;
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
      parts.push(out); o = end;
    }
    return parts;
  }

  /* Protobuf strings that look like PostScript font names ("HelveticaNeue-Bold") */
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

  async function analyze(zip, opts) {
    var token = opts.token || {};
    function check() { if (token.cancelled) throw new CL.CancelError(); }
    function progress(step, frac) { if (opts.onProgress) opts.onProgress(step, frac); }
    var res = { type: 'key', fileName: opts.fileName, fileBytes: opts.fileBytes, slideW: 0, slideH: 0, slides: [], fonts: [], embeddedFonts: [], media: [], images: [], links: [] };
    var names = Object.keys(zip.files), iwas = names.filter(function (n) { return /^Index\/.*\.iwa$/.test(n); });
    if (!iwas.length) throw new Error('notKey');

    // One Slide-*.iwa per slide (master slides are TemplateSlide-*)
    iwas.forEach(function (n) { if (/\/Slide[-\d]*\.iwa$/.test(n)) res.slides.push({ index: res.slides.length + 1, hidden: null, empty: null, notes: null }); });

    var prev = zip.file('preview.jpg');
    if (prev) { var p = CL.probe.image(await prev.async('uint8array'), 'preview.jpg'); res.slideW = p.width; res.slideH = p.height; }

    // Fonts (estimate) from styles and slides; "HelveticaNeue" -> "Helvetica Neue"
    var found = {}, errors = 0;
    for (var i = 0; i < iwas.length; i++) {
      check();
      progress('fonts', i / iwas.length);
      try { iwa(await zip.file(iwas[i]).async('uint8array')).forEach(function (d) { fontStrings(d, found); }); } catch (e) { errors++; }
    }
    res.fontsUnreadable = errors === iwas.length;
    Object.keys(found).forEach(function (ps) {
      res.fonts.push({ name: ps.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/MT$/, ' MT'), slides: [] });
    });

    // Images, videos and audio: the real files in Data/ (Keynote's own thumbnails skipped)
    // Videos and audio first, then images, so the progress steps run in order
    var data = names.filter(function (n) { return /^Data\/[^/]+$/.test(n) && !/(^|-)small(-|\.)|thumbnail|^Data\/st-|preview/i.test(n); });
    var isImage = function (n) { return CL.probe.kindOf(CL.probe.sniff([], n)) === 'image'; };
    data.sort(function (a, b) { return isImage(a) - isImage(b); });
    var nMedia = data.filter(function (n) { return !isImage(n); }).length;
    for (var j = 0; j < data.length; j++) {
      check();
      progress(j < nMedia ? 'media' : 'images', j < nMedia ? j / nMedia : (j - nMedia) / (data.length - nMedia));
      var name = data[j].slice(5), bytes = await zip.file(data[j]).async('uint8array');
      var kind = CL.probe.kindOf(CL.probe.sniff(bytes, name));
      if (kind === 'image') {
        res.images.push({ path: data[j], name: name, info: CL.probe.image(bytes, name), uses: [] });
      } else if (kind !== 'other') {
        var m = CL.probe.media(bytes, name);
        res.media.push({ path: data[j], name: name, slides: [], bytes: bytes.length, info: m, kind: m.kind });
      }
    }
    return res;
  }

  CL.analyzeKey = analyze;
})(typeof window !== 'undefined' ? window : globalThis);
