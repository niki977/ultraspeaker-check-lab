/*
 * The Ultraspeaker Check Lab — probe.js
 * Reads the real format of a media or image file from its bytes, not its extension.
 * © Nicholas D. Charles — The Ultraspeaker
 */
(function (g) {
  'use strict';
  var CL = g.CheckLab = g.CheckLab || {};

  var IMAGES = ['png', 'jpeg', 'gif', 'bmp', 'tiff', 'webp', 'heic', 'avif', 'svg', 'emf', 'wmf', 'eps', 'pdf'];
  var AUDIO = ['mp3', 'wav', 'm4a', 'wma', 'aiff'];
  var VIDEO_FORMATS = ['mp4', 'mov', 'avi', 'wmv', 'mpeg', 'mpg', 'mkv', 'webm', 'flv'];
  var ISO = ['mp4', 'mov', 'm4a'];

  function str(b, o, n) { var s = ''; for (var i = 0; i < n && o + i < b.length; i++) s += String.fromCharCode(b[o + i]); return s; }
  function be16(b, o) { return (b[o] << 8) | b[o + 1]; }
  function be32(b, o) { return b[o] * 16777216 + ((b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]); }
  function le16(b, o) { return b[o] | (b[o + 1] << 8); }
  function le32(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16)) + b[o + 3] * 16777216; }

  /* Format name from the first bytes; the extension is only a fallback. */
  function sniff(b, name) {
    var ext = String(name).split('.').pop().toLowerCase();
    var h4 = str(b, 0, 4);
    if (h4 === '\x89PNG') return 'png';
    if (b[0] === 0xFF && b[1] === 0xD8) return 'jpeg';
    if (h4 === 'GIF8') return 'gif';
    if (h4.slice(0, 2) === 'BM') return 'bmp';
    if (h4 === 'II*\x00' || h4 === 'MM\x00*') return 'tiff';
    if (h4 === 'RIFF') return { WEBP: 'webp', 'AVI ': 'avi', WAVE: 'wav' }[str(b, 8, 4)] || ext;
    if (str(b, 40, 4) === ' EMF') return 'emf';
    if (h4 === '\xD7\xCD\xC6\x9A') return 'wmf';
    if (h4 === '%!PS' || h4 === '\xC5\xD0\xD3\xC6') return 'eps';
    if (h4 === '%PDF') return 'pdf';
    if (h4 === '0&\xB2u') return ext === 'wma' ? 'wma' : 'wmv';
    if (str(b, 4, 4) === 'ftyp') {
      var brand = str(b, 8, 4);
      if (/^(heic|heix|mif1|msf1)$/.test(brand)) return 'heic';
      if (brand === 'avif') return 'avif';
      if (brand === 'qt  ') return 'mov';
      return /^M4[AB] $/.test(brand) || ext === 'm4a' ? 'm4a' : 'mp4';
    }
    if (/^(moov|mdat|wide|free)$/.test(str(b, 4, 4))) return 'mov';
    if (/<svg[\s>]/i.test(str(b, 0, 1024))) return 'svg';
    if (h4.slice(0, 3) === 'ID3') return 'mp3';
    return { jpg: 'jpeg', tif: 'tiff', aif: 'aiff', m4v: 'mp4', '3gp': 'mp4' }[ext] || ext;
  }
  function kindOf(format) {
    return IMAGES.indexOf(format) >= 0 ? 'image' : AUDIO.indexOf(format) >= 0 ? 'audio' : VIDEO_FORMATS.indexOf(format) >= 0 ? 'video' : 'other';
  }

  /* ---------- ISO-BMFF / QuickTime ---------- */
  function boxes(b, start, end) {
    var out = [], o = start;
    while (o + 8 <= end) {
      var size = be32(b, o), hdr = 8;
      if (size === 1) { size = be32(b, o + 8) * 4294967296 + be32(b, o + 12); hdr = 16; }
      else if (size === 0) size = end - o;
      if (size < hdr) break;
      out.push({ type: str(b, o + 4, 4), data: o + hdr, end: Math.min(o + size, end) });
      o += size;
    }
    return out;
  }
  function child(b, box, type) { return boxes(b, box.data, box.end).filter(function (x) { return x.type === type; })[0]; }

  var VIDEO = {
    avc1: 'H.264', avc3: 'H.264', hvc1: 'HEVC (H.265)', hev1: 'HEVC (H.265)', mp4v: 'MPEG-4 Part 2', s263: 'H.263',
    apcn: 'Apple ProRes 422', apch: 'Apple ProRes 422 HQ', apcs: 'Apple ProRes 422 LT', apco: 'Apple ProRes 422 Proxy', ap4h: 'Apple ProRes 4444', ap4x: 'Apple ProRes 4444 XQ',
    'dvc ': 'DV', dvcp: 'DV', dvpp: 'DV', dv5n: 'DV', dv5p: 'DV', dvh5: 'DVCPRO HD', dvh6: 'DVCPRO HD', hdv1: 'HDV', hdv2: 'HDV',
    jpeg: 'Motion JPEG', mjpa: 'Motion JPEG', mjpb: 'Motion JPEG', 'raw ': 'Uncompressed', 'rle ': 'QuickTime Animation', 'png ': 'PNG sequence', cvid: 'Cinepak', SVQ3: 'Sorenson 3'
  };
  var AUDIO_CODEC = { mp4a: 'AAC', alac: 'Apple Lossless', sowt: 'PCM', twos: 'PCM', lpcm: 'PCM', 'ac-3': 'Dolby AC-3', '.mp3': 'MP3' };

  function probeIso(b, info) {
    var moov = boxes(b, 0, b.length).filter(function (x) { return x.type === 'moov'; })[0];
    if (!moov) { info.incomplete = true; return; }
    var mvhd = child(b, moov, 'mvhd');
    if (mvhd) {
      var v1 = b[mvhd.data] === 1, ts = be32(b, mvhd.data + (v1 ? 20 : 12));
      var dur = v1 ? be32(b, mvhd.data + 24) * 4294967296 + be32(b, mvhd.data + 28) : be32(b, mvhd.data + 16);
      if (ts) info.duration = dur / ts;
    }
    boxes(b, moov.data, moov.end).forEach(function (trak) {
      if (trak.type !== 'trak') return;
      var mdia = child(b, trak, 'mdia'), hdlr = mdia && child(b, mdia, 'hdlr');
      var minf = mdia && child(b, mdia, 'minf'), stbl = minf && child(b, minf, 'stbl'), stsd = stbl && child(b, stbl, 'stsd');
      if (!hdlr || !stsd) return;
      var handler = str(b, hdlr.data + 8, 4), e = stsd.data + 8, fourcc = str(b, e + 4, 4);
      if (handler === 'vide' && !info.video) {
        info.video = { fourcc: fourcc, codec: VIDEO[fourcc] || fourcc.trim() };
        info.width = be16(b, e + 32); info.height = be16(b, e + 34);
      } else if (handler === 'soun' && !info.audio) {
        info.audio = { codec: AUDIO_CODEC[fourcc] || fourcc.trim() };
      }
    });
    info.kind = info.video ? 'video' : 'audio';
  }

  function media(bytes, name) {
    var b = new Uint8Array(bytes), f = sniff(b, name);
    var info = { format: f, kind: kindOf(f) === 'audio' ? 'audio' : 'video', video: null, audio: null, width: 0, height: 0, duration: 0 };
    if (ISO.indexOf(f) >= 0) probeIso(b, info);
    else if (info.kind === 'video') info.video = { codec: f === 'wmv' ? 'Windows Media Video' : f.toUpperCase() };
    else info.audio = { codec: f.toUpperCase() };
    return info;
  }

  /* ---------- Images ---------- */
  function exif(b, t) {
    var le = b[t] === 0x49, res = { exif: true, gps: false, make: '', model: '', date: '' };
    function r16(o) { return le ? le16(b, t + o) : be16(b, t + o); }
    function r32(o) { return le ? le32(b, t + o) : be32(b, t + o); }
    function text(count, vo) { var o = count > 4 ? r32(vo) : vo; return str(b, t + o, Math.min(count, 64)).replace(/\x00.*$/, '').trim(); }
    function ifd(off, cb) {
      if (!off || t + off + 2 > b.length) return;
      for (var i = 0, n = Math.min(r16(off), 300); i < n; i++) { var e = off + 2 + i * 12; cb(r16(e), r32(e + 4), e + 8); }
    }
    var sub = 0, gps = 0;
    ifd(r32(4), function (tag, count, vo) {
      if (tag === 0x010F) res.make = text(count, vo);
      else if (tag === 0x0110) res.model = text(count, vo);
      else if (tag === 0x0132) res.date = text(count, vo);
      else if (tag === 0x8769) sub = r32(vo);
      else if (tag === 0x8825) gps = r32(vo);
    });
    ifd(sub, function (tag, count, vo) { if (tag === 0x9003) res.date = text(count, vo); });
    ifd(gps, function (tag) { if (tag === 2 || tag === 4) res.gps = true; });
    return res;
  }

  function image(bytes, name) {
    var b = new Uint8Array(bytes), f = sniff(b, name);
    var info = { format: f, bytes: b.length, width: 0, height: 0, meta: null, cmyk: false, vector: /^(svg|emf|wmf|eps|pdf)$/.test(f) };
    try {
      if (f === 'png') { info.width = be32(b, 16); info.height = be32(b, 20); }
      else if (f === 'gif') { info.width = le16(b, 6); info.height = le16(b, 8); }
      else if (f === 'bmp') { info.width = Math.abs(le32(b, 18) | 0); info.height = Math.abs(le32(b, 22) | 0); }
      else if (f === 'jpeg') {
        for (var o = 2; o + 9 < b.length;) {
          if (b[o] !== 0xFF) { o++; continue; }
          var m = b[o + 1], len = be16(b, o + 2);
          if (m === 0xDA || m === 0xD9) break;
          if (m === 0xE1 && !info.meta && str(b, o + 4, 4) === 'Exif') info.meta = exif(b, o + 10);
          if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) {
            info.height = be16(b, o + 5); info.width = be16(b, o + 7); info.cmyk = b[o + 9] === 4;
          }
          o += 2 + len;
        }
      } else if (f === 'tiff') {
        var le = b[0] === 0x49, r16 = le ? le16 : be16, ifd = (le ? le32 : be32)(b, 4);
        for (var i = 0, n = r16(b, ifd); i < n && i < 200; i++) {
          var e = ifd + 2 + i * 12, tag = r16(b, e), val = r16(b, e + 2) === 3 ? r16(b, e + 8) : (le ? le32 : be32)(b, e + 8);
          if (tag === 256) info.width = val;
          if (tag === 257) info.height = val;
          if (tag === 262 && val === 5) info.cmyk = true;
        }
      } else if (f === 'webp') {
        var c = str(b, 12, 4);
        if (c === 'VP8 ') { info.width = le16(b, 26) & 0x3FFF; info.height = le16(b, 28) & 0x3FFF; }
        else if (c === 'VP8L') { var bits = le32(b, 21); info.width = (bits & 0x3FFF) + 1; info.height = ((bits >> 14) & 0x3FFF) + 1; }
        else { info.width = (b[24] | b[25] << 8 | b[26] << 16) + 1; info.height = (b[27] | b[28] << 8 | b[29] << 16) + 1; }
      } else if (f === 'heic') {
        // largest 'ispe' (image size) box, and an 'Exif' item if present
        var head = str(b, 0, Math.min(b.length, 65536));
        for (var p = head.indexOf('ispe'); p >= 0; p = head.indexOf('ispe', p + 4)) {
          var w = be32(b, p + 8), hh = be32(b, p + 12);
          if (w * hh > info.width * info.height) { info.width = w; info.height = hh; }
        }
        if (head.indexOf('Exif') >= 0) info.meta = { exif: true, gps: null, make: '', model: '', date: '' };
      } else if (f === 'svg') {
        var svg = (/<svg[^>]*>/i.exec(str(b, 0, 4096)) || [''])[0];
        var vb = /viewBox\s*=\s*["'][-\d.\s,]+?([\d.]+)[\s,]+([\d.]+)\s*["']/i.exec(svg);
        var sw = /\swidth\s*=\s*["']([\d.]+)/i.exec(svg), sh = /\sheight\s*=\s*["']([\d.]+)/i.exec(svg);
        info.width = sw ? +sw[1] : (vb ? +vb[1] : 0); info.height = sh ? +sh[1] : (vb ? +vb[2] : 0);
      }
    } catch (err) { /* unreadable header: dimensions stay 0 and the image is skipped */ }
    return info;
  }

  /* Progress through files to decompress, by their size and not their number: share done before file k */
  function byBytes(zip, files) {
    var sizes = files.map(function (x) { var f = x.path && zip.file(x.path); return (f && f._data && f._data.uncompressedSize) || 1; });
    var total = sizes.reduce(function (a, b) { return a + b; }, 0), done = 0;
    return function (k) { done += k > 0 ? sizes[k - 1] : 0; return total ? done / total : 1; };
  }

  CL.probe = { sniff: sniff, kindOf: kindOf, media: media, image: image, byBytes: byBytes };
})(window);
