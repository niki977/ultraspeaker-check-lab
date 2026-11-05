/*
 * The Ultraspeaker Check Lab — color.js
 * Colour and legibility arithmetic shared by the Stile tab, the report and the template:
 *  - contrast ratio (WCAG 2) between text and background, also as projected in a lit room
 *  - the nearest readable colour, the best text colours for a background
 *  - deuteranopia simulation (Viénot 1999), for previews and charts
 *  - minimum text size for the farthest row (AVIXA DISCAS, basic decision making: element height ≥ distance / 200)
 * © Nicholas D. Charles — The Ultraspeaker
 */
(function (g) {
  'use strict';
  var CL = g.CheckLab = g.CheckLab || {};

  /* Light in the room: share of the projector's white that the room light adds to every point of the image (estimate) */
  var LIGHT = { buio: 0.01, penombra: 0.05, accese: 0.12 };
  /* Farthest row, in screen heights */
  var ROWS = { aula: 4, sala: 6, plenaria: 8 };

  function hex(v) {
    var m = /^#?([0-9a-f]{6}|[0-9a-f]{3})$/i.exec(String(v || '').trim());
    if (!m) return null;
    var h = m[1].length === 3 ? m[1].replace(/./g, '$&$&') : m[1];
    return '#' + h.toUpperCase();
  }
  function rgb(h) { var n = parseInt(h.slice(1), 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; }
  function toHex(c) { return '#' + c.map(function (v) { return ('0' + Math.round(Math.max(0, Math.min(255, v))).toString(16)).slice(-2); }).join('').toUpperCase(); }
  function lin(c) { c = c / 255; return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); }
  function enc(v) { v = Math.max(0, Math.min(1, v)); return 255 * (v <= 0.0031308 ? 12.92 * v : 1.055 * Math.pow(v, 1 / 2.4) - 0.055); }
  function lumRgb(c) { return 0.2126 * lin(c[0]) + 0.7152 * lin(c[1]) + 0.0722 * lin(c[2]); }
  function lum(h) { return lumRgb(rgb(h)); }
  /* Contrast of two relative luminances; amb = room light added to the projected image */
  function ratioL(a, b, amb) {
    amb = amb || 0;
    var x = amb + (1 - amb) * a, y = amb + (1 - amb) * b;
    return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05);
  }
  function ratio(fg, bg, amb) { return ratioL(lum(fg), lum(bg), amb); }

  function hsl(h) {
    var c = rgb(h).map(function (v) { return v / 255; }), mx = Math.max.apply(null, c), mn = Math.min.apply(null, c), l = (mx + mn) / 2, hue = 0, s = 0, d = mx - mn;
    if (d) {
      s = d / (1 - Math.abs(2 * l - 1));
      hue = mx === c[0] ? ((c[1] - c[2]) / d) % 6 : (mx === c[1] ? (c[2] - c[0]) / d + 2 : (c[0] - c[1]) / d + 4);
      hue *= 60; if (hue < 0) hue += 360;
    }
    return [hue, s, l];
  }
  function fromHsl(hue, s, l) {
    hue = ((hue % 360) + 360) % 360;
    var c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs((hue / 60) % 2 - 1)), m = l - c / 2, k = Math.floor(hue / 60) % 6;
    var p = [[c, x, 0], [x, c, 0], [0, c, x], [0, x, c], [x, 0, c], [c, 0, x]][k];
    return toHex(p.map(function (v) { return (v + m) * 255; }));
  }

  /* Nearest colour with the same hue that reaches `need`, moving away from the background; null if none */
  function fix(fg, bg, need, amb) {
    var h = hsl(fg), darker = lum(bg) > 0.18;
    for (var i = 1; i <= 100; i++) {
      var c = fromHsl(h[0], h[1], darker ? h[2] * (1 - i / 100) : h[2] + (1 - h[2]) * i / 100);
      if (ratio(c, bg, amb) >= need * 1.05) return c;
    }
    return null;
  }

  /* The best text colours for a background: the stronger neutral, then its own hue, the complementary and a nearby hue,
     each as colourful as it can be while readable; first those that read in the room, then those that need the lights down */
  function bestText(bg, amb, count) {
    var hb = hsl(bg), grey = hb[1] < 0.12;
    var light = ratio('#FFFFFF', bg, amb) >= ratio('#000000', bg, amb);
    function tone(hue, lt, a) {
      for (var i = 0; i <= 60; i++) {
        var c = fromHsl(hue, 0.7, lt ? 0.55 + i * 0.0075 : 0.45 - i * 0.0075);
        if (ratio(c, bg, a) >= 4.5 * 1.1) return c;
      }
      return null;
    }
    var hues = grey ? [[40, 'warm'], [205, 'cold'], [150, 'teal']] : [[hb[0], 'same'], [hb[0] + 180, 'complement'], [hb[0] + 35, 'near'], [hb[0] - 35, 'near']];
    function build(lt, a) {
      return [[lt ? '#FFFFFF' : '#000000', lt ? 'white' : 'black']]
        .concat(hues.map(function (x) { return [tone(x[0], lt, a), x[1]]; }), [[lt ? '#F5F1E6' : '#1F1F1F', lt ? 'ivory' : 'graphite']]);
    }
    function far(a, b) { var p = rgb(a), q = rgb(b); return Math.pow(p[0] - q[0], 2) + Math.pow(p[1] - q[1], 2) + Math.pow(p[2] - q[2], 2) > 1600; }
    var out = [];
    build(light, amb).concat(build(light, 0), build(!light, 0)).forEach(function (x) {
      if (!x[0] || out.length >= (count || 4)) return;
      var r = ratio(x[0], bg, 0);
      if (r >= 4.5 && out.every(function (o) { return far(o.color, x[0]); })) out.push({ color: x[0], kind: x[1], ratio: r, room: ratio(x[0], bg, amb) });
    });
    return out.sort(function (a, b) { return (b.room >= 4.5) - (a.room >= 4.5); });
  }

  /* Deuteranopia (Viénot 1999), in linear RGB */
  function deutan(h) {
    var c = rgb(h).map(lin), r = 0.29275 * c[0] + 0.70725 * c[1];
    return toHex([enc(r), enc(r), enc(-0.02234 * c[0] + 0.02234 * c[1] + c[2])]);
  }
  /* CIE76 colour difference (Lab, D65) */
  function lab(h) {
    var c = rgb(h).map(lin);
    var x = (0.4124 * c[0] + 0.3576 * c[1] + 0.1805 * c[2]) / 0.95047, y = 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2], z = (0.0193 * c[0] + 0.1192 * c[1] + 0.9505 * c[2]) / 1.08883;
    function f(t) { return t > 0.008856 ? Math.cbrt(t) : 7.787 * t + 16 / 116; }
    return [116 * f(y) - 16, 500 * (f(x) - f(y)), 200 * (f(y) - f(z))];
  }
  function deltaE(a, b) { var p = lab(a), q = lab(b); return Math.sqrt(Math.pow(p[0] - q[0], 2) + Math.pow(p[1] - q[1], 2) + Math.pow(p[2] - q[2], 2)); }

  /* Minimum size (in the slide's own units) for text read from the farthest row:
     lowercase height ≥ distance / 200. rows = farthest row in image heights; xh = x-height / em;
     slideH = slide height in the units of the font size (points in PowerPoint, pixels in Keynote);
     the slide fills the screen height only if it is not wider than the screen. */
  function minSize(rows, xh, slideH, slideRatio, roomRatio) {
    var fill = roomRatio && slideRatio ? Math.min(1, roomRatio / slideRatio) : 1;
    return rows * slideH / (200 * (xh || 0.5) * fill);
  }
  /* Rows from the talk settings: a room type, or the measured screen height and farthest row (metres) */
  function rows(settings) {
    var h = parseFloat(String(settings.scrH).replace(',', '.')), d = parseFloat(String(settings.rowD).replace(',', '.'));
    if (settings.rows === 'misure') return h > 0 && d > 0 ? d / h : ROWS.sala;
    return ROWS[settings.rows] || ROWS.sala;
  }

  CL.color = { LIGHT: LIGHT, ROWS: ROWS, hex: hex, rgb: rgb, toHex: toHex, lum: lum, lumRgb: lumRgb, ratio: ratio, ratioL: ratioL, hsl: hsl, fromHsl: fromHsl,
    fix: fix, bestText: bestText, deutan: deutan, deltaE: deltaE, minSize: minSize, rows: rows };
})(window);
