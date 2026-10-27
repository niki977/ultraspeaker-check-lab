/*
 * The Ultraspeaker Check Lab — i18n.js
 * IT, EN, ES, FR, DE. The chosen language is remembered on this device only.
 * © Nicholas D. Charles — The Ultraspeaker
 */
(function (g) {
  'use strict';
  var CL = g.CheckLab = g.CheckLab || {}, LANGS = ['it', 'en', 'es', 'fr', 'de'], current = 'en';

  function fill(s, p) { return String(s).replace(/\{(\w+)\}/g, function (m, k) { return p && p[k] !== undefined ? p[k] : m; }); }
  function dict() { return CL.I18N[current]; }

  CL.i18n = {
    LANGS: LANGS,
    get lang() { return current; },
    get locale() { return dict()._lang; },
    /* ?lang= in the address, then the saved choice, then the browser or Office language */
    detect: function (hint) {
      var saved = null;
      try { saved = g.localStorage.getItem('ultraspeaker-checklab-lang'); } catch (e) { /* not available */ }
      var l = (/[?&]lang=([a-z]{2})/.exec(g.location.search) || [])[1] || saved || String(hint || g.navigator.language).slice(0, 2).toLowerCase();
      return LANGS.indexOf(l) >= 0 ? l : 'en';
    },
    set: function (l) {
      current = l;
      g.document.documentElement.lang = l;
      try { g.localStorage.setItem('ultraspeaker-checklab-lang', l); } catch (e) { /* not remembered */ }
    },
    t: function (key, params) { var v = dict()[key]; return Array.isArray(v) ? v.slice() : fill(v === undefined ? key : v, params); },
    /* part: 't' title, 'r' risk, 'f' fix */
    issue: function (id, part, params) {
      var p = Object.assign({}, params);
      if (p.bars) p.bars = dict()['bars.' + p.bars];
      return fill(dict().issues[id][part] || '', p);
    }
  };
})(window);
