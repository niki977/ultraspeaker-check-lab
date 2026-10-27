/*
 * The Ultraspeaker Check Lab — ui.js
 * Interface shared by the PowerPoint task pane and the online version.
 * The host page passes a "platform" object (app-addin.js / app-web.js) that gets
 * the file, moves to a slide, opens links and hands files to the user.
 * © Nicholas D. Charles — The Ultraspeaker
 */
(function (g) {
  'use strict';
  var CL = g.CheckLab, I = CL.i18n, t = I.t;
  var PROJECTOR_LAB = 'https://niki977.github.io/ultraspeaker-projector-lab/';
  var BASE = g.document.currentScript.src.replace(/assets\/js\/ui\.js.*$/, '');
  var STEPS = ['read', 'fonts', 'media', 'images', 'evaluate'];

  function h(s) { return String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function $(sel) { return g.document.querySelector(sel); }
  function on(sel, ev, fn) { Array.prototype.forEach.call(g.document.querySelectorAll(sel), function (el) { el.addEventListener(ev, function (e) { fn(el, e); }); }); }
  function load(k, def) { try { return JSON.parse(g.localStorage.getItem('ultraspeaker-checklab-' + k)) || def; } catch (e) { return def; } }
  function save(k, v) { try { g.localStorage.setItem('ultraspeaker-checklab-' + k, JSON.stringify(v)); } catch (e) { /* private mode: not remembered */ } }
  function mb(b) { return (b / 1048576).toFixed(0); }

  var ICON = {
    logo: '<svg width="40" height="40" viewBox="0 0 40 40" aria-hidden="true" fill="none" stroke="#FFFFFF" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="34" height="22" rx="2"/><path d="M12 16l5 5 11-11"/><path d="M20 27v6"/><path d="M12 35h16"/></svg>',
    lock: '<svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true" fill="none" stroke="#3DDC84" stroke-width="1.6" stroke-linecap="round"><rect x="3.5" y="8" width="11" height="8" rx="1.5"/><path d="M6 8V5.5a3 3 0 016 0V8"/></svg>',
    info: '<svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true" fill="none" stroke="#BDBDBD" stroke-width="1.6" stroke-linecap="round"><circle cx="9" cy="9" r="7"/><path d="M9 8v4.5"/><path d="M9 5.5v.1"/></svg>',
    done: '<svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true" fill="none" stroke="#3DDC84" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="10" cy="10" r="8"/><path d="M6.5 10.2l2.4 2.4 4.6-4.8"/></svg>',
    run: '<svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true" fill="none" stroke-width="2" stroke-linecap="round"><circle cx="10" cy="10" r="8" stroke="#333333"/><path class="cl-spin" d="M10 2a8 8 0 018 8" stroke="#FFFFFF"/></svg>',
    todo: '<svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true" fill="none" stroke="#4D4D4D" stroke-width="2"><circle cx="10" cy="10" r="8"/></svg>',
    video: '<svg width="22" height="22" viewBox="0 0 22 22" aria-hidden="true" fill="none" stroke="#FFFFFF" stroke-width="1.7" stroke-linejoin="round"><rect x="2" y="5" width="13" height="12" rx="2"/><path d="M15 9.5l5-2.5v8l-5-2.5"/></svg>',
    image: '<svg width="22" height="22" viewBox="0 0 22 22" aria-hidden="true" fill="none" stroke="#FFFFFF" stroke-width="1.7" stroke-linejoin="round"><rect x="2.5" y="3.5" width="17" height="15" rx="2"/><circle cx="8" cy="9" r="1.8"/><path d="M3 16l5-5 4 4 3-3 4 4"/></svg>',
    projector: '<svg width="36" height="36" viewBox="0 0 36 36" aria-hidden="true" fill="none" stroke="#FFFFFF" stroke-width="1.8" stroke-linejoin="round"><rect x="4" y="11" width="20" height="14" rx="2"/><circle cx="14" cy="18" r="3.5"/><path d="M24 15l8-4v14l-8-4"/></svg>'
  };

  /* ---------- state ---------- */
  var P;   // platform
  var S = { tab: 'check', running: false, token: null, error: '', notice: '', read: 0, total: 0, steps: {}, zip: null, res: null, ev: null, kitBusy: false, kitStatus: {} };
  var settings = Object.assign({}, CL.DEFAULT_SETTINGS, load('settings', {}));
  var delivery = load('delivery', { date: '', time: '', room: '', surname: '', talk: '' });
  function defaultList() { return [0, 1, 2, 3, 4, 5, 6].map(function (i) { return { def: i, done: false }; }); }
  var checklist = load('checklist', null) || defaultList();
  function saveAll() { save('settings', settings); save('delivery', delivery); save('checklist', checklist); }
  function evaluate() { S.ev = CL.evaluate(S.res, Object.assign({}, settings, { hostOS: P.hostOS })); }

  /* ---------- pieces ---------- */
  function note(icon, key) { return '<div class="cl-note">' + ICON[icon] + '<span>' + h(t(key)) + '</span></div>'; }
  function alerts() {
    return (S.error ? '<div class="cl-alert" role="alert">' + h(S.error) + '</div>' : '') +
      (S.notice ? '<div class="cl-alert info" role="status">' + h(S.notice) + '</div>' : '');
  }
  function select(key, values, prefix) {
    return '<select class="cl-select" data-set="' + key + '">' + values.map(function (v) {
      return '<option value="' + v + '"' + (v === settings[key] ? ' selected' : '') + '>' + h(t(prefix + v)) + '</option>';
    }).join('') + '</select>';
  }
  function number(key) { return '<input class="cl-input" type="number" min="1" data-set="' + key + '" value="' + h(settings[key]) + '">'; }
  function field(label, control) { return '<label class="cl-label">' + h(t(label)) + control + '</label>'; }
  function slideLabel(slides) { return !slides.length ? '' : (slides.length > 4 ? t('rep.slides', { count: slides.length }) : t('rep.slide', { list: slides.join(', ') })); }
  function itemDetail(x) {
    return [x.px && x.need ? x.px + ' → ' + x.need + ' px' : (x.px ? x.px + ' px' : ''), x.codec, x.format, x.mb && x.mb + ' MB', x.save && '−' + x.save + ' MB',
      x.pct && x.pct + '%', x.device, x.date && x.date.replace(/^(\d{4}):(\d{2}):(\d{2})/, '$1-$2-$3'), x.gps && 'GPS', x.slides && slideLabel(x.slides)].filter(Boolean).join(' · ');
  }

  /* ---------- tabs ---------- */
  function checkTab() {
    if (S.running) {
      var pct = Math.round(60 * (S.total ? S.read / S.total : 0) + 10 * STEPS.filter(function (k) { return S.steps[k] === 'done' && k !== 'read'; }).length);
      return '<section class="cl-panel" aria-live="polite"><h2 class="cl-h2" style="font-size:26px">' + h(t('prog.title')) + '</h2><span class="cl-muted">' + h(t('prog.sub')) + '</span>' +
        '<div style="display:flex;justify-content:space-between;align-items:baseline"><span class="cl-progress-num">' + pct + '%</span><span class="cl-muted">' +
        (S.total ? h(t('prog.bytes', { a: mb(S.read), b: mb(S.total) })) : '') + '</span></div>' +
        '<div class="cl-bar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="' + pct + '"><div style="width:' + pct + '%"></div></div>' +
        '<ol class="cl-steps">' + STEPS.map(function (k) {
          var s = S.steps[k] || 'todo';
          return '<li class="' + s + '">' + ICON[s] + '<span>' + h(t('step.' + k)) + '</span><span class="st">' + (s === 'todo' ? '' : h(t(s === 'done' ? 'prog.done' : 'prog.running'))) + '</span></li>';
        }).join('') + '</ol><button class="cl-btn" id="cl-cancel"' + (S.token.cancelled ? ' disabled' : '') + '>' + h(t('btn.cancel')) + '</button></section>';
    }
    var source = P.mode === 'web'
      ? '<h1 class="cl-h1">' + h(t('app.tagline')) + '</h1><p class="cl-muted" style="font-size:15px;margin:0">' + h(t('app.lead')) + '</p>' +
        '<label class="cl-drop" id="cl-drop"><span style="font-size:16px;font-weight:600">' + h(t('file.drop')) + '</span><span class="cl-btn primary">' + h(t('file.choose')) + '</span>' +
        '<input id="cl-file" type="file" accept=".pptx,.key" class="sr-only"><span class="cl-muted">' + h(t('file.formats')) + '</span></label>'
      : '<section class="cl-card"><span class="cl-overline">' + h(t('file.open')) + '</span><div style="font-size:15px;font-weight:600;overflow-wrap:anywhere">' +
        h(P.documentName() || t('file.current')) + '</div><span class="cl-muted">' + h(P.platformLabel) + '</span></section>';
    return source + alerts() +
      '<section class="cl-panel" style="gap:14px"><h2 class="cl-h2">' + h(t('talk.title')) + '</h2>' +
      field('set.venueOS', select('venueOS', ['win', 'mac', 'own', 'unknown'], 'os.')) + field('set.minutes', number('minutes')) +
      field('set.room', select('room', ['16:9', '16:10', '4:3', '3:1', 'unknown'], 'room.')) +
      '<details class="cl-adv"><summary>' + h(t('set.advanced')) + '</summary><div class="cl-grid2">' +
      field('set.warnMB', number('warnMB')) + field('set.critMB', number('critMB')) + field('set.minSec', number('minSec')) + field('set.paceSec', number('paceSec')) + '</div></details></section>' +
      '<section class="cl-panel" style="gap:8px"><span class="cl-overline">' + h(t('checks.title')) + '</span><div class="cl-chips">' +
      ['fonts', 'media', 'images', 'format', 'time', 'privacy'].map(function (c) { return '<span>' + h(t('cat.' + c)) + '</span>'; }).join('') + '</div></section>' +
      (P.mode === 'web' ? '' : '<button class="cl-btn primary big" id="cl-start">▶ ' + h(t('btn.start')) + '</button>') + note('lock', 'privacy.local');
  }

  function issueCard(iss) {
    var risk = I.issue(iss.id, 'r', iss.params), fix = I.issue(iss.id, 'f', iss.params);
    var goto = function (n) { return P.canGoto ? ' <button class="cl-link" data-goto="' + n + '">' + h(t('btn.goto', { n: n })) + ' →</button>' : ''; };
    var items = iss.items.filter(itemDetail);
    return '<article class="cl-issue ' + iss.sev + '"><div class="head"><span class="title">' + h(I.issue(iss.id, 't', iss.params)) + '</span>' +
      '<span class="cl-muted">' + h(t('cat.' + iss.cat)) + (iss.slides.length ? ' · ' + h(slideLabel(iss.slides)) : '') + '</span></div>' +
      (risk ? '<p>' + h(risk) + '</p>' : '') + (fix ? '<p class="fix"><strong>' + h(t('rep.fix')) + '</strong> ' + h(fix) + '</p>' : '') +
      (items.length ? '<details><summary>' + h(t('btn.showList', { count: items.length })) + '</summary><ul class="cl-items">' + items.map(function (x) {
        return '<li><span>' + h(x.name) + '</span><span>' + h(itemDetail(x)) + (x.slides && x.slides.length ? goto(x.slides[0]) : '') + '</span></li>';
      }).join('') + '</ul></details>' : '') +
      (iss.slides.length && P.canGoto ? '<div>' + iss.slides.slice(0, 3).map(goto).join('') + '</div>' : '') + '</article>';
  }

  function reportTab() {
    if (!S.ev) return alerts() + '<p class="cl-empty">' + h(t('rep.empty')) + '</p>';
    var ev = S.ev, sm = ev.summary, col = ev.crit ? '#FF5A5A' : (ev.warn ? '#FFB547' : '#3DDC84'), c = 2 * Math.PI * 48;
    function list(sev) { return ev.issues.filter(function (i) { return i.sev === sev; }); }
    var ok = list('ok');
    return alerts() + '<section class="cl-score"><svg width="112" height="112" viewBox="0 0 112 112" role="img" aria-label="' + ev.score + ' / 100" style="flex:none">' +
      '<circle cx="56" cy="56" r="48" fill="none" stroke="#262626" stroke-width="10"/><circle cx="56" cy="56" r="48" fill="none" stroke="' + col + '" stroke-width="10" stroke-linecap="round" stroke-dasharray="' +
      (c * ev.score / 100).toFixed(1) + ' ' + c.toFixed(1) + '" transform="rotate(-90 56 56)"/>' +
      '<text x="56" y="62" text-anchor="middle" fill="#FFFFFF" font-family="Barlow Condensed, sans-serif" font-weight="700" font-size="38">' + ev.score + '</text>' +
      '<text x="56" y="80" text-anchor="middle" fill="#A3A3A3" font-family="Inter, sans-serif" font-size="11">' + h(t('rep.of100')) + '</text></svg>' +
      '<div class="cl-panel" style="gap:6px"><span class="cl-overline">' + h(t('rep.ready')) + '</span><span class="cl-verdict">' + h(t('rep.verdict.' + ev.verdict)) + '</span>' +
      '<span class="cl-muted">' + h(t('rep.sub.' + ev.verdict)) + '</span><span class="cl-muted" style="overflow-wrap:anywhere">' + h(sm.fileName) + '</span></div></section>' +
      '<div class="cl-counts">' + ['crit', 'warn', 'ok'].map(function (k) { return '<div class="cl-count ' + k + '"><b>' + ev[k] + '</b>' + h(t('rep.' + k)) + '</div>'; }).join('') + '</div>' +
      ['crit', 'warn'].map(function (sev) { var l = list(sev); return l.length ? '<h3 class="cl-h3 ' + sev + '">' + h(t('rep.sec.' + sev)) + '</h3>' + l.map(issueCard).join('') : ''; }).join('') +
      '<h3 class="cl-h3 ok">' + h(t('rep.sec.ok', { count: ok.length })) + '</h3><ul class="cl-oklist">' + ok.map(function (i) {
        return '<li><span>' + h(I.issue(i.id, 't', i.params)) + i.items.map(function (x) { return '<br><span class="cl-muted">' + h(x.name + ' · ' + itemDetail(x)) + '</span>'; }).join('') + '</span><span class="tick" aria-hidden="true">✓</span></li>';
      }).join('') + '</ul>' +
      '<section class="cl-card cl-muted" style="line-height:1.6"><strong style="color:#fff">' + h(t('rep.summary')) + '</strong><br>' +
      h(t('rep.sum.format', { aspect: sm.aspect, mb: sm.fileMB < 10 ? sm.fileMB.toFixed(1) : sm.fileMB.toFixed(0) })) + '<br>' +
      h(t('rep.sum.slides', { total: sm.slides, visible: sm.visible })) + '<br>' + h(t('rep.sum.media', { videos: sm.videos, audios: sm.audios, mb: sm.videoMB.toFixed(0) })) +
      (sm.fonts.length ? '<br><strong style="color:#fff">' + h(t('rep.sum.fonts')) + '</strong><ul class="cl-fonts">' + sm.fonts.map(function (f) {
        var cls = f.embedded || f.cls === 'office' ? 'ok' : (f.cls === 'custom' ? 'crit' : 'warn');
        return '<li><span>' + h(f.name) + '</span><span class="cl-badge ' + cls + '">' + h(t(f.embedded ? 'rep.font.embedded' : 'rep.font.' + f.cls)) + '</span></li>';
      }).join('') + '</ul>' : '') + '</section>' +
      note('info', 'note.fontsRef') + (sm.type === 'key' ? note('info', 'note.keyLimits') : '') +
      (P.canGoto ? '' : note('info', P.mode === 'web' ? 'note.gotoWeb' : 'note.gotoUnavailable')) + (P.downloadNote ? note('info', P.downloadNote) : '') +
      '<div class="cl-actions"><button class="cl-btn primary" id="cl-pdf">' + h(t('btn.pdf')) + '</button><button class="cl-btn" id="cl-copy">' + h(t('btn.copy')) + '</button></div>' +
      '<button class="cl-btn" id="cl-again">' + h(t(P.mode === 'web' ? 'btn.newFile' : 'btn.recheck')) + '</button>';
  }

  function deliveryName() {
    function slug(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-+|-+$/g, ''); }
    return [delivery.date, slug(delivery.room).replace(/-/g, ''), (delivery.time || '').replace(':', ''), slug(delivery.surname), slug(delivery.talk)]
      .filter(Boolean).join('_') + (S.res && S.res.type === 'key' ? '.key' : '.pptx');
  }
  function venueTab() {
    function input(k, type) { return field('del.' + k, '<input class="cl-input" type="' + type + '" data-del="' + k + '" value="' + h(delivery[k]) + '">'); }
    var videos = S.res ? S.res.media.filter(function (m) { return !m.external && !m.missing; }) : [];
    var images = S.res ? CL.kit.myImages(S.zip, S.res, S.ev, t).length : 0;
    function kit(kind) {
      var mine = kind === 'video' ? videos.length : images;
      var info = !S.res ? 'del.tests.needCheck' : (!mine ? (kind === 'video' ? 'del.tests.noVideo' : 'del.tests.noImage') : (kind === 'video' ? 'del.tests.videoInfo' : 'del.tests.imageInfo'));
      return '<div class="cl-card cl-panel" style="gap:10px"><strong>' + ICON[kind] + ' ' + h(t('del.tests.' + kind)) + '</strong>' +
        '<span class="cl-muted">' + h(t(info, { count: mine, mb: mb(videos.reduce(function (a, m) { return a + m.bytes; }, 0)) })) + '</span><div class="cl-grid2">' +
        '<button class="cl-btn primary sm" data-kit="' + kind + ':mine"' + (!mine || S.kitBusy ? ' disabled' : '') + '>' + h(t('del.tests.mine.' + kind)) + '</button>' +
        '<button class="cl-btn sm" data-kit="' + kind + ':generic"' + (S.kitBusy ? ' disabled' : '') + '>' + h(t('del.tests.generic')) + '</button></div>' +
        '<span class="cl-muted" role="status">' + h(S.kitStatus[kind] || '') + '</span></div>';
    }
    var done = checklist.filter(function (c) { return c.done; }).length;
    return '<section class="cl-panel" style="gap:12px"><h2 class="cl-h2">' + h(t('del.name.title')) + '</h2>' +
      '<div class="cl-grid2">' + input('date', 'date') + input('time', 'time') + input('room', 'text') + input('surname', 'text') + '</div>' + input('talk', 'text') +
      '<div class="cl-card cl-panel" style="gap:10px"><code class="cl-filename" id="cl-fname">' + h(deliveryName()) + '</code>' +
      '<button class="cl-btn sm" id="cl-copyname" style="align-self:flex-start">' + h(t('del.copy')) + '</button></div><span class="cl-muted">' + h(t('del.hint')) + '</span></section>' +
      '<section class="cl-panel" style="gap:10px"><h2 class="cl-h2">' + h(t('del.check.title')) + ' <span class="cl-muted">' + h(t('del.check.count', { done: done, total: checklist.length })) + '</span></h2>' +
      '<ul class="cl-check">' + checklist.map(function (c, i) {
        var text = c.def !== undefined && c.def !== null ? t('del.items')[c.def] : c.text;
        return '<li class="' + (c.done ? 'done' : '') + '"><label><input type="checkbox" data-check="' + i + '"' + (c.done ? ' checked' : '') + '><span>' + h(text) + '</span></label>' +
          '<button class="cl-x" data-remove="' + i + '" aria-label="' + h(t('del.remove') + ': ' + text) + '">×</button></li>';
      }).join('') + '</ul>' +
      '<form class="cl-addrow" id="cl-addform"><input id="cl-newitem" class="cl-input" aria-label="' + h(t('del.addPrompt')) + '" placeholder="' + h(t('del.addPrompt')) + '"><button class="cl-btn sm">' + h(t('del.add')) + '</button></form>' +
      '<button class="cl-btn sm" id="cl-resetlist" style="align-self:flex-start">' + h(t('del.reset')) + '</button></section>' +
      '<button class="cl-linkcard" id="cl-projector">' + ICON.projector + '<span style="flex:1"><strong class="t">PROJECTOR LAB</strong><br><span class="cl-muted">' + h(t('del.projector')) + '</span></span>›</button>' +
      '<section class="cl-panel" style="gap:12px"><h2 class="cl-h2">' + h(t('del.tests.title')) + '</h2><span class="cl-muted">' + h(t('del.tests.intro')) + '</span>' +
      kit('video') + kit('image') + '<span class="cl-muted">' + h(t('del.tests.key')) + '</span></section>';
  }

  function render() {
    var main = S.tab === 'report' ? reportTab() : (S.tab === 'venue' ? venueTab() : checkTab());
    $('#app').innerHTML =
      '<header class="cl-header"><span class="cl-brand">' + ICON.logo + '<span class="cl-brand-text"><span class="cl-brand-over">THE ULTRASPEAKER</span><span class="cl-brand-name">CHECK LAB</span></span></span>' +
      '<select id="cl-lang" class="cl-select small" aria-label="' + h(t('lang.label')) + '">' + I.LANGS.map(function (l) {
        return '<option value="' + l + '"' + (l === I.lang ? ' selected' : '') + '>' + l.toUpperCase() + '</option>';
      }).join('') + '</select></header>' +
      '<nav class="cl-tabs" role="tablist" aria-label="' + h(t('nav.sections')) + '">' + ['check', 'report', 'venue'].map(function (k) {
        return '<button class="cl-tab" role="tab" data-tab="' + k + '" aria-selected="' + (S.tab === k) + '">' + h(t('nav.' + k)) + '</button>';
      }).join('') + '</nav><main class="cl-main">' + main + '</main>' +
      '<footer class="cl-footer">' + ['guide', 'privacy', 'terms', 'support'].map(function (k) { return '<button class="cl-link" data-page="' + k + '">' + h(t('foot.' + k)) + '</button>'; }).join('') + '</footer>';
    bind();
  }
  function go(tab) { S.tab = tab; S.error = ''; render(); g.scrollTo(0, 0); }

  /* ---------- events ---------- */
  function bind() {
    on('#cl-lang', 'change', function (el) { I.set(el.value); render(); });
    on('[data-tab]', 'click', function (el) { go(el.dataset.tab); });
    on('[data-page]', 'click', function (el) { P.openUrl(BASE + el.dataset.page + '.html?lang=' + I.lang); });
    on('[data-set]', 'change', function (el) {
      settings[el.dataset.set] = el.type === 'number' && el.value !== '' ? Number(el.value) : el.value;
      saveAll();
      if (S.res) evaluate();
    });
    on('#cl-start', 'click', function () { run(); });
    on('#cl-cancel', 'click', function (el) { S.token.cancelled = true; el.disabled = true; });
    on('#cl-file', 'change', function (el) { if (el.files[0]) run(el.files[0]); });
    on('#cl-drop', 'dragover', function (el, e) { e.preventDefault(); el.classList.add('over'); });
    on('#cl-drop', 'dragleave', function (el) { el.classList.remove('over'); });
    on('#cl-drop', 'drop', function (el, e) {
      e.preventDefault();
      var item = e.dataTransfer.items && e.dataTransfer.items[0], entry = item && item.webkitGetAsEntry && item.webkitGetAsEntry();
      if (entry && entry.isDirectory) { S.error = t(/\.key$/i.test(entry.name) ? 'err.package' : 'err.type'); render(); }
      else if (e.dataTransfer.files[0]) run(e.dataTransfer.files[0]);
    });
    on('[data-goto]', 'click', function (el) { P.goto(Number(el.dataset.goto)); });
    on('#cl-pdf', 'click', exportPdf);
    on('#cl-copy', 'click', function (el) { copy(reportText(), el); });
    on('#cl-again', 'click', function () { if (P.mode === 'web') { S.res = S.ev = S.zip = null; go('check'); } else run(); });
    on('[data-del]', 'input', function (el) { delivery[el.dataset.del] = el.value; saveAll(); $('#cl-fname').textContent = deliveryName(); });
    on('#cl-copyname', 'click', function (el) { copy(deliveryName(), el); });
    on('[data-check]', 'change', function (el) { checklist[el.dataset.check].done = el.checked; saveAll(); render(); });
    on('[data-remove]', 'click', function (el) { checklist.splice(el.dataset.remove, 1); saveAll(); render(); });
    on('#cl-addform', 'submit', function (el, e) {
      e.preventDefault();
      var v = $('#cl-newitem').value.trim();
      if (v) { checklist.push({ text: v, done: false }); saveAll(); render(); $('#cl-newitem').focus(); }
    });
    on('#cl-resetlist', 'click', function () { checklist = defaultList(); saveAll(); render(); });
    on('#cl-projector', 'click', function () { P.openUrl(PROJECTOR_LAB); });
    on('[data-kit]', 'click', function (el) { var p = el.dataset.kit.split(':'); makeKit(p[0], p[1] === 'mine'); });
  }

  /* ---------- the check ---------- */
  function step(k, state) { S.steps[k] = state; render(); }
  async function run(webFile) {
    if (S.running) return;
    var token = S.token = { cancelled: false };
    Object.assign(S, { running: true, tab: 'check', error: '', notice: '', read: 0, total: 0, steps: { read: 'run' } });
    render();
    try {
      var file = await P.getFile(webFile, function (read, total) { S.read = read; S.total = total; render(); }, token);
      step('read', 'done');
      var zip;
      try { zip = await g.JSZip.loadAsync(file.bytes); } catch (e) { throw new Error(/\.key$/i.test(file.name) ? 'notKey' : 'notPptx'); }
      file.bytes = null;
      var opts = { fileName: file.name, fileBytes: file.size, token: token, onProgress: function (k) {
        if (S.steps[k] === 'run') return;
        STEPS.slice(1, STEPS.indexOf(k)).forEach(function (p) { S.steps[p] = 'done'; });
        step(k, 'run');
      } };
      var res = /\.key$/i.test(file.name) ? await CL.analyzeKey(zip, opts) : await CL.analyzePptx(zip, opts);
      S.zip = zip; S.res = res; S.kitStatus = {};
      evaluate();
      S.running = false;
      go('report');
    } catch (e) {
      var code = e.name === 'CancelError' ? 'cancelled' : (e instanceof RangeError ? 'memory' : e.message);
      var known = ['cancelled', 'notPptx', 'notKey', 'type', 'getFile', 'memory'];
      if (known.indexOf(code) < 0) g.console.error(e);
      S.running = false;
      S.error = t('err.' + (known.indexOf(code) >= 0 ? code : 'read'));
      go('check');
    }
  }

  /* ---------- report as text and PDF ---------- */
  function settingsLine() { return t('pdf.settings', { os: t('os.' + settings.venueOS), room: settings.room === 'unknown' ? '?' : settings.room, minutes: settings.minutes || '-' }); }
  function reportLines() {
    var ev = S.ev, out = [];
    ['crit', 'warn', 'ok'].forEach(function (sev) {
      var l = ev.issues.filter(function (i) { return i.sev === sev; });
      if (!l.length) return;
      out.push({ sev: sev, head: sev === 'ok' ? t('rep.sec.ok', { count: l.length }) : t('rep.sec.' + sev) });
      l.forEach(function (i) {
        out.push({ sev: sev, title: I.issue(i.id, 't', i.params) + (i.slides.length ? ' (' + slideLabel(i.slides) + ')' : '') });
        if (sev !== 'ok') ['r', 'f'].forEach(function (part) { var s = I.issue(i.id, part, i.params); if (s) out.push({ text: (part === 'f' ? t('rep.fix') + ' ' : '') + s }); });
        i.items.slice(0, 25).forEach(function (x) { out.push({ item: '- ' + x.name + (itemDetail(x) ? ': ' + itemDetail(x) : '') }); });
      });
    });
    return out;
  }
  function reportText() {
    var ev = S.ev;
    return ['THE ULTRASPEAKER · CHECK LAB', t('pdf.title') + ' — ' + ev.summary.fileName, t('rep.date', { date: new Date().toLocaleString(I.locale) }),
      t('rep.ready') + ': ' + ev.score + '/100 · ' + t('rep.verdict.' + ev.verdict), settingsLine(), ''].concat(reportLines().map(function (l) {
      return l.head ? '\n== ' + l.head.toUpperCase() + ' ==' : (l.title ? '- ' + l.title : '  ' + (l.text || l.item));
    }), ['', t('note.fontsRef')]).join('\n');
  }
  function exportPdf() {
    var ev = S.ev, doc = new g.jspdf.jsPDF({ unit: 'pt', format: 'a4' });
    var W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight(), M = 48, y = 124;
    var COLOR = { crit: [208, 40, 40], warn: [214, 140, 20], ok: [30, 150, 80] };
    function write(s, size, bold, color, indent) {
      doc.setFont('helvetica', bold ? 'bold' : 'normal'); doc.setFontSize(size); doc.setTextColor.apply(doc, color || [20, 20, 20]);
      // jsPDF's built-in font covers Latin-1 only
      var safe = String(s).replace(/[‘’]/g, "'").replace(/[“”«»]/g, '"').replace(/›/g, '>').replace(/[–—−]/g, '-').replace(/→/g, '->').replace(/…/g, '...').replace(/[^\x00-\xFF]/g, '?');
      doc.splitTextToSize(safe, W - 2 * M - (indent || 0)).forEach(function (line) {
        if (y > H - M) { doc.addPage(); y = M; }
        doc.text(line, M + (indent || 0), y); y += size * 1.35;
      });
    }
    doc.setFillColor(10, 10, 10); doc.rect(0, 0, W, 92, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(163, 163, 163); doc.text('THE ULTRASPEAKER', M, 38);
    doc.setFontSize(22); doc.setTextColor(255, 255, 255); doc.text('CHECK LAB', M, 62);
    write(ev.summary.fileName, 13, true);
    write(t('rep.date', { date: new Date().toLocaleString(I.locale) }) + ' · ' + settingsLine(), 9, false, [110, 110, 110]);
    y += 10;
    write(ev.score + '/100 · ' + t('rep.verdict.' + ev.verdict).toUpperCase() + '   (' + t('rep.crit') + ' ' + ev.crit + ' · ' + t('rep.warn') + ' ' + ev.warn + ' · ' + t('rep.ok') + ' ' + ev.ok + ')', 16, true, COLOR[ev.crit ? 'crit' : (ev.warn ? 'warn' : 'ok')]);
    reportLines().forEach(function (l) {
      if (l.head) { y += 10; write(l.head.toUpperCase(), 11, true, COLOR[l.sev]); }
      else if (l.title) { y += 4; write(l.title, 10.5, l.sev !== 'ok', null, 10); }
      else if (l.text) write(l.text, 9.5, false, [60, 60, 60], 10);
      else write(l.item, 8.5, false, [90, 90, 90], 20);
    });
    y += 10; write(t('note.fontsRef'), 8.5, false, [110, 110, 110]);
    for (var p = 1, n = doc.getNumberOfPages(); p <= n; p++) {
      doc.setPage(p); doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(140, 140, 140);
      doc.text('The Ultraspeaker Check Lab · niki977.github.io/ultraspeaker-check-lab', M, H - 24); doc.text(p + ' / ' + n, W - M, H - 24, { align: 'right' });
    }
    P.saveFile(new Uint8Array(doc.output('arraybuffer')), 'CheckLab_' + ev.summary.fileName.replace(/\.(pptx|key)$/i, '').replace(/[^A-Za-z0-9]+/g, '-') + '.pdf', 'application/pdf');
  }
  function copy(text, btn) {
    var label = btn.textContent;
    function done() { btn.textContent = t('del.copied'); setTimeout(function () { btn.textContent = label; }, 1500); }
    function fallback() { // Office task panes may block the Clipboard API
      var ta = g.document.createElement('textarea'); ta.value = text; g.document.body.appendChild(ta); ta.select();
      try { g.document.execCommand('copy'); done(); } catch (e) { /* nothing else to try */ }
      ta.remove();
    }
    if (g.navigator.clipboard) g.navigator.clipboard.writeText(text).then(done, fallback); else fallback();
  }

  /* ---------- test kits ---------- */
  function asset(name) { return g.fetch(BASE + 'assets/kit/' + name).then(function (r) { return r.arrayBuffer(); }).then(function (b) { return new Uint8Array(b); }); }
  async function makeKit(kind, mine) {
    S.kitBusy = true; S.kitStatus[kind] = t('del.tests.building'); render();
    try {
      var items = !mine ? CL.kit.genericItems(kind, asset, t) : (kind === 'video' ? CL.kit.myVideos(S.zip, S.res, t) : CL.kit.myImages(S.zip, S.res, S.ev, t));
      var bytes = await CL.kit.build({ kind: kind, items: items, t: t, lang: I.locale, template: await asset('template.pptx'), poster: await asset('poster-video.png'),
        title: t('kit.title.' + (mine ? 'my' : 'gen') + (kind === 'video' ? 'Video' : 'Image')), intro: t(kind === 'video' ? 'kit.intro.video' : 'kit.intro.image') });
      var name = 'CheckLab_' + (mine ? 'MyTest' : 'RoomTest') + '_' + (kind === 'video' ? 'Video' : 'Images') + '.pptx';
      var how = await P.openKit(bytes, name);
      S.kitStatus[kind] = how === 'opened' ? t('del.tests.opened') : (how === 'downloaded' ? t('del.tests.ready', { name: name }) : t('note.downloadMac'));
    } catch (e) {
      g.console.error(e);
      S.kitStatus[kind] = t('err.read');
    }
    S.kitBusy = false; render();
  }

  CL.UI = {
    init: function (platform) { P = platform; I.set(I.detect(P.langHint)); g.document.body.classList.add(P.mode); render(); },
    notice: function (msg) { S.notice = msg; render(); },
    download: function (bytes, name, type) {
      var a = g.document.createElement('a');
      a.href = URL.createObjectURL(new Blob([bytes], { type: type })); a.download = name;
      g.document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 5000);
    }
  };
})(window);
