/*
 * The Ultraspeaker Check Lab — analyze-pptx.js
 * Reads a PowerPoint .pptx (Open XML) with JSZip, in the browser, and collects
 * the facts the report needs. The rules that turn facts into problems are in evaluate.js.
 * © Nicholas D. Charles — The Ultraspeaker
 */
(function (g) {
  'use strict';
  var CL = g.CheckLab = g.CheckLab || {};

  function CancelError() { this.name = 'CancelError'; }
  CL.CancelError = CancelError;

  function unxml(s) {
    return String(s).replace(/&#x([0-9a-f]+);/gi, function (m, h) { return String.fromCharCode(parseInt(h, 16)); })
      .replace(/&#(\d+);/g, function (m, d) { return String.fromCharCode(+d); })
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  }
  function attr(tag, name) { var m = new RegExp('\\s' + name + '="([^"]*)"').exec(tag); return m ? unxml(m[1]) : null; }
  function all(re, s) { var out = [], m; re.lastIndex = 0; while ((m = re.exec(s))) out.push(m); return out; }
  function baseName(p) { return String(p).split(/[\\/]/).pop(); }
  function resolve(from, target) {
    var out = [];
    (from.replace(/[^/]*$/, '') + target).split('/').forEach(function (p) { if (p === '..') out.pop(); else if (p && p !== '.') out.push(p); });
    return out.join('/');
  }
  async function read(zip, path) { var f = zip.file(path); return f ? f.async('string') : ''; }
  /* Relationships of a part, as a map id -> { type, target, external, path } */
  async function rels(zip, path) {
    var xml = await read(zip, path.replace(/([^/]*)$/, '_rels/$1.rels')), map = {};
    all(/<Relationship\s[^>]*>/g, xml).forEach(function (m) {
      var r = { type: (attr(m[0], 'Type') || '').split('/').pop(), target: attr(m[0], 'Target') || '', external: attr(m[0], 'TargetMode') === 'External' };
      r.path = r.external ? null : resolve(path, r.target);
      map[attr(m[0], 'Id')] = r;
    });
    return map;
  }
  function list(map) { return Object.keys(map).map(function (k) { var r = map[k]; r.id = k; return r; }); }
  function text(xml) { return all(/<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g, xml).map(function (m) { return unxml(m[1]); }).join(''); }

  async function analyze(zip, opts) {
    var token = opts.token || {};
    function check() { if (token.cancelled) throw new CancelError(); }
    function progress(step, frac) { if (opts.onProgress) opts.onProgress(step, frac); }
    var res = { type: 'pptx', fileName: opts.fileName, fileBytes: opts.fileBytes, slideW: 0, slideH: 0, slides: [], fonts: [], embeddedFonts: [], media: [], images: [], links: [] };

    var pres = await read(zip, 'ppt/presentation.xml');
    if (!pres) throw new Error('notPptx');
    var sz = /<p:sldSz\s[^>]*>/.exec(pres);
    if (sz) { res.slideW = +attr(sz[0], 'cx'); res.slideH = +attr(sz[0], 'cy'); }
    res.embeddedFonts = all(/<p:font\s[^>]*typeface="([^"]*)"/g, (/<p:embeddedFontLst>[\s\S]*?<\/p:embeddedFontLst>/.exec(pres) || [''])[0]).map(function (m) { return unxml(m[1]); });
    var presRels = await rels(zip, 'ppt/presentation.xml');

    // Fonts: theme fonts, fonts named in masters/layouts, fonts named on each slide
    var fonts = {};
    function addFont(name, slide) {
      if (!name || /^\+/.test(name)) return;
      var f = fonts[name.toLowerCase()] = fonts[name.toLowerCase()] || { name: name, slides: [] };
      if (slide && f.slides.indexOf(slide) < 0) f.slides.push(slide);
    }
    function fontNames(xml) { return all(/<a:(?:latin|ea|cs|sym|buFont)\s[^>]*typeface="([^"]+)"/g, xml).map(function (m) { return unxml(m[1]); }); }
    var themeRel = list(presRels).filter(function (r) { return r.type === 'theme'; })[0];
    var theme = await read(zip, themeRel ? themeRel.path : 'ppt/theme/theme1.xml');
    var themeFont = {};
    ['major', 'minor'].forEach(function (k) {
      var m = new RegExp('<a:' + k + 'Font>\\s*<a:latin\\s[^>]*typeface="([^"]+)"').exec(theme);
      if (m) themeFont[k === 'major' ? '+mj-lt' : '+mn-lt'] = unxml(m[1]);
    });
    var layouts = Object.keys(zip.files).filter(function (p) { return /^ppt\/(slideMasters|slideLayouts)\/[^/]+\.xml$/.test(p); });
    for (var l = 0; l < layouts.length; l++) fontNames(await read(zip, layouts[l])).forEach(function (n) { addFont(n); });
    addFont(themeFont['+mj-lt']); addFont(themeFont['+mn-lt']);

    var ids = all(/<p:sldId\s[^>]*r:id="([^"]+)"/g, pres).map(function (m) { return m[1]; });
    var imageUses = {}, mediaByKey = {}, shownName = {};

    for (var i = 0; i < ids.length; i++) {
      check();
      progress('fonts', i / ids.length);
      var num = i + 1, path = presRels[ids[i]] && presRels[ids[i]].path;
      if (!path) continue;
      var xml = await read(zip, path), srel = await rels(zip, path), srels = list(srel);
      var tree = (/<p:spTree>([\s\S]*)<\/p:spTree>/.exec(xml) || ['', ''])[1];
      var ownShapes = all(/<p:sp[\s>][\s\S]*?<\/p:sp>/g, tree).filter(function (m) { return m[0].indexOf('<p:ph') < 0; }).length;
      var slide = {
        index: num, hidden: attr((/<p:sld[\s>][^>]*>/.exec(xml) || [''])[0], 'show') === '0', notes: false,
        empty: !text(xml).trim() && !ownShapes && !/<p:(pic|graphicFrame|cxnSp|contentPart)[\s>]/.test(tree)
      };
      fontNames(xml).forEach(function (n) { addFont(themeFont[n] || n, num); });
      ['+mj-lt', '+mn-lt'].forEach(function (k) { if (themeFont[k] && xml.indexOf('typeface="' + k.slice(0, 4)) >= 0 && text(xml).trim()) addFont(themeFont[k], num); });

      var notes = srels.filter(function (r) { return r.type === 'notesSlide'; })[0];
      if (notes) {
        var nx = await read(zip, notes.path);
        slide.notes = all(/<p:sp[\s>][\s\S]*?<\/p:sp>/g, nx).some(function (m) { return /<p:ph\s[^>]*type="body"/.test(m[0]) && text(m[0]).trim(); });
      }

      // Video and audio (embedded or linked); one entry per file, with every slide using it
      srels.forEach(function (r) {
        if (!/^(video|audio|media)$/.test(r.type)) return;
        var key = r.external ? r.target : r.path, m = mediaByKey[key];
        if (!m) {
          m = mediaByKey[key] = { key: key, path: r.path, target: r.target, slides: [], external: r.external, online: r.external && /^https?:/i.test(r.target), name: r.external ? r.target : baseName(r.path) };
          if (m.external && !m.online) m.name = baseName(decodeURIComponent(r.target));
          res.media.push(m);
        }
        if (r.type !== 'media') m.relKind = r.type;
        if (m.slides.indexOf(num) < 0) m.slides.push(num);
      });

      // Pictures with their size on the slide; media posters are marked
      all(/<p:pic[\s>][\s\S]*?<\/p:pic>/g, tree).forEach(function (pm) {
        var pb = pm[0], isMedia = /<a:(videoFile|audioFile|quickTimeFile)\s/.test(pb);
        if (isMedia) {
          var link = /<a:(?:videoFile|audioFile|quickTimeFile)\s[^>]*r:link="([^"]+)"/.exec(pb), mr = link && srel[link[1]];
          if (mr) shownName[mr.external ? mr.target : mr.path] = attr((/<p:cNvPr\s[^>]*>/.exec(pb) || [''])[0], 'name');
        }
        var blip = /<a:blip\s[^>]*r:embed="([^"]+)"/.exec(pb), r1 = blip && srel[blip[1]];
        if (!r1) return;
        r1.used = true;
        if (r1.external) { res.links.push({ slide: num, kind: 'image', target: r1.target }); return; }
        var ext = /<a:ext\s[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec((/<a:xfrm[\s\S]*?<\/a:xfrm>/.exec(pb) || [''])[0]);
        var src = (/<a:srcRect\s[^>]*>/.exec(pb) || [''])[0], crop = {};
        ['l', 't', 'r', 'b'].forEach(function (s) { crop[s] = Math.max(0, +(attr(src, s) || 0)) / 100000; });
        var use = { slide: num, cx: ext ? +ext[1] : 0, crop: crop, poster: isMedia };
        var svg = /<asvg:svgBlip\s[^>]*r:embed="([^"]+)"/.exec(pb), r2 = svg && srel[svg[1]];
        if (r2 && r2.path) { r2.used = true; (imageUses[r2.path] = imageUses[r2.path] || []).push({ slide: num, cx: use.cx, crop: crop, svg: true, fallback: r1.path }); }
        (imageUses[r1.path] = imageUses[r1.path] || []).push(use);
      });
      srels.forEach(function (r) {
        if (r.type === 'image' && !r.used) {
          if (r.external) res.links.push({ slide: num, kind: 'image', target: r.target });
          else (imageUses[r.path] = imageUses[r.path] || []).push({ slide: num, cx: 0, crop: { l: 0, t: 0, r: 0, b: 0 } }); // backgrounds and fills
        }
        if (r.type === 'hyperlink' && r.external && !/^mailto:/i.test(r.target)) res.links.push({ slide: num, kind: /^https?:/i.test(r.target) ? 'web' : 'file', target: r.target });
        if (r.type === 'oleObject' && r.external) res.links.push({ slide: num, kind: 'ole', target: r.target });
      });
      var charts = srels.filter(function (r) { return r.type === 'chart'; });
      for (var c = 0; c < charts.length; c++) {
        list(await rels(zip, charts[c].path)).forEach(function (r) { if (r.external && !/^https?:/i.test(r.target)) res.links.push({ slide: num, kind: 'chart', target: r.target }); });
      }
      res.slides.push(slide);
    }
    Object.keys(fonts).forEach(function (k) { res.fonts.push(fonts[k]); });

    // Probe the embedded media
    for (var j = 0; j < res.media.length; j++) {
      check();
      progress('media', j / res.media.length);
      var m = res.media[j];
      if (shownName[m.key] && !m.external) m.name = /\.[a-z0-9]{2,4}$/i.test(shownName[m.key]) ? shownName[m.key] : shownName[m.key] + '.' + m.key.split('.').pop();
      if (m.external) { m.kind = m.relKind || (/\.(mp3|wav|wma|m4a|aiff?)$/i.test(m.target) ? 'audio' : 'video'); continue; }
      if (!zip.file(m.path)) { m.missing = true; m.kind = m.relKind || 'video'; continue; }
      var bytes = await zip.file(m.path).async('uint8array');
      m.bytes = bytes.length;
      m.info = CL.probe.media(bytes, m.path);
      m.kind = m.info.kind === 'audio' ? 'audio' : 'video';
    }

    // Probe the images
    var paths = Object.keys(imageUses);
    for (var q = 0; q < paths.length; q++) {
      check();
      progress('images', q / paths.length);
      var uses = imageUses[paths[q]], file = zip.file(paths[q]);
      if (!file) continue;
      res.images.push({ path: paths[q], name: baseName(paths[q]), info: CL.probe.image(await file.async('uint8array'), paths[q]), uses: uses,
        poster: uses.every(function (u) { return u.poster; }) });
    }
    return res;
  }

  CL.analyzePptx = analyze;
})(typeof window !== 'undefined' ? window : globalThis);
