/*
 * The Ultraspeaker Check Lab — app-addin.js
 * PowerPoint add-in: reads the open presentation with Office.context.document.getFileAsync
 * (compressed format, 4 MB slices). What a platform cannot do is said in the pane.
 * © Nicholas D. Charles — The Ultraspeaker
 */
(function (g) {
  'use strict';
  var CL = g.CheckLab, t = CL.i18n.t;

  function call(fn) {
    return new Promise(function (ok, ko) { fn(function (r) { if (r.status === Office.AsyncResultStatus.Succeeded) ok(r.value); else ko(new Error('getFile')); }); });
  }
  function base64(bytes) {
    var s = '';
    for (var i = 0; i < bytes.length; i += 32768) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 32768));
    return g.btoa(s);
  }

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.PowerPoint) {
      CL.UI.init({ mode: 'addin', documentName: function () { return ''; }, platformLabel: '' });
      CL.UI.notice(t('err.office'));
      return;
    }
    var ctx = Office.context, doc = ctx.document, plat = ctx.platform, isMac = plat === Office.PlatformType.Mac;
    var canCreate = g.PowerPoint && PowerPoint.createPresentation && ctx.requirements.isSetSupported('PowerPointApi', '1.1');
    CL.UI.init({
      mode: 'addin', langHint: ctx.displayLanguage,
      hostOS: isMac ? 'mac' : (plat === Office.PlatformType.PC ? 'win' : (/Mac/.test(g.navigator.userAgent) ? 'mac' : 'win')),
      platformLabel: 'PowerPoint · ' + (isMac ? 'Mac' : (plat === Office.PlatformType.PC ? 'Windows' : 'web')),
      downloadNote: isMac ? 'note.downloadMac' : null,
      canGoto: typeof doc.goToByIdAsync === 'function',
      documentName: function () { return decodeURIComponent(String(doc.url || '').split(/[\\/]/).pop()); },
      getFile: async function (_, onProgress, token) {
        var file = await call(function (cb) { doc.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, cb); });
        try {
          var out = new Uint8Array(file.size), off = 0;
          for (var i = 0; i < file.sliceCount; i++) {
            if (token.cancelled) throw new CL.CancelError();
            var data = (await call(function (cb) { file.getSliceAsync(i, cb); })).data;
            out.set(data, off); off += data.length;
            onProgress(off, file.size);
          }
          return { bytes: out, name: (this.documentName() || 'presentation').replace(/(\.pptx)?$/i, '.pptx'), size: file.size };
        } finally { file.closeAsync(); }
      },
      goto: function (n) {
        doc.goToByIdAsync(n, Office.GoToType.Index, function (r) { if (r.status !== Office.AsyncResultStatus.Succeeded) CL.UI.notice(t('note.gotoUnavailable')); });
      },
      openUrl: function (url) {
        if (ctx.ui.openBrowserWindow) ctx.ui.openBrowserWindow(url); else g.open(url, '_blank');
      },
      saveFile: function (bytes, name, type) {
        if (isMac) CL.UI.notice(t('note.downloadMac')); else CL.UI.download(bytes, name, type);
      },
      openKit: async function (bytes, name) {
        if (canCreate) {
          try { await PowerPoint.createPresentation(base64(bytes)); return 'opened'; } catch (e) { /* fall back to a download */ }
        }
        if (isMac) return 'failed';
        CL.UI.download(bytes, name, 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
        return 'downloaded';
      }
    });
  });
})(window);
