/*
 * The Ultraspeaker Check Lab — fonts-db.js
 * Fonts normally present on a venue computer. The pane cannot see the fonts
 * really installed there: these lists are a reference, and the report says so.
 * © Nicholas D. Charles — The Ultraspeaker
 */
(function (g) {
  'use strict';
  var CL = g.CheckLab = g.CheckLab || {};

  // Installed with Microsoft Office (Windows and Mac) or present on both systems
  var BOTH = 'Aptos,Aptos Display,Aptos Narrow,Aptos Light,Aptos SemiBold,Aptos ExtraBold,Aptos Black,Aptos Serif,Aptos Mono,Arial,Arial Black,Arial Narrow,' +
    'Arial Rounded MT Bold,Book Antiqua,Bookman Old Style,Calibri,Calibri Light,Cambria,Cambria Math,Candara,Century,Century Gothic,Century Schoolbook,' +
    'Comic Sans MS,Consolas,Constantia,Corbel,Courier New,Franklin Gothic Book,Franklin Gothic Medium,Garamond,Georgia,Gill Sans MT,Impact,Lucida Console,' +
    'Lucida Sans Unicode,Palatino Linotype,Rockwell,Symbol,Tahoma,Times New Roman,Trebuchet MS,Tw Cen MT,Verdana,Webdings,Wingdings,Wingdings 2,Wingdings 3,Andale Mono';
  // Windows only
  var WIN = 'Bahnschrift,Cascadia Code,Cascadia Mono,Ebrima,Gabriola,Gadugi,Ink Free,Javanese Text,Leelawadee UI,Malgun Gothic,Marlett,Microsoft Himalaya,' +
    'Microsoft JhengHei,Microsoft New Tai Lue,Microsoft PhagsPa,Microsoft Sans Serif,Microsoft Tai Le,Microsoft YaHei,Microsoft Yi Baiti,MingLiU-ExtB,Mongolian Baiti,' +
    'MS Gothic,MS PGothic,MV Boli,Myanmar Text,Nirmala UI,Segoe MDL2 Assets,Segoe Fluent Icons,Segoe Print,Segoe Script,Segoe UI,Segoe UI Emoji,Segoe UI Historic,' +
    'Segoe UI Symbol,Segoe UI Variable,SimSun,NSimSun,Sitka,Sylfaen,Yu Gothic,Yu Gothic UI';
  // macOS only
  var MAC = 'AppleSystemUIFont,San Francisco,SF Pro,SF Pro Text,SF Pro Display,SF Pro Rounded,SF Compact,SF Mono,New York,Helvetica,Helvetica Neue,Avenir,Avenir Next,' +
    'Avenir Next Condensed,Futura,Gill Sans,Optima,Palatino,Times,Courier,Menlo,Monaco,Baskerville,Didot,Hoefler Text,American Typewriter,Chalkboard,Chalkboard SE,' +
    'Chalkduster,Copperplate,Big Caslon,Cochin,Marker Felt,Noteworthy,Papyrus,Phosphate,Skia,Superclarendon,Bradley Hand,Brush Script MT,Charter,Seravek,Athelas,' +
    'Iowan Old Style,Lucida Grande,Geneva,Apple Chancery,Apple SD Gothic Neo,Hiragino Sans,PingFang SC,Snell Roundhand,Zapfino,Trattatello,Herculanum,Luminari,Kefa,' +
    'DIN Alternate,DIN Condensed,Arial Hebrew,Geeza Pro,Kohinoor Devanagari';

  var WEIGHTS = /(Regular|Roman|Book|Light|ExtraLight|UltraLight|Thin|Medium|SemiBold|DemiBold|Demi|Bold|ExtraBold|Heavy|Black|Italic|Oblique|Condensed|Narrow|Display|Text|Rounded)+$/i;
  function norm(name) { return String(name).toLowerCase().replace(/[\s_.\-]/g, ''); }
  function set(list) { var s = {}; list.split(',').forEach(function (f) { s[norm(f)] = 1; }); return s; }
  var SETS = { office: set(BOTH), win: set(WIN), mac: set(MAC) };

  /* 'office' (Windows and Mac), 'win', 'mac' or 'custom' */
  function classify(name) {
    if (!name || /^\+(mj|mn)-/.test(name)) return 'office';
    var full = norm(name), family = norm(String(name).split('-')[0].replace(WEIGHTS, ''));
    for (var k in SETS) if (SETS[k][full] || SETS[k][family]) return k;
    return 'custom';
  }

  CL.fonts = { classify: classify, norm: norm };
})(window);
