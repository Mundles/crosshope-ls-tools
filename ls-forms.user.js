// ==UserScript==
// @name         CrossHope LS Forms
// @namespace    https://crosshopetreatment.com/ls-tools
// @version      2.3.0
// @description  CrossHope helpers for LightningStep: Open a form, New patient (Teams VOB + prescreen PDF to an admitted chart), Ready for Assessment, Discharge, Change level of care, Treatment plans and New letter (CrossHope letterhead in Word).
// @author       CrossHope
// @match        https://chtc.lightningstep.com/*
// @match        https://restaurabh.sharepoint.com/sites/RestauraBH/*
// @noframes
// @grant        none
// @run-at       document-idle
// @updateURL    https://raw.githubusercontent.com/Mundles/crosshope-ls-tools/main/ls-forms.user.js
// @downloadURL  https://raw.githubusercontent.com/Mundles/crosshope-ls-tools/main/ls-forms.user.js
// ==/UserScript==

// CrossHope LS Forms. No patient data is stored in this file or sent anywhere
// except LightningStep itself. It reads the census and the chart from
// LightningStep with the user's own sign-in, then opens the form's page.

(function () {
  'use strict';
  // New letter, part 2: runs on SharePoint only when the LightningStep button
  // opened it (?chletter=1). Copies the letterhead template into CrossHope
  // Letters with a new name, then opens the copy in Word for the web.
  // The template itself is never opened or changed.
  if (location.hostname !== 'restaurabh.sharepoint.com') return;
  if (!/[?&]chletter=1/.test(location.search)) return;
  var SITE = '/sites/RestauraBH';
  var FOLDER = SITE + '/Shared Documents/Arkansas Clinic/CrossHope Letters';
  var TEMPLATE = FOLDER + '/_Template (do not edit)/CrossHope Letterhead.docx';
  var H = { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' };
  function say(html, bad) {
    var d = document.getElementById('chletter-msg');
    if (!d) { d = document.createElement('div'); d.id = 'chletter-msg'; document.documentElement.appendChild(d); }
    d.style.cssText = 'position:fixed;top:20px;left:50%;transform:translateX(-50%);z-index:2147483647;background:#fff;border:3px solid ' + (bad ? '#b91c1c' : '#1a73e8') + ';border-radius:10px;padding:16px 20px;font:16px system-ui,"Segoe UI",sans-serif;max-width:560px;box-shadow:0 6px 24px rgba(0,0,0,.25)';
    d.innerHTML = html;
  }
  function stamp() {
    var n = new Date(), h = n.getHours(), ap = h < 12 ? 'AM' : 'PM'; h = h % 12 || 12;
    var p = function (x) { return (x < 10 ? '0' : '') + x; };
    return n.getFullYear() + '-' + p(n.getMonth() + 1) + '-' + p(n.getDate()) + ' ' + h + '-' + p(n.getMinutes()) + '-' + p(n.getSeconds()) + ' ' + ap;
  }
  async function go() {
    say('<b>Making a new CrossHope letter...</b>');
    try {
      var ci = await fetch(SITE + '/_api/contextinfo', { method: 'POST', headers: H, credentials: 'include' });
      if (!ci.ok) throw new Error('no-access');
      var dg = (await ci.json()).FormDigestValue;
      var name = 'Letter ' + stamp() + '.docx';
      var dst = FOLDER + '/' + name;
      var r = await fetch(SITE + '/_api/SP.MoveCopyUtil.CopyFileByPath(overwrite=false)', { method: 'POST', credentials: 'include', headers: Object.assign({ 'X-RequestDigest': dg }, H),
        body: JSON.stringify({ srcPath: { DecodedUrl: location.origin + TEMPLATE }, destPath: { DecodedUrl: location.origin + dst } }) });
      if (r.status === 401 || r.status === 403) throw new Error('no-access');
      if (!r.ok) throw new Error('copy ' + r.status);
      var f = await fetch(SITE + "/_api/web/GetFileByServerRelativePath(decodedurl='" + encodeURIComponent(dst).replace(/'/g, "''") + "')?$select=UniqueId", { headers: H, credentials: 'include' });
      if (!f.ok) throw new Error('find ' + f.status);
      var id = (await f.json()).UniqueId;
      say('<b>Opening your letter in Word...</b><br>It saves itself in CrossHope Letters as "' + name + '". Rename it when you\'re done (use initials, not the full name).');
      location.replace(SITE + '/_layouts/15/Doc.aspx?sourcedoc={' + id + '}&action=edit');
    } catch (e) {
      if (e.message === 'no-access') say('<b>You don\'t have access to the CrossHope Letters folder yet.</b><br>Ask Mendy to share it with you, then click New letter again.', true);
      else say('<b>The letter didn\'t open.</b> (' + e.message + ')<br>Try New letter again. If it keeps happening, tell Mendy.', true);
    }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', go); else go();
})();

(function () {
  'use strict';
  if (location.hostname !== 'chtc.lightningstep.com') return;

  // Only on signed-in pages.
  if (/^\/(login|password|session|logout)/i.test(location.pathname)) return;
  if (!document.querySelector('meta[name="csrf-token"]')) return;
  if (document.getElementById('lsfp-btn')) return;

  var FORMS = [
    ['MH Biopsych', 'Mental Health Biopsychosocial'],
    ['SUD Biopsych', 'Biopsychosocial Assessment AR'],
    ['LOCUS', 'LOCUS Reassessment'],
    ['ASAM', 'ASAM Reassessment'],
    ['CIWA-Ar', 'Custom CIWA AR Assessment'],
    ['COWS', 'Custom COWS Assessment'],
    ['C-SSRS + Safety Plan', 'CSSRS Screener with Safety Plan'],
    ['Safety Plan', 'Stanley Brown Safety Plan'],
    ['Individual Note', '#note'],
    ['Treatment Plan Review', '#tpr']
  ];
  var sel = 'LOCUS Reassessment';
  try { sel = localStorage.getItem('lsfp-form') || sel; } catch (e) {}
  if (!FORMS.some(function (x) { return x[1] === sel; })) sel = 'LOCUS Reassessment';

  // Never print the button or the box.
  var st = document.createElement('style');
  st.textContent = '@media print{#lsfp-bar,#lsfp,#lsnp,#lslc,#lstp{display:none!important}}';
  document.head.appendChild(st);

  var btn = document.createElement('button');
  btn.id = 'lsfp-btn';
  btn.type = 'button';
  btn.textContent = 'Open a form';
  btn.title = 'Pick a form, type the patient\'s name, and the form opens on their chart.';
  btn.style.cssText = 'position:fixed;right:18px;bottom:18px;z-index:2147483646;background:#1a73e8;color:#fff;border:0;border-radius:24px;padding:11px 18px;font:600 15px system-ui,"Segoe UI",sans-serif;cursor:pointer;box-shadow:0 4px 14px rgba(0,0,0,.25)';
  btn.onclick = function () { togglePicker(); };
  btn.style.position = 'static'; btn.style.boxShadow = 'none';
  var bar = document.createElement('div');
  bar.id = 'lsfp-bar';
  bar.style.cssText = 'position:fixed;right:18px;bottom:18px;z-index:2147483646;display:flex;gap:8px;border-radius:24px';
  var mk = function (id, label, color, fn) { var b = document.createElement('button'); b.id = id; b.type = 'button'; b.textContent = label; b.style.cssText = 'background:' + color + ';color:#fff;border:0;border-radius:24px;padding:11px 18px;font:600 15px system-ui,"Segoe UI",sans-serif;cursor:pointer;box-shadow:0 4px 14px rgba(0,0,0,.25)'; b.onclick = fn; return b; };
  if (/\/(client\/index|inquiry\/show)\/\d+/.test(location.pathname)) {
    bar.appendChild(mk('lsnp-ready', 'Ready for Assessment', '#7c3aed', function () { NP.readyCheck(); }));
    bar.appendChild(mk('lslc-loc', 'Level of care', '#b45309', function () { LC.openLoc(); }));
    bar.appendChild(mk('lslc-dc', 'Discharge', '#b91c1c', function () { LC.openDischarge(); }));
  }
  var tpBtn = mk('lstp-btn', 'Treatment plans', '#0e7490', function () { TP.open(); });
  bar.appendChild(tpBtn);
  setTimeout(function () { try { TP.badge(tpBtn); } catch (e) {} }, 1500);
  bar.appendChild(mk('lsltr-btn', 'New letter', '#1e3a8a', function () {
    // Opens a fresh copy of the CrossHope letterhead in Word for the web.
    window.open('https://restaurabh.sharepoint.com/sites/RestauraBH/Shared%20Documents/Forms/AllItems.aspx?chletter=1&id=' + encodeURIComponent('/sites/RestauraBH/Shared Documents/Arkansas Clinic/CrossHope Letters'), '_blank');
  }));
  bar.appendChild(mk('lsnp-btn', 'New patient', '#0f766e', function () { NP.openWizard(); }));
  btn.style.boxShadow = '0 4px 14px rgba(0,0,0,.25)';
  bar.appendChild(btn);
  document.body.appendChild(bar);

  function togglePicker() {
    var old = document.getElementById('lsfp');
    if (old) { old.remove(); return; }
    openPicker();
  }

  function openPicker() {
    var box = document.createElement('div');
    box.id = 'lsfp';
    box.style.cssText = 'position:fixed;top:60px;left:50%;transform:translateX(-50%);z-index:2147483647;background:#fff;border:3px solid #1a73e8;border-radius:12px;padding:16px;width:min(560px,94vw);max-height:80vh;overflow:auto;font:16px system-ui,"Segoe UI",sans-serif;color:#111;box-shadow:0 8px 30px rgba(0,0,0,.3)';
    box.innerHTML = '<div style="display:flex;justify-content:space-between;align-items:center"><b style="font-size:18px">Open a form for a patient</b><button id="lsfpx" type="button" style="font-size:22px;border:0;background:none;cursor:pointer" title="Close">&times;</button></div>' +
      '<div id="lsfpf" style="display:flex;flex-wrap:wrap;gap:6px;margin:10px 0"></div>' +
      '<input id="lsfpq" placeholder="Type the patient\'s name or MRN" autocomplete="off" style="width:100%;box-sizing:border-box;font-size:17px;padding:10px;border:2px solid #999;border-radius:8px">' +
      '<div id="lsfpm" style="margin:8px 0;color:#555;font-size:14px">Loading the census...</div><div id="lsfpr"></div>';
    document.body.appendChild(box);
    var $ = function (id) { return document.getElementById(id); };
    $('lsfpx').onclick = function () { box.remove(); };

    function drawForms() {
      var f = $('lsfpf'); f.innerHTML = '';
      FORMS.forEach(function (x) {
        var b = document.createElement('button');
        b.type = 'button';
        b.textContent = x[0];
        var on = x[1] === sel;
        b.style.cssText = 'padding:7px 10px;border-radius:8px;cursor:pointer;font-size:14px;border:2px solid #1a73e8;background:' + (on ? '#1a73e8;color:#fff' : '#fff;color:#1a73e8');
        b.onclick = function () {
          sel = x[1];
          try { localStorage.setItem('lsfp-form', sel); } catch (e) {}
          drawForms(); $('lsfpq').focus();
        };
        f.appendChild(b);
      });
    }
    drawForms();

    var strip = function (s) { return String(s == null ? '' : s).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(); };
    var rows = [];
    var t = document.querySelector('meta[name="csrf-token"]').content;
    fetch('/home/census', {
      method: 'POST', credentials: 'same-origin',
      headers: { 'X-CSRF-TOKEN': t, 'X-Requested-With': 'XMLHttpRequest', 'Content-Type': 'application/x-www-form-urlencoded', 'Accept': 'application/json' },
      body: '_token=' + encodeURIComponent(t) + '&draw=1&start=0&length=2000&programfilter%5B%5D=0'
    }).then(function (r) { return r.json(); }).then(function (j) {
      rows = (j.data || []).map(function (r) {
        var a = String(r.name || '').match(/<a\b[^>]*>([^<]*)<\/a>/i);
        return { c: r.client_id, n: strip(a ? a[1] : r.lnamefname), m: strip(r.mrn), p: strip(r.pgmname), l: strip(r.location) };
      });
      $('lsfpm').textContent = rows.length + ' patients on the census. Type a name, then click the patient (or press Enter).';
      draw();
    }).catch(function () { $('lsfpm').textContent = 'Could not load the census. Make sure you are signed in to LightningStep.'; });

    function matches() {
      var q = $('lsfpq').value.toLowerCase().trim();
      if (q.length < 2) return [];
      var w = q.split(/\s+/);
      return rows.filter(function (r) { var s = (r.n + ' ' + r.m).toLowerCase(); return w.every(function (x) { return s.indexOf(x) > -1; }); }).slice(0, 15);
    }
    function draw() {
      var out = $('lsfpr'); out.innerHTML = '';
      matches().forEach(function (r) {
        var d = document.createElement('div');
        d.style.cssText = 'padding:10px;border:1px solid #ccc;border-radius:8px;margin:5px 0;cursor:pointer';
        d.innerHTML = '<b></b> <span style="color:#555"></span>';
        d.querySelector('b').textContent = r.n;
        d.querySelector('span').textContent = 'MRN ' + r.m + ' · ' + r.p + ' · ' + r.l;
        d.onmouseenter = function () { d.style.background = '#e8f0fe'; };
        d.onmouseleave = function () { d.style.background = ''; };
        d.onclick = function () { go(r); };
        out.appendChild(d);
      });
    }
    $('lsfpq').oninput = draw;
    $('lsfpq').onkeydown = function (e) {
      if (e.key === 'Enter') { var m = matches(); if (m.length) go(m[0]); }
      if (e.key === 'Escape') box.remove();
    };
    setTimeout(function () { $('lsfpq').focus(); }, 50);

    function go(r) {
      var label = FORMS.filter(function (x) { return x[1] === sel; })[0][0];
      $('lsfpm').textContent = 'Opening ' + label + ' for ' + r.n + '...';
      $('lsfpr').innerHTML = '';
      fetch('/client/index/' + r.c, { credentials: 'same-origin' }).then(function (x) { return x.text(); }).then(function (h) {
        var url = null;
        if (sel === '#note') { url = (h.match(/\/listnote\/\d+\/create\/prognote\/edit\?episodeId=\d+/) || [])[0]; }
        else if (sel === '#tpr') {
          url = (h.match(/\/txplan\/create\/\d+\/1\?episodeId=\d+/) || [])[0];
          if (!url) { $('lsfpm').textContent = r.n + ' has no treatment plan yet, so there is nothing to review. Start a Treatment Plan from the chart first.'; return; }
        } else {
          var d = new DOMParser().parseFromString(h, 'text/html');
          var el = [].slice.call(d.querySelectorAll('[data-url]')).filter(function (e) { return strip(e.textContent) === sel; })[0];
          url = el && el.getAttribute('data-url');
        }
        if (!url) { $('lsfpm').textContent = 'Could not find ' + label + ' on ' + r.n + '\'s chart.'; return; }
        location.href = url;
      }).catch(function () { $('lsfpm').textContent = 'Could not open the chart.'; });
    }
  }

// Minimal text extractor for Chrome/Skia PDFs (Jotform prescreen). No outside code.
// pdfText(bytes: Uint8Array, inflate: async (Uint8Array) => Uint8Array) -> array of lines,
// each line = array of {x, s} chunks sorted left to right; caller joins.
async function pdfText(bytes, inflate) {
  const latin = new TextDecoder('latin1').decode(bytes);
  const objs = {};
  const re = /(\d+)\s+0\s+obj\b/g;
  let m;
  const starts = [];
  while ((m = re.exec(latin))) starts.push([+m[1], m.index + m[0].length]);
  for (let i = 0; i < starts.length; i++) {
    const [num, s] = starts[i];
    const e = latin.indexOf('endobj', s);
    const body = latin.slice(s, e);
    const o = { dict: body, stream: null };
    const si = body.indexOf('stream');
    if (si >= 0 && /\/Length/.test(body.slice(0, si))) {
      let ds = s + si + 6;
      if (latin[ds] === '\r') ds++;
      if (latin[ds] === '\n') ds++;
      const lenM = body.slice(0, si).match(/\/Length\s+(\d+)(\s+0\s+R)?/);
      let len;
      if (lenM && !lenM[2]) len = +lenM[1];
      else len = latin.indexOf('endstream', ds) - ds;
      o.dict = body.slice(0, si);
      o.raw = bytes.subarray(ds, ds + len);
    }
    objs[num] = o;
  }
  const getStream = async (n) => {
    const o = objs[n];
    if (!o || !o.raw) return '';
    if (o.text != null) return o.text;
    let b = o.raw;
    if (/\/FlateDecode/.test(o.dict)) b = await inflate(b);
    o.text = new TextDecoder('latin1').decode(b);
    return o.text;
  };
  const hex2 = (h) => { const a = []; for (let i = 0; i + 1 < h.length; i += 2) a.push(parseInt(h.substr(i, 2), 16)); return a; };
  const cmapCache = {};
  const cmap = async (n) => {
    if (cmapCache[n]) return cmapCache[n];
    const t = await getStream(n);
    const map = {}; let bytesPer = 2;
    const cs = t.match(/begincodespacerange\s*<([0-9a-fA-F]+)>/); if (cs) bytesPer = cs[1].length / 2;
    t.replace(/beginbfchar([\s\S]*?)endbfchar/g, (_, blk) => {
      blk.replace(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>/g, (__, a, b) => { map[parseInt(a, 16)] = uni(b); return ''; });
      return '';
    });
    t.replace(/beginbfrange([\s\S]*?)endbfrange/g, (_, blk) => {
      blk.replace(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(<([0-9a-fA-F]+)>|\[([^\]]*)\])/g, (__, a, b, c, d, arr) => {
        const lo = parseInt(a, 16), hi = parseInt(b, 16);
        if (d) { let base = parseInt(d, 16); for (let k = lo; k <= hi; k++) map[k] = String.fromCodePoint(base + k - lo); }
        else { const items = arr.match(/<([0-9a-fA-F]+)>/g) || []; items.forEach((it, j) => { map[lo + j] = uni(it.slice(1, -1)); }); }
        return '';
      });
      return '';
    });
    return (cmapCache[n] = { map, bytesPer });
  };
  function uni(h) { let s = ''; for (let i = 0; i + 3 < h.length + 1 && i < h.length; i += 4) s += String.fromCharCode(parseInt(h.substr(i, 4), 16)); return s; }
  const ref = (dict, key) => { const r = dict.match(new RegExp('/' + key + '\\s+(\\d+)\\s+0\\s+R')); return r ? +r[1] : null; };
  const dictOf = (txt, key) => { // inline dict after /Key << ... >>
    const i = txt.indexOf('/' + key); if (i < 0) return null;
    let j = txt.indexOf('<<', i); if (j < 0) return null;
    // if a ref comes first, follow it
    const between = txt.slice(i + key.length + 1, j);
    const r = between.match(/^\s*(\d+)\s+0\s+R/);
    if (r) return objs[+r[1]] ? objs[+r[1]].dict : null;
    let depth = 0, k = j;
    for (; k < txt.length; k++) { if (txt.startsWith('<<', k)) { depth++; k++; } else if (txt.startsWith('>>', k)) { depth--; k++; if (!depth) break; } }
    return txt.slice(j, k + 1);
  };
  // pages in order
  const pages = Object.keys(objs).map(Number).filter(n => /\/Type\s*\/Page\b(?!s)/.test(objs[n].dict)).sort((a, b) => a - b);
  // order by Kids if available
  const root = Object.values(objs).find(o => /\/Type\s*\/Pages\b/.test(o.dict) && !/\/Parent/.test(o.dict));
  let order = pages;
  if (root) {
    const flat = [];
    const walk = (d) => { const k = d.match(/\/Kids\s*\[([^\]]*)\]/); if (!k) return; (k[1].match(/(\d+)\s+0\s+R/g) || []).forEach(x => { const n = +x.split(/\s+/)[0]; const o = objs[n]; if (!o) return; if (/\/Type\s*\/Pages\b/.test(o.dict)) walk(o.dict); else flat.push(n); }); };
    walk(root.dict);
    if (flat.length) order = flat;
  }
  const lines = [];
  let pageOffset = 0;
  for (const pn of order) {
    const pd = objs[pn].dict;
    const res = dictOf(pd, 'Resources') || '';
    const fontDict = dictOf(res, 'Font') || '';
    const fonts = {};
    (fontDict.match(/\/([A-Za-z0-9_.+-]+)\s+(\d+)\s+0\s+R/g) || []).forEach(x => { const mm = x.match(/\/([A-Za-z0-9_.+-]+)\s+(\d+)/); const fo = objs[+mm[2]]; const tu = fo && ref(fo.dict, 'ToUnicode'); fonts[mm[1]] = tu; });
    const cRefs = [];
    const cm = pd.match(/\/Contents\s*(\[([^\]]*)\]|(\d+)\s+0\s+R)/);
    if (cm) { if (cm[3]) cRefs.push(+cm[3]); else (cm[2].match(/(\d+)\s+0\s+R/g) || []).forEach(x => cRefs.push(+x.split(/\s+/)[0])); }
    let content = '';
    for (const c of cRefs) content += (await getStream(c)) + '\n';
    // tokenize
    const toks = content.match(/<[0-9a-fA-F\s]*>|\((?:\\.|[^\\)])*\)|\[|\]|\/[^\s/\[\]()<>]+|-?\d*\.?\d+|[A-Za-z'"*]+/g) || [];
    let stack = [], arr = null, font = null, fsize = 10, ctm = [1, 0, 0, 1, 0, 0], tm = [1, 0, 0, 1, 0, 0], tlm = [1, 0, 0, 1, 0, 0];
    const gstack = [];
    const mul = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
    const chunks = [];
    const emit = async (hexes) => {
      if (!font || !fonts[font]) return;
      const cmp = await cmap(fonts[font]);
      let s = '';
      for (const h of hexes) { const b = hex2(h.replace(/\s/g, '')); for (let i = 0; i + cmp.bytesPer <= b.length; i += cmp.bytesPer) { let code = 0; for (let k = 0; k < cmp.bytesPer; k++) code = code * 256 + b[i + k]; s += cmp.map[code] != null ? cmp.map[code] : ''; } }
      if (!s) return;
      const p = mul(tm, ctm);
      const sz = Math.abs(fsize * Math.hypot(p[0], p[1])) || 10;
      chunks.push({ x: p[4], y: pageOffset - p[5], s, sz });
    };
    for (const t of toks) {
      if (t === '[') { arr = []; continue; }
      if (t === ']') { stack.push(arr); arr = null; continue; }
      if (arr) { arr.push(t); continue; }
      if (/^[A-Za-z'"*]+$/.test(t)) {
        const op = t, a = stack; stack = [];
        if (op === 'q') gstack.push(ctm.slice());
        else if (op === 'Q') ctm = gstack.pop() || [1, 0, 0, 1, 0, 0];
        else if (op === 'cm') ctm = mul(a.slice(-6).map(Number), ctm);
        else if (op === 'BT') { tm = [1, 0, 0, 1, 0, 0]; tlm = tm.slice(); }
        else if (op === 'Tm') { tm = a.slice(-6).map(Number); tlm = tm.slice(); }
        else if (op === 'Td' || op === 'TD') { const [tx, ty] = a.slice(-2).map(Number); tlm = mul([1, 0, 0, 1, tx, ty], tlm); tm = tlm.slice(); }
        else if (op === 'Tf') { font = String(a[a.length - 2] || '').slice(1); fsize = +a[a.length - 1] || 10; }
        else if (op === 'Tj' || op === "'" || op === '"') { const h = a[a.length - 1]; if (h && h[0] === '<') await emit([h.slice(1, -1)]); }
        else if (op === 'TJ') { const ar = a[a.length - 1] || []; await emit(ar.filter(x => x[0] === '<').map(x => x.slice(1, -1))); }
        continue;
      }
      stack.push(t);
    }
    // group by y (Skia flips y: larger y = lower), page coordinates
    let maxY = -1e9;
    chunks.forEach(c => { if (c.y > maxY) maxY = c.y; });
    chunks.sort((a, b) => a.y - b.y || a.x - b.x);
    for (const c of chunks) {
      const L = lines.length && Math.abs(lines[lines.length - 1].y - c.y) < c.sz * 0.6 ? lines[lines.length - 1] : null;
      if (L) L.parts.push(c); else lines.push({ y: c.y, parts: [c] });
    }
    pageOffset = maxY + 2000;
  }
  const wid = (str) => { let w = 0; for (const ch of str) w += /[mwMW]/.test(ch) ? 0.86 : /[A-Z0-9]/.test(ch) ? 0.66 : /[iljI.,;:!|'ft]/.test(ch) ? 0.3 : ch === ' ' ? 0.26 : /[r]/.test(ch) ? 0.36 : 0.53; return w; };
  // merge glyphs into words and columns: returns array of lines, each an array of column strings
  return lines.map(l => {
    l.parts.sort((a, b) => a.x - b.x);
    const cols = []; let cur = '', end = -1e9;
    for (const c of l.parts) {
      const gap = c.x - end;
      if (cur && gap > c.sz * 2.2) { cols.push(cur); cur = ''; }
      else if (cur && gap > c.sz * 0.16) cur += ' ';
      cur += c.s;
      end = c.x + wid(c.s) * c.sz;
    }
    if (cur) cols.push(cur);
    return cols.map(x => x.replace(/\s+/g, ' ').trim());
  });
}

// ---------- New patient wizard and Ready for Assessment ----------
var NP = (function () {
  var $id = function (id) { return document.getElementById(id); };
  var sleep = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var tc = function (s) { return String(s || '').toLowerCase().replace(/(^|[\s'-])([a-z])/g, function (a, b, c) { return b + c.toUpperCase(); }); };
  var STAFF = { mikaila: '133', james: '167', jimmy: '167', connie: '218', mendy: '92' };
  var CT = function () {
    var p = new Intl.DateTimeFormat('en-US', { timeZone: 'America/Chicago', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false }).formatToParts(new Date());
    var g = function (t) { return (p.find(function (x) { return x.type === t; }) || {}).value; };
    return { date: g('year') + '-' + g('month') + '-' + g('day'), time: (g('hour') === '24' ? '00' : g('hour')) + ':' + g('minute') };
  };
  var isoDate = function (s) {
    s = String(s || '').trim();
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) { var y = +m[3]; if (y < 100) y += y > 30 ? 1900 : 2000; return y + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2); }
    var d = new Date(s.replace(/\s+\d{1,2}:\d{2}.*$/, ''));
    if (!isNaN(d)) return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
    return '';
  };

  // ----- parse the Teams VOB -----
  function parseVob(raw) {
    var t = String(raw || '').replace(/\r/g, '').replace(/ /g, ' ');
    var line = function (re) { var m = t.match(re); return m ? m[1].replace(/\s+/g, ' ').trim() : ''; };
    var v = {};
    v.name = line(/^[ \t]*(?:CLIENT[ \t]+)?NAME[ \t]*:[ \t]*(.+)$/im);
    v.dob = line(/^[ \t]*DOB[ \t]*:?[ \t]*(\d{1,2}\/\d{1,2}\/\d{2,4})/im) || line(/DOB[ \t]*:?[ \t]*(\d{1,2}\/\d{1,2}\/\d{2,4})/i);
    v.phone = line(/^[ \t]*(?:PHONE|CELL|PHONE[ \t]*NUMBER)[ \t]*#?[ \t]*:[ \t]*(.+)$/im);
    v.address = line(/^[ \t]*ADDRESS[ \t]*:[ \t]*(.+)$/im);
    v.insurance = line(/^[ \t]*INSURANCE[ \t]+PROVIDER[ \t]*:[ \t]*(.+)$/im) || line(/^[ \t]*Primary[ \t]+Insurance[ \t]*:[ \t]*(.+)$/im) || line(/^[ \t]*INSURANCE[ \t]*:[ \t]*(.+)$/im);
    v.plan = line(/^[ \t]*PLAN[ \t]+NAME[ \t]*:[ \t]*(.+)$/im);
    v.planType = line(/^[ \t]*PLAN[ \t]+TYPE[ \t]*:[ \t]*(.+)$/im);
    v.medicaidId = line(/Medicaid[ \t]*(?:ID|#)[ \t]*#?[ \t]*:?[ \t]*(\d{10})\b/i);
    v.memberId = line(/MEMBER[ \t]*ID[ \t]*#?[ \t]*:?[ \t]*([A-Za-z0-9-]{5,})/i).replace(/-/g, '').toUpperCase();
    v.group = line(/^[ \t]*GROUP[ \t]*(?:#|NUMBER|NO\.?)[ \t]*:[ \t]*(.+)$/im);
    v.effective = line(/EFFECTIVE[ \t]+DATE[ \t]*:?[ \t]*(\d{1,2}\/\d{1,2}\/\d{4})/i);
    v.authReq = line(/AUTH(?:ORIZATION)?[ \t]+REQUIRED[ \t]*:?[ \t]*(Yes|No)/i);
    v.verified = line(/DATE[ \t]+VERIFIED[ \t]*:?[ \t]*(\d{1,2}\/\d{1,2}\/\d{4})/i);
    v.callRef = line(/CALL[ \t]+REF(?:ERENCE)?[ \t]*#?[ \t]*:?[ \t]*([A-Za-z0-9-]+)/i);
    v.rep = line(/INSURANCE[ \t]+REP[ \t]*:?[ \t]*(.+)/i);
    if (/^n\/?a$/i.test(v.phone)) v.phone = '';
    if (/^n\/?a$/i.test(v.group)) v.group = '';
    if (/^n\/?a$/i.test(v.address)) v.address = '';
    var am = v.address.match(/^(.*?)[ ,]+([A-Za-z .'-]+?)[ ,]+([A-Z]{2})[ ,]+(\d{5})(?:-\d{4})?$/);
    if (am) { v.street = tc(am[1].replace(/\s*,\s*$/, '')); v.city = tc(am[2]); v.state = am[3]; v.zip = am[4]; }
    else if (v.address) { v.street = v.address; }
    return v;
  }

  // ----- parse the prescreen (PDF lines or pasted text) -----
  function parsePrescreen(lines) {
    // lines: array of column arrays
    var norm = function (s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); };
    var get = function (label) {
      var k = norm(label);
      for (var i = 0; i < lines.length; i++) { var L = lines[i]; if (L.length > 1 && norm(L[0]).indexOf(k) === 0) return L.slice(1).join(' ').replace(/\s+/g, ' ').trim(); }
      return '';
    };
    var p = {};
    p.name = get('Client name'); p.dob = get('Date of birth'); p.insurance = get('Insurance');
    p.referral = get('Referral source'); p.screenedBy = get('Screened by'); p.housing = get('Housing'); p.bed = get('Confirm bed availability');
    p.program = get('Program wanted'); p.allergiesYN = get('Any allergies'); p.allergies = get('Which allergies') || get('Allergies');
    p.substances = get('Used in the last'); p.lastUse = get('Date of last use'); p.symptoms = get('Symptoms affecting daily life');
    p.pregnant = get('Pregnant or possibly'); p.phone = get('Client phone') || get('Phone');
    var g = (p.housing + ' ' + p.bed).toLowerCase();
    p.gender = /\bfemale\b/.test(g) ? 'F' : /\bmale\b/.test(g) ? 'M' : '';
    return p;
  }
  function textToLines(txt) {
    return String(txt || '').replace(/\r/g, '').split('\n').map(function (l) {
      var parts = l.split(/\t| {2,}| \| /).map(function (x) { return x.trim(); }).filter(Boolean);
      return parts;
    }).filter(function (x) { return x.length; });
  }

  // ----- contract map (Mendy's table, tenant findings 23 Nov 2026) -----
  function pickContract(v, demo) {
    var t = (v.insurance + ' ' + v.plan + ' ' + v.planType).toLowerCase();
    var res = { id: '', name: '', program: '', flags: [] };
    if (/medicare|empower|summit|\buhc\b|united|optum|\bumr\b|aetna|cigna|anthem/.test(t)) { res.stop = 'This plan (' + v.insurance + ') isn\'t in the contract table yet. Ask Mendy which contract to use.'; return res; }
    if (/caresource/.test(t)) { res.id = '70'; res.name = 'Caresource Arkansas'; res.program = 'PHP'; }
    else if (/total\s*care|\batc\b/.test(t)) { res.id = '60'; res.name = 'Arkansas Total Care'; res.program = 'PHP'; }
    else if (/ambetter|celtic|qualchoice/.test(t)) { res.id = '27'; res.name = 'Ambetter of Arkansas'; res.program = 'PHP'; }
    else if (/health\s*advantage|arkansas\s*blue|blue\s*cross.*arkansas|arkansas.*blue\s*cross|bcbs\s*(of\s*)?(ar|arkansas)\b|arhome.*(bcbs|blue)/.test(t)) { res.id = '3'; res.name = 'BCBS Arkansas'; res.program = 'PHP'; }
    else if (/medicaid|fee.for.service|\barkids\b/.test(t)) { res.id = '43'; res.name = 'Medicaid Arkansas'; res.program = 'IOP'; res.flags.push('The VOB says Medicaid only. A PASSE or ARHOME carrier would show in MMIS, not the VOB. If one shows up, the payer and program change (PHP).'); }
    else { res.stop = 'The VOB insurance (' + (v.insurance || 'blank') + ') doesn\'t match the contract table. Ask Mendy which contract to use.'; return res; }
    return res;
  }
  var PROGRAMS = { IOP: '2', PHP: '7' };

  // ----- UI -----
  var state = null;
  function box(html) {
    var b = $id('lsnp');
    if (!b) {
      b = document.createElement('div'); b.id = 'lsnp';
      b.style.cssText = 'position:fixed;top:40px;left:50%;transform:translateX(-50%);z-index:2147483647;background:#fff;border:3px solid #0f766e;border-radius:12px;padding:16px;width:min(720px,96vw);max-height:88vh;overflow:auto;font:15px system-ui,"Segoe UI",sans-serif;color:#111;box-shadow:0 8px 30px rgba(0,0,0,.3)';
      document.body.appendChild(b);
    }
    b.innerHTML = '<div style="display:flex;justify-content:space-between;align-items:center"><b style="font-size:18px">New patient</b><button type="button" id="lsnpx" style="font-size:22px;border:0;background:none;cursor:pointer" title="Close">&times;</button></div>' + html;
    $id('lsnpx').onclick = function () { if (state && state.running) { alert('Still working. Please wait until it finishes.'); return; } b.remove(); };
    return b;
  }
  var btnCss = 'padding:9px 16px;border-radius:8px;border:0;background:#0f766e;color:#fff;font-weight:600;cursor:pointer;font-size:15px';

  function openWizard() {
    state = { vob: null, pre: null, running: false };
    box('<p style="margin:8px 0">1. In Teams, open the VOB reply in <b>Crosshope VOB-Insurance-Billing</b>, select the whole message, copy it, and paste it here.</p>' +
      '<textarea id="lsnpv" rows="5" style="width:100%;box-sizing:border-box;border:2px solid #99f6e4;border-radius:8px;padding:8px" placeholder="Paste the VOB here (Ctrl+V)"></textarea>' +
      '<p style="margin:12px 0 6px">2. Drop the prescreen PDF here (from the Outlook email or OneDrive).</p>' +
      '<div id="lsnpd" style="border:2px dashed #0f766e;border-radius:8px;padding:18px;text-align:center;color:#0f766e;cursor:pointer">Drop the prescreen PDF here, or click to pick it</div>' +
      '<input type="file" id="lsnpf" accept="application/pdf" style="display:none">' +
      '<div id="lsnpm" style="margin:10px 0;color:#555;font-size:14px">The VOB is used for name, date of birth, address and payer. The prescreen is used for everything else.</div>' +
      '<button type="button" id="lsnpgo" style="' + btnCss + '">Check and review</button>');
    var drop = $id('lsnpd'), file = $id('lsnpf');
    drop.onclick = function () { file.click(); };
    drop.ondragover = function (e) { e.preventDefault(); drop.style.background = '#f0fdfa'; };
    drop.ondragleave = function () { drop.style.background = ''; };
    drop.ondrop = function (e) { e.preventDefault(); drop.style.background = ''; if (e.dataTransfer.files[0]) readPdf(e.dataTransfer.files[0]); };
    file.onchange = function () { if (file.files[0]) readPdf(file.files[0]); };
    $id('lsnpgo').onclick = review;
  }
  async function readPdf(f) {
    var drop = $id('lsnpd');
    drop.textContent = 'Reading ' + f.name + '...';
    try {
      var buf = new Uint8Array(await f.arrayBuffer());
      var inflate = async function (u) { var ds = new DecompressionStream('deflate'); var out = new Response(new Blob([u]).stream().pipeThrough(ds)); return new Uint8Array(await out.arrayBuffer()); };
      var lines = await pdfText(buf, inflate);
      state.pre = parsePrescreen(lines);
      if (!state.pre.name && !state.pre.screenedBy) throw new Error('not a prescreen');
      drop.textContent = 'Prescreen read: ' + (state.pre.name || '(no name)') + ', screened by ' + (state.pre.screenedBy || '?') + '.';
      drop.style.borderStyle = 'solid';
    } catch (e) {
      state.pre = null;
      drop.innerHTML = 'Couldn\'t read that file as a prescreen. Open the PDF in Chrome, press Ctrl+A then Ctrl+C, and paste it here instead:<textarea id="lsnpt" rows="3" style="width:100%;margin-top:8px"></textarea>';
      $id('lsnpt').oninput = function () { state.pre = parsePrescreen(textToLines($id('lsnpt').value)); };
    }
  }

  async function review() {
    var v = parseVob($id('lsnpv').value);
    var m = $id('lsnpm');
    if (!v.name || !v.dob) { m.innerHTML = '<span style="color:#b91c1c">The VOB paste needs at least NAME and DOB lines. Copy the whole VOB reply from Teams.</span>'; return; }
    var pre = state.pre || {};
    var nm = v.name.split(/\s+/);
    var first = tc(nm[0]), last = tc(nm[nm.length - 1]);
    var demo = /^demo$/i.test(first);
    var c = pickContract(v, demo);
    var warn = [];
    if (!state.pre) warn.push('No prescreen was read. Gender, allergies and referral source must be filled in below.');
    if (pre.name && pre.name.toLowerCase().replace(/[^a-z]/g, '').indexOf(last.toLowerCase().replace(/[^a-z]/g, '')) < 0) warn.push('The prescreen name (' + pre.name + ') doesn\'t match the VOB. The VOB spelling is used.');
    if (pre.dob && isoDate(pre.dob) && isoDate(pre.dob) !== isoDate(v.dob)) warn.push('The prescreen date of birth doesn\'t match the VOB. The VOB date is used.');
    if (/yes/i.test(v.authReq)) warn.push('The VOB says authorization is required. Get it before services.');
    c.flags.forEach(function (f) { warn.push(f); });
    // duplicate check (census and open inquiries, by name)
    m.textContent = 'Checking for an existing record...';
    var dups = [];
    try {
      var t = document.querySelector('meta[name="csrf-token"]').content;
      var post = async function (u, o) { var b = new URLSearchParams(Object.assign({ _token: t }, o)); var r = await fetch(u, { method: 'POST', credentials: 'same-origin', headers: { 'X-CSRF-TOKEN': t, 'X-Requested-With': 'XMLHttpRequest', 'Accept': 'application/json' }, body: b }); return r.json(); };
      var strip = function (s) { return String(s == null ? '' : s).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(); };
      var cen = await post('/home/census', { draw: 1, start: 0, length: 2000, 'programfilter[]': 0 });
      var inq = await post('/home/inquirys', { draw: 1, start: 0, length: 1000 });
      var L = last.toLowerCase(), F = first.toLowerCase();
      (cen.data || []).forEach(function (r) { var s = strip(r.lnamefname || r.name).toLowerCase(); if (s.indexOf(L) > -1 && s.indexOf(F.slice(0, 3)) > -1) dups.push('On the census: ' + strip(r.lnamefname || r.name) + ' (MRN ' + strip(r.mrn) + ', ' + strip(r.pgmname) + ')'); });
      (inq.data || []).forEach(function (r) { var s = strip(r.fnamelname || r.name).toLowerCase(); if (s.indexOf(L) > -1 && s.indexOf(F.slice(0, 3)) > -1) dups.push('Open inquiry: ' + strip(r.fnamelname || r.name) + ' (inquiry ' + r.id + ')'); });
    } catch (e) { warn.push('Couldn\'t check the census for a duplicate. LightningStep will still check on save.'); }
    var now = CT();
    var gender = pre.gender || '';
    var allergies = /^\s*no\b/i.test(pre.allergiesYN || '') ? 'NKDA' : (pre.allergies || '');
    var suggestSud = !!(pre.substances && !/^\s*none|^\s*no\b/i.test(pre.substances));
    state.data = { first: first, last: last, demo: demo, v: v, pre: pre, contract: c, suggestSud: suggestSud };
    var row = function (label, html) { return '<tr><td style="padding:4px 8px 4px 0;color:#555;white-space:nowrap;vertical-align:top">' + label + '</td><td style="padding:4px 0">' + html + '</td></tr>'; };
    var inp = function (id, val, w) { return '<input id="' + id + '" value="' + esc(val) + '" style="width:' + (w || '100%') + ';box-sizing:border-box;padding:5px;border:1px solid #bbb;border-radius:6px">'; };
    var dupHtml = dups.length ? '<div style="background:#fef2f2;border:2px solid #b91c1c;border-radius:8px;padding:10px;margin:8px 0"><b>This person may already be in LightningStep:</b><br>' + dups.map(esc).join('<br>') + '<br><label><input type="checkbox" id="lsnpdupok"> I checked. This is a different person.</label></div>' : '';
    var stopHtml = c.stop ? '<div style="background:#fef2f2;border:2px solid #b91c1c;border-radius:8px;padding:10px;margin:8px 0"><b>Stopped:</b> ' + esc(c.stop) + '</div>' : '';
    box(dupHtml + stopHtml +
      (warn.length ? '<div style="background:#fffbeb;border:1px solid #d97706;border-radius:8px;padding:8px 10px;margin:8px 0;font-size:14px">' + warn.map(function (w) { return '&bull; ' + esc(w); }).join('<br>') + '</div>' : '') +
      '<p style="margin:6px 0;font-size:14px;color:#555">Check everything. Fix anything here. Nothing is saved until you click <b>Save to LightningStep</b>.</p>' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
      row('First name', inp('np_first', first)) + row('Last name', inp('np_last', last)) +
      row('Date of birth', '<input id="np_dob" type="date" value="' + esc(isoDate(v.dob)) + '" style="padding:5px">') +
      row('Gender', '<select id="np_gender" style="padding:5px"><option value="">Pick one</option><option value="F"' + (gender === 'F' ? ' selected' : '') + '>Female</option><option value="M"' + (gender === 'M' ? ' selected' : '') + '>Male</option></select>' + (gender ? ' <span style="color:#555">(from the prescreen housing answer)</span>' : '')) +
      row('Phone', inp('np_phone', (v.phone || pre.phone || '').replace(/[^\d]/g, ''), '160px')) +
      row('Street', inp('np_street', v.street || '')) +
      row('City, state, ZIP', inp('np_city', v.city || '', '45%') + ' ' + inp('np_state', v.state || 'AR', '60px') + ' ' + inp('np_zip', v.zip || '', '90px')) +
      row('Payer', c.stop ? '<b style="color:#b91c1c">Not set</b>' : '<b>' + esc(c.name) + '</b> (contract ' + c.id + ') <span style="color:#555">from the VOB: ' + esc(v.insurance) + (v.plan ? ', ' + esc(v.plan) : '') + '</span>') +
      row(c.id === '43' ? 'Medicaid ID' : 'Member ID', inp('np_member', c.id === '43' ? (v.medicaidId || v.memberId) : (v.memberId || v.medicaidId), '220px')) +
      row('Group number', inp('np_group', v.group || '', '220px')) +
      row('Effective date', '<input id="np_eff" type="date" value="' + esc(isoDate(v.effective)) + '" style="padding:5px">') +
      row('Auth required', '<select id="np_auth" style="padding:5px"><option' + (/yes/i.test(v.authReq) ? ' selected' : '') + '>Yes</option><option' + (/yes/i.test(v.authReq) ? '' : ' selected') + '>No</option></select>') +
      row('Program', '<select id="np_prog" style="padding:5px"><option value="IOP"' + (c.program === 'IOP' ? ' selected' : '') + '>IOP</option><option value="PHP"' + (c.program === 'PHP' ? ' selected' : '') + '>PHP</option></select> <span style="color:#555">' + (c.program ? 'CrossHope rule: ' + (c.id === '43' ? 'Arkansas Medicaid goes to IOP' : 'this payer goes to PHP') : '') + (pre.program ? '. Prescreen asked for: ' + esc(pre.program) : '') + '</span>' + (demo ? ' <b style="color:#0f766e">(test patient: no-charge test programs)</b>' : '')) +
      row('Admission', '<input id="np_adate" type="date" value="' + now.date + '" style="padding:5px"> <input id="np_atime" type="time" value="' + now.time + '" style="padding:5px"> <span style="color:#555">Can\'t be changed after saving.</span>') +
      row('Allergies', inp('np_allergy', allergies) + '<div style="color:#555;font-size:12px">NKDA, or a list separated by commas. Each one is saved as a medication allergy with reaction "Unknown": fix the type and reaction on the chart if you know them.</div>') +
      row('Referral source', inp('np_ref', pre.referral || '')) +
      row('Screened by', inp('np_screen', pre.screenedBy || '')) +
      row('Emergency contact', inp('np_ecname', '', '45%') + ' ' + inp('np_ecrel', '', '25%') + ' ' + inp('np_ecphone', '', '25%') + '<div style="color:#555;font-size:12px">Optional: name, relationship, phone.</div>') +
      '</table>' +
      '<div id="lsnpm" style="margin:10px 0;font-size:14px"></div>' +
      '<button type="button" id="lsnpback" style="padding:9px 14px;border-radius:8px;border:1px solid #999;background:#fff;cursor:pointer">Back</button> ' +
      (c.stop ? '' : '<button type="button" id="lsnprun" style="' + btnCss + '">Save to LightningStep</button>'));
    $id('lsnpback').onclick = openWizard;
    if ($id('lsnprun')) $id('lsnprun').onclick = function () {
      if (dups.length && !($id('lsnpdupok') && $id('lsnpdupok').checked)) { $id('lsnpm').innerHTML = '<span style="color:#b91c1c">Tick the box in the red section first, or close this and open the existing record.</span>'; return; }
      run();
    };
  }

  // ----- the run -----
  function frame(url) {
    return new Promise(function (res, rej) {
      var f = document.createElement('iframe');
      f.style.cssText = 'position:fixed;left:-4000px;top:0;width:1280px;height:1000px;border:0';
      var t = setTimeout(function () { rej(new Error('LightningStep didn\'t load ' + url)); }, 60000);
      f.onload = function () { clearTimeout(t); res(f); };
      f.src = url; document.body.appendChild(f);
    });
  }
  function waitUrl(f, re, ms) {
    return new Promise(function (res, rej) {
      var end = Date.now() + (ms || 60000);
      (function poll() {
        var u = ''; try { u = f.contentWindow.location.pathname + f.contentWindow.location.search; } catch (e) {}
        var mm = u.match(re); if (mm) { setTimeout(function () { res(mm); }, 800); return; }
        if (Date.now() > end) { rej(new Error('Save didn\'t finish (page: ' + u + ')')); return; }
        setTimeout(poll, 400);
      })();
    });
  }
  async function idle(w, ms) { var end = Date.now() + (ms || 15000); await sleep(300); while (Date.now() < end && w.jQuery && w.jQuery.active > 0) await sleep(250); await sleep(300); }
  function setv(w, sel, val, root) {
    var $ = w.jQuery; var el = root ? $(root).find(sel) : $(sel);
    if (!el.length) throw new Error('missing field ' + sel);
    el.val(val).trigger('input').trigger('change').trigger('blur');
    return el.val();
  }
  function popupNo(w) {
    var d = w.document; var p = d.querySelector('.swal2-popup');
    if (p && getComputedStyle(p).display !== 'none') { var b = p.querySelector('.swal2-cancel') || p.querySelector('.swal2-deny'); if (b && getComputedStyle(b).display !== 'none') { b.click(); return true; } }
    return false;
  }

  async function run() {
    var D = state.data, g = function (id) { return ($id(id).value || '').trim(); };
    var f = {
      first: tc(g('np_first')), last: tc(g('np_last')), dob: g('np_dob'), gender: g('np_gender'), phone: g('np_phone').replace(/\D/g, ''),
      street: g('np_street'), city: g('np_city'), state: g('np_state').toUpperCase(), zip: g('np_zip'),
      member: g('np_member').replace(/[\s-]/g, '').toUpperCase(), group: g('np_group'), eff: g('np_eff'), auth: g('np_auth'), prog: g('np_prog'),
      adate: g('np_adate'), atime: g('np_atime'), allergy: g('np_allergy'), ref: g('np_ref'), screen: g('np_screen'),
      ecname: g('np_ecname'), ecrel: g('np_ecrel'), ecphone: g('np_ecphone').replace(/\D/g, '')
    };
    var miss = [];
    ['first', 'last', 'dob', 'gender', 'member', 'eff', 'adate', 'atime'].forEach(function (k) { if (!f[k]) miss.push(k); });
    if (miss.length) { $id('lsnpm').innerHTML = '<span style="color:#b91c1c">Fill in: ' + miss.join(', ') + '</span>'; return; }
    var demo = /^demo$/i.test(f.first);
    var sub = PROGRAMS[f.prog];
    var daysBack = (new Date(CT().date) - new Date(f.adate)) / 864e5;
    if (daysBack < 0 || daysBack > 14) { $id('lsnpm').innerHTML = '<span style="color:#b91c1c">The admission date must be today or within the last 14 days.</span>'; return; }
    state.running = true;
    var steps = ['Inquiry with name, date of birth, gender, phone and address', 'Payer (' + D.contract.name + ')', 'Admission to ' + f.prog, 'Allergies'];
    if (f.ecname) steps.push('Emergency contact');
    steps.push('Read back');
    var made = {};
    box('<p style="margin:6px 0">Saving. Don\'t close this box or leave the page.</p><ol id="lsnps" style="margin:8px 0 8px 18px;padding:0;line-height:1.8">' + steps.map(function (s, i) { return '<li id="lsnps' + i + '">' + esc(s) + ' <span></span></li>'; }).join('') + '</ol><div id="lsnpm" style="margin:10px 0"></div>');
    var mark = function (i, ok, note) { var li = $id('lsnps' + i); if (!li) return; li.style.color = ok === true ? '#166534' : ok === false ? '#b91c1c' : '#111'; li.querySelector('span').textContent = ok === true ? '✓ ' + (note || '') : ok === false ? '✗ ' + (note || '') : '… ' + (note || ''); };
    var note = 'Entered with LS Forms from the Teams VOB' + (D.v.verified ? ' (verified ' + D.v.verified + ')' : '') + ' and the prescreen' + (f.screen ? ' by ' + f.screen : '') + '. Referral: ' + (f.ref || 'not given') + '.' + (D.contract.flags.length ? ' ' + D.contract.flags.join(' ') : '');
    var i = 0, fr;
    try {
      // 1. inquiry
      mark(i, null, 'working');
      fr = await frame('/ice/create');
      var w = fr.contentWindow; await idle(w);
      var sc = (f.screen || 'CrossHope Staff').split(/\s+/);
      setv(w, '#callerfname', tc(sc[0])); setv(w, '#callerlname', tc(sc.slice(1).join(' ') || 'Staff'));
      setv(w, '#reltoclient', 'Other'); await sleep(300); setv(w, '#reltoclientother', 'CrossHope staff (prescreen)');
      setv(w, '#staffreferral_id', STAFF[(sc[0] || '').toLowerCase()] || '0');
      setv(w, '#fname', f.first); setv(w, '#lname', f.last); setv(w, '#dob', f.dob);
      setv(w, '#gender_id', f.gender === 'F' ? '93' : '92'); setv(w, '#legal_sex', f.gender);
      if (f.phone) { w.jQuery('#cphone').val(f.phone).trigger('input').trigger('change').trigger('blur'); setv(w, '#compref', 'Cell Phone'); }
      if (f.street) { setv(w, '#address', f.street); setv(w, '#city', f.city); setv(w, '#usstate', f.state || 'AR'); setv(w, '#zip', f.zip); }
      setv(w, '#location_id', '21'); await sleep(1200);
      for (var k = 0; k < 30 && !w.document.querySelector('#potentialpgm_id option[value="' + sub + '"]'); k++) await sleep(300);
      setv(w, '#potentialpgm_id', sub); setv(w, '#projdtadmission', f.adate); setv(w, '#additionalnote', note);
      await sleep(2000); await idle(w);
      var fd = new FormData(w.document.querySelector('#frmice'));
      var want = { fname: f.first, lname: f.last, dob: f.dob, legal_sex: f.gender, location_id: '21', potentialpgm_id: sub };
      var bad = Object.keys(want).filter(function (k2) { return (fd.get(k2) || '').trim() !== want[k2]; });
      if (fd.get('fundingsrc')) bad.push('funding source set');
      if (bad.length) throw new Error('Inquiry form check failed: ' + bad.join(', '));
      var dupRows = [].filter.call(w.document.querySelectorAll('#tbldups tbody tr'), function (tr) { return tr.cells.length > 1 && !/no data|no matching/i.test(tr.textContent); }).length;
      if (dupRows && !($id('lsnpdupok') || {}).checked) throw new Error('LightningStep found a possible duplicate. Nothing was saved.');
      if (w.document.getElementById('btnsave').disabled) throw new Error('LightningStep blocked the save (possible duplicate). Nothing was saved.');
      w.document.getElementById('btnsave').click();
      for (var k3 = 0; k3 < 20; k3++) { await sleep(500); if (popupNo(w)) break; if (!w.document.getElementById('btnsave')) break; }
      var mInq = await waitUrl(fr, /\/inquiry\/show\/(\d+)/);
      made.inquiry = mInq[1];
      var html = w.document.documentElement.innerHTML;
      var pc = html.match(/\/payor\/create\/(\d+)\/(\d+)/);
      made.episode = pc ? pc[2] : '';
      fr.remove(); mark(i, true, 'inquiry ' + made.inquiry); i++;
      if (!made.episode) throw new Error('Couldn\'t find the new episode for the payer.');

      // 2. payer
      mark(i, null, 'working');
      fr = await frame('/payor/create/' + made.inquiry + '/' + made.episode);
      w = fr.contentWindow; await idle(w);
      var $ = w.jQuery;
      var frm = [].find.call(w.document.forms, function (x) { return /\/payor$/.test(x.getAttribute('action') || ''); });
      if (!frm) throw new Error('No payer form');
      $('#collapseOne,#collapseThree,#bvf').collapse('show');
      setv(w, '#insreltoclient', 'Self', frm); for (var k4 = 0; k4 < 25 && !$('#insfname').val(); k4++) await sleep(300);
      setv(w, '#fundingsrc', 'Insurance', frm); await sleep(900);
      setv(w, '#mccontract_id', D.contract.id, frm); for (var k5 = 0; k5 < 30 && !($('#payorname').val() && $.active === 0); k5++) await sleep(300); await sleep(600);
      setv(w, '#planpriority', '1', frm); await sleep(500);
      var sw = w.document.querySelector('.swal2-popup'); if (sw && getComputedStyle(sw).display !== 'none') { var cf = sw.querySelector('.swal2-confirm'); if (cf) cf.click(); }
      setv(w, '#policyhomestate', 'AR', frm);
      $('#memberidnumber').val(f.member).trigger('input').trigger('change').trigger('blur');
      setv(w, '#dateeffective', f.eff, frm);
      if (f.group && frm.querySelector('#groupnumber')) $(frm).find('#groupnumber').val(f.group).trigger('change');
      var o = function (n, val) { var el = frm.querySelector('#' + n); if (el) $(el).val(val).trigger('change'); };
      o('contractyn', 'Yes'); o('eligibleyn', 'Yes'); o('preauthyn', f.auth); o('groupnumreq', f.group ? 'Yes' : 'No'); o('labyn', 'Yes'); o('otherpayoryn', 'No');
      $(frm.querySelector('select#vstatus')).val('Verified').trigger('change');
      var iv = frm.querySelector('[name=isverified][type=checkbox]'); if (iv && !iv.checked) iv.click();
      setv(w, '#verifyresult', 'Teams VOB' + (D.v.verified ? ' verified ' + D.v.verified : '') + (D.v.rep ? ' by ' + D.v.rep : '') + (D.v.callRef ? ', call ref ' + D.v.callRef : '') + '. Entered with LS Forms.', frm);
      var sober = frm.querySelector('[name=issoberliving]'); if (sober && sober.checked) sober.click();
      await sleep(600);
      var pfd = new FormData(frm), pg = function (k6) { return (pfd.get(k6) || '').trim(); };
      var pbad = [];
      if (pg('mccontract_id') !== D.contract.id) pbad.push('contract'); if (pg('planpriority') !== '1') pbad.push('priority');
      if (pg('memberidnumber') !== f.member) pbad.push('member id'); if (pg('insreltoclient') !== 'Self') pbad.push('relationship');
      if (!pg('insdob')) pbad.push('insured date of birth'); if (pfd.has('issoberliving')) pbad.push('sober living');
      if (pbad.length) throw new Error('Payer form check failed: ' + pbad.join(', '));
      var payList = async function () { var ph = await (await fetch('/inquiry/show/' + made.inquiry + '?section=payors', { credentials: 'same-origin' })).text(); return (ph.match(/\/payor\/(\d+)\/edit/g) || []).map(function (x) { return x.match(/\d+/)[0]; }); };
      var p0 = await payList();
      frm.querySelector('.SavePayor').click();
      var newPayer = null;
      for (var k7 = 0; k7 < 60 && !newPayer; k7++) { await sleep(700); popupNo(w); var p1 = await payList(); newPayer = p1.filter(function (x) { return p0.indexOf(x) < 0; })[0] || null; }
      if (!newPayer) throw new Error('The payer didn\'t save. Check the Payers tab before trying again.');
      made.payer = newPayer;
      fr.remove(); mark(i, true, 'payer ' + newPayer); i++;

      // 3. admission
      mark(i, null, 'working');
      fr = await frame('/admission/create/' + made.inquiry);
      w = fr.contentWindow; await idle(w);
      var af = [].find.call(w.document.forms, function (x) { return /\/admission\/store$/.test(x.getAttribute('action') || ''); });
      if (!af) { var bodyTxt = w.document.body.innerText.slice(0, 200); throw new Error('Admission page didn\'t open: ' + bodyTxt); }
      var subSel = af.querySelector('[name=subprogram_id]');
      if (!subSel || subSel.value !== sub) throw new Error('The admission page shows a different program (' + (subSel ? subSel.value : 'none') + '). Nothing was admitted.');
      if (demo) throw new Error('Test patient: stopped before the admission. Test patients are admitted through the connector to the no-charge test program.');
      var dEl = af.querySelector('[name=dateadmission]'), tEl = af.querySelector('[name=timeadmission]');
      if (dEl) w.jQuery(dEl).val(f.adate).trigger('change');
      if (tEl) w.jQuery(tEl).val(f.atime).trigger('change');
      await idle(w);
      var afd = new FormData(af);
      var pgmVal = afd.get('subprogram_id') || afd.get('potentialpgm_id') || afd.get('program_id');
      made.admitProgram = pgmVal;
      var saveBtn = [].find.call(af.querySelectorAll('button,input[type=submit]'), function (b) { return /save/i.test(b.textContent || b.value); });
      if (!saveBtn) throw new Error('No Save button on the admission page');
      saveBtn.click();
      for (var k8 = 0; k8 < 40; k8++) {
        await sleep(500);
        var u8 = ''; try { u8 = w.location.pathname; } catch (e) {}
        if (/\/client\/index\//.test(u8)) break;
        var ap = [].find.call(w.document.querySelectorAll('button,a'), function (b) { return /approve change/i.test(b.textContent) && b.offsetParent; });
        if (ap) ap.click();
        var err = w.document.querySelector('.alert-danger:not(.hidden), .help-block.with-errors li');
        if (err && err.offsetParent && /required/i.test(err.textContent)) throw new Error('Admission refused: ' + err.textContent.trim().slice(0, 160));
      }
      var mCl = await waitUrl(fr, /\/client\/index\/(\d+)/);
      made.client = mCl[1];
      fr.remove(); mark(i, true, 'chart ' + made.client); i++;

      // 4. allergies
      mark(i, null, 'working');
      fr = await frame('/client/index/' + made.client + '?section=medicalassess');
      w = fr.contentWindow; await idle(w); $ = w.jQuery;
      var rows = function () { return [].filter.call(w.document.querySelectorAll('#taballergies table tbody tr'), function (tr) { return !/No allergies to display/i.test(tr.textContent); }).map(function (tr) { return tr.textContent.replace(/\s+/g, ' ').trim(); }); };
      var items = /^(nkda|none|no)$/i.test(f.allergy) || !f.allergy ? ['NKDA'] : f.allergy.split(/\s*[,;]\s*/).filter(Boolean);
      for (var a = 0; a < items.length; a++) {
        w.document.querySelector('#btnaddallergy').click();
        var shown = function () { var mo = w.document.querySelector('#mdl_allergy_add'); return mo && $(mo).is(':visible') && /\b(in|show)\b/.test(mo.className) && w.getComputedStyle(mo).opacity === '1'; };
        for (var k9 = 0; k9 < 40 && !shown(); k9++) await sleep(250); await sleep(800);
        var n0 = rows().length;
        if (items[a] === 'NKDA') { $('#mdl_allergy_add_allergytype').val('NKDA').trigger('change'); }
        else {
          $('#mdl_allergy_add_allergytype').val('Medication').trigger('change'); await sleep(300);
          $('#mdl_allergy_add_substance').val(items[a]).trigger('input').trigger('change');
          $('#mdl_allergy_add_reaction').val('Unknown (from prescreen)').trigger('input').trigger('change');
          $('#mdl_allergy_add_howaddress').val('Ask the patient at intake').trigger('input').trigger('change');
        }
        await sleep(300);
        w.document.querySelector('#btn_mdl_allergy_add_save').click();
        for (var k10 = 0; k10 < 50 && rows().length <= n0; k10++) await sleep(300);
        if (rows().length <= n0) throw new Error('Allergy "' + items[a] + '" didn\'t save. Add it on the chart.');
      }
      fr.remove(); mark(i, true, items.join(', ')); i++;

      // 5. emergency contact (optional)
      if (f.ecname) {
        mark(i, null, 'working');
        fr = await frame('/contact/create/' + made.inquiry);
        w = fr.contentWindow; await idle(w); $ = w.jQuery;
        var cf2 = [].find.call(w.document.forms, function (x) { return /\/contact$/.test(x.getAttribute('action') || ''); });
        if (!cf2) throw new Error('No contact form');
        var en = f.ecname.split(/\s+/);
        var sv = function (n, val) { var el = cf2.querySelector('[name="' + n + '"]'); if (el) $(el).val(val).trigger('input').trigger('change'); return !!el; };
        sv('fname', tc(en[0])); sv('lname', tc(en.slice(1).join(' ')));
        sv('hphone', f.ecphone); sv('cphone', f.ecphone); sv('usstate', 'AR');
        sv('reltoclient', tc(f.ecrel || 'Emergency contact'));
        var ec = cf2.querySelector('[name=isemergencycontact]'); if (ec && !ec.checked) ec.click();
        var csb = [].find.call(cf2.querySelectorAll('button,input[type=submit]'), function (b) { return /save/i.test(b.textContent || b.value); });
        csb.click();
        await waitUrl(fr, /\/client\/index\/\d+|\/inquiry\/show\/\d+/);
        fr.remove(); mark(i, true); i++;
      }

      // 6. read back
      mark(i, null, 'working');
      var ch = await (await fetch('/client/index/' + made.client, { credentials: 'same-origin' })).text();
      var okName = ch.indexOf(f.last) > -1;
      mark(i, okName, okName ? 'chart opens' : 'check the chart');
      state.running = false;
      sessionStorage.setItem('lsnp-last', JSON.stringify(made));
      var sudFirst = D.suggestSud;
      $id('lsnpm').innerHTML = '<div style="background:#f0fdf4;border:2px solid #166534;border-radius:8px;padding:12px"><b>Done. ' + esc(f.first + ' ' + f.last) + ' is admitted to ' + f.prog + '.</b><br>Next: the primary diagnosis, then the Biopsych.' +
        '<div style="margin-top:10px"><b>Ready for Assessment:</b> ' +
        '<button type="button" id="lsnpmh" style="' + btnCss + ';background:' + (sudFirst ? '#64748b' : '#0f766e') + '">MH Biopsych</button> ' +
        '<button type="button" id="lsnpsud" style="' + btnCss + ';background:' + (sudFirst ? '#0f766e' : '#64748b') + '">SUD Biopsych</button>' +
        '<div style="font-size:13px;color:#555;margin-top:6px">' + (sudFirst ? 'The prescreen lists recent substance use, so SUD is suggested.' : 'No recent substance use on the prescreen, so MH is suggested.') + ' Opening one starts its draft.</div></div>' +
        '<div style="margin-top:8px"><a href="/client/index/' + made.client + '">Open the chart</a></div></div>';
      $id('lsnpmh').onclick = function () { openForm(made.client, 'Mental Health Biopsychosocial'); };
      $id('lsnpsud').onclick = function () { openForm(made.client, 'Biopsychosocial Assessment AR'); };
    } catch (e) {
      state.running = false;
      try { if (fr) fr.remove(); } catch (x) {}
      mark(i, false, e.message);
      var list = Object.keys(made).map(function (k11) { return k11 + ' ' + made[k11]; }).join(', ');
      $id('lsnpm').innerHTML = '<div style="background:#fef2f2;border:2px solid #b91c1c;border-radius:8px;padding:12px"><b>Stopped.</b> ' + esc(e.message) + '<br>' + (list ? 'Already saved: ' + esc(list) + '. <b>Don\'t run New patient again for this person.</b> Finish the rest on the chart or ask Mendy.' : 'Nothing was saved.') + '</div>';
    }
  }

  async function openForm(client, title) {
    var h = await (await fetch('/client/index/' + client, { credentials: 'same-origin' })).text();
    var d = new DOMParser().parseFromString(h, 'text/html');
    var el = [].filter.call(d.querySelectorAll('[data-url]'), function (e) { return e.textContent.replace(/\s+/g, ' ').trim() === title; })[0];
    if (!el) { alert('Couldn\'t find ' + title + ' on this chart.'); return; }
    location.href = el.getAttribute('data-url');
  }

  // Ready for Assessment on a chart or inquiry page
  async function readyCheck() {
    var path = location.pathname, client = (path.match(/\/client\/index\/(\d+)/) || [])[1];
    if (!client) {
      var inq = (path.match(/\/inquiry\/show\/(\d+)/) || [])[1];
      if (inq) { var ih = await (await fetch('/inquiry/show/' + inq, { credentials: 'same-origin' })).text(); client = (ih.match(/\/client\/index\/(\d+)/) || [])[1];
        if (!client) { box('<p>This person isn\'t admitted yet, so the Biopsych can\'t open. ' + (/gender/i.test(ih) ? '' : '') + 'Admit them first (an admission needs gender and date of birth on the record).</p>'); return; } }
    }
    if (!client) { box('<p>Open the patient\'s chart first, then click Ready for Assessment.</p>'); return; }
    var h = await (await fetch('/client/index/' + client, { credentials: 'same-origin' })).text();
    var d = new DOMParser().parseFromString(h, 'text/html');
    var has = function (t) { return [].some.call(d.querySelectorAll('[data-url]'), function (e) { return e.textContent.replace(/\s+/g, ' ').trim() === t; }); };
    box('<p style="margin:8px 0"><b>Ready for Assessment</b>: pick the Biopsych. Opening one starts its draft on this chart.</p>' +
      (has('Mental Health Biopsychosocial') ? '<button type="button" id="lsnpmh" style="' + btnCss + '">MH Biopsych</button> ' : '') +
      (has('Biopsychosocial Assessment AR') ? '<button type="button" id="lsnpsud" style="' + btnCss + '">SUD Biopsych</button>' : '') +
      '<div style="font-size:13px;color:#555;margin-top:8px">MH Biopsych for a mental health primary diagnosis, SUD Biopsych for substance use.</div>');
    if ($id('lsnpmh')) $id('lsnpmh').onclick = function () { openForm(client, 'Mental Health Biopsychosocial'); };
    if ($id('lsnpsud')) $id('lsnpsud').onclick = function () { openForm(client, 'Biopsychosocial Assessment AR'); };
  }

  return { openWizard: openWizard, readyCheck: readyCheck, parseVob: parseVob, parsePrescreen: parsePrescreen, pickContract: pickContract };
})();


// ---------- Discharge and Level of care ----------
var LC = (function () {
  var $id = function (id) { return document.getElementById(id); };
  var sleep = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var strip = function (s) { return String(s == null ? '' : s).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(); };
  var tok = function () { return document.querySelector('meta[name="csrf-token"]').content; };
  var H = { 'X-Requested-With': 'XMLHttpRequest', 'Accept': 'application/json' };
  var post = async function (u, o) { var t = tok(); var r = await fetch(u, { method: 'POST', credentials: 'same-origin', headers: Object.assign({ 'X-CSRF-TOKEN': t }, H), body: new URLSearchParams(Object.assign({ _token: t }, o)) }); return r.json(); };
  var getJ = async function (u) { var r = await fetch(u, { credentials: 'same-origin', headers: H }); return r.json(); };
  var CT = function (addDays) {
    var d = new Date(Date.now() + (addDays || 0) * 864e5);
    var p = new Intl.DateTimeFormat('en-US', { timeZone: 'America/Chicago', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false }).formatToParts(d);
    var g = function (t) { return (p.find(function (x) { return x.type === t; }) || {}).value; };
    return { date: g('year') + '-' + g('month') + '-' + g('day'), time: (g('hour') === '24' ? '00' : g('hour')) + ':' + g('minute') };
  };
  var mdy = function (iso) { var p = String(iso).split('-'); return p[1] + '/' + p[2] + '/' + p[0]; };
  var PROG = { '2': { name: 'IOP', svc: '20' }, '7': { name: 'PHP', svc: '219' }, '10': { name: 'IOP. 1 (test)', svc: '230' }, '13': { name: 'PHP. (test)', svc: '233' } };
  var btnCss = 'padding:9px 16px;border-radius:8px;border:0;background:#0f766e;color:#fff;font-weight:600;cursor:pointer;font-size:15px';
  var state = { running: false };

  function box(title, html, color) {
    var b = $id('lslc');
    if (!b) {
      b = document.createElement('div'); b.id = 'lslc';
      b.style.cssText = 'position:fixed;top:40px;left:50%;transform:translateX(-50%);z-index:2147483647;background:#fff;border:3px solid ' + (color || '#b45309') + ';border-radius:12px;padding:16px;width:min(680px,96vw);max-height:88vh;overflow:auto;font:15px system-ui,"Segoe UI",sans-serif;color:#111;box-shadow:0 8px 30px rgba(0,0,0,.3)';
      document.body.appendChild(b);
    }
    b.style.borderColor = color || '#b45309';
    b.innerHTML = '<div style="display:flex;justify-content:space-between;align-items:center"><b style="font-size:18px">' + esc(title) + '</b><button type="button" id="lslcx" style="font-size:22px;border:0;background:none;cursor:pointer" title="Close">&times;</button></div>' + html;
    $id('lslcx').onclick = function () { if (state.running) { alert('Still working. Please wait until it finishes.'); return; } b.remove(); };
    return b;
  }
  function frame(url) {
    return new Promise(function (res, rej) {
      var f = document.createElement('iframe');
      f.style.cssText = 'position:fixed;left:-4000px;top:0;width:1440px;height:1000px;border:0';
      var t = setTimeout(function () { rej(new Error('LightningStep didn\'t load ' + url)); }, 60000);
      f.onload = function () { clearTimeout(t); res(f); };
      f.src = url; document.body.appendChild(f);
    });
  }
  async function idle(w, ms) { var end = Date.now() + (ms || 15000); await sleep(400); while (Date.now() < end && w.jQuery && w.jQuery.active > 0) await sleep(250); await sleep(300); }
  var vis = function (w, el) { return el && w.getComputedStyle(el).display !== 'none' && w.getComputedStyle(el).visibility !== 'hidden' && (el.offsetWidth || el.offsetHeight || el.getClientRects().length); };
  var href = function (w) { try { return w.location.pathname + w.location.search; } catch (e) { return ''; } };

  // The LightningStep review and approval steps after Save on an admission
  async function approveFlow(w, reason, log) {
    var $ = w.jQuery, d = w.document;
    for (var k = 0; k < 60; k++) {
      await sleep(500);
      if (/\/inquiry\/show\/\d+/.test(href(w))) return 'saved';
      var ap = [].find.call(d.querySelectorAll('button,a'), function (b) { return /approve change/i.test(b.textContent) && vis(w, b); });
      if (ap) { log('Approving the change'); ap.click(); break; }
      if (vis(w, d.getElementById('approvalReason'))) break;
      var err = [].find.call(d.querySelectorAll('.alert-danger,.help-block'), function (e) { return vis(w, e) && strip(e.textContent).length > 3; });
      if (err && k > 6) throw new Error('LightningStep said: ' + strip(err.textContent).slice(0, 200));
    }
    if (/\/inquiry\/show\/\d+/.test(href(w))) return 'saved';
    // approval modal
    for (var k2 = 0; k2 < 30 && !vis(w, d.getElementById('approvalReason')); k2++) await sleep(300);
    await sleep(800);
    var yes = d.getElementById('approvalYes');
    if (!yes) throw new Error('The approval box didn\'t open.');
    if (!yes.checked) { var lab = d.querySelector('label[for="approvalYes"]'); if (lab) lab.click(); if (!yes.checked) { yes.checked = true; $(yes).trigger('change'); } }
    if (!yes.checked) throw new Error('Couldn\'t tick "update the affected items".');
    $('#approvalReason').val(reason).trigger('input').trigger('change');
    d.getElementById('approvalButton').click();
    // SweetAlert: type CONFIRM
    for (var k3 = 0; k3 < 30; k3++) {
      await sleep(400);
      var inp = d.querySelector('.swal2-input');
      if (inp && vis(w, inp)) { inp.value = 'CONFIRM'; inp.dispatchEvent(new w.Event('input', { bubbles: true })); await sleep(300); d.querySelector('.swal2-confirm').click(); break; }
      if (/\/inquiry\/show\/\d+/.test(href(w))) return 'saved';
    }
    for (var k4 = 0; k4 < 60; k4++) { await sleep(500); if (/\/inquiry\/show\/\d+/.test(href(w))) return 'saved'; }
    throw new Error('The change didn\'t finish saving. Check the Admissions tab before trying again.');
  }

  // Who is this chart?
  async function patient() {
    var cid = (location.pathname.match(/\/client\/index\/(\d+)/) || [])[1];
    var inq = (location.pathname.match(/\/inquiry\/show\/(\d+)/) || [])[1];
    if (!cid && inq) { var ih = await (await fetch('/inquiry/show/' + inq, { credentials: 'same-origin' })).text(); cid = (ih.match(/\/client\/index\/(\d+)/) || [])[1]; }
    if (!cid) throw new Error('Open the patient\'s chart first.');
    var cen = await post('/home/census', { draw: 1, start: 0, length: 2000, 'programfilter[]': 0 });
    var row = (cen.data || []).find(function (r) { return String(r.client_id) === String(cid) && /Cross Hope/i.test(strip(r.location)); });
    if (!row) throw new Error('This patient isn\'t on today\'s CrossHope census, so there\'s no active admission to change.');
    inq = inq || String(row.inquiry_id);
    var adm = await post('/admission/getadmissionsdatatable', { inquiry_id: inq, client_id: cid, profile: 'client', draw: 1 });
    var cur = (adm.data || []).find(function (a) { return /admitted/i.test(strip(a.status)); });
    if (!cur) throw new Error('Couldn\'t find an open admission for this patient.');
    var nm = String(row.lnamefname || '').split('|');
    return { cid: cid, inq: inq, adm: String(cur.id), program: strip(row.pgmname), payer: strip(row.payor), name: (nm[1] || '') + ' ' + (nm[0] || ''), short: (nm[1] || '').replace(/^./, function (c) { return c.toUpperCase(); }) + ' ' + (nm[0] || '').slice(0, 1).toUpperCase() + '.', since: row.datepgmadmission || '', admDate: String(cur.datetimeadmission || '') };
  }

  // Charges already built for this patient on or after a date
  async function chargesFrom(inq, fromIso) {
    try {
      var j = await getJ('/account/insurancecharges/datatables/' + inq + '?draw=1&start=0&length=500');
      return (j.data || []).filter(function (c) { var p = String(c.servicedate).split('/'); var iso = p[2] + '-' + p[0] + '-' + p[1]; return iso >= fromIso; })
        .map(function (c) { return c.servicedate + ' ' + strip(c.servicename) + ' ' + strip(c.billingcode) + ' $' + c.amount + (c.claim_id ? ' (claim ' + c.claim_id + ')' : ' (no claim yet)'); });
    } catch (e) { return ['Couldn\'t read charges: check the Financial tab.']; }
  }

  // ---------------- Discharge ----------------
  var DC = {
    done: { label: 'Completed the program', type: '1', code: '01' },
    left: { label: 'Left on their own / against advice', type: '3', code: '07' },
    staff: { label: 'Staff discharged them', type: '2', code: '01' },
    hosp: { label: 'Went to a medical hospital', type: '5', code: '02' },
    psych: { label: 'Went to a psychiatric hospital', type: '5', code: '65' }
  };
  async function openDischarge() {
    var P;
    box('Discharge', '<p>Loading...</p>');
    try { P = await patient(); } catch (e) { box('Discharge', '<p style="color:#b91c1c">' + esc(e.message) + '</p>'); return; }
    var now = CT();
    box('Discharge ' + P.short, '<p style="margin:6px 0">' + esc(P.name) + ' is in <b>' + esc(P.program) + '</b>' + (P.payer ? ' (' + esc(P.payer) + ')' : '') + '.</p>' +
      '<p style="margin:10px 0 4px"><b>Why are they leaving?</b></p>' +
      Object.keys(DC).map(function (k) { return '<label style="display:block;margin:4px 0"><input type="radio" name="lcwhy" value="' + k + '"> ' + esc(DC[k].label) + '</label>'; }).join('') +
      '<p style="margin:10px 0 4px"><b>When did they leave?</b> <input id="lc_d" type="date" value="' + now.date + '"> <input id="lc_t" type="time" value="' + now.time + '"></p>' +
      '<div style="font-size:13px;color:#555">"Last night" means 8:00 PM, "this morning" 9:00 AM. Use the last day they were here.</div>' +
      '<p style="margin:10px 0 4px"><b>What happened</b> (one line, it goes in the discharge reason)</p><input id="lc_r" style="width:100%;box-sizing:border-box;padding:6px" placeholder="e.g. Left after group, told staff she was going home">' +
      '<div id="lclog" style="margin:10px 0;font-size:14px"></div><button type="button" id="lc_go" style="' + btnCss + '">Check and discharge</button>', '#b91c1c');
    $id('lc_go').onclick = async function () {
      var why = (document.querySelector('input[name=lcwhy]:checked') || {}).value;
      var dd = $id('lc_d').value, tt = $id('lc_t').value, rr = $id('lc_r').value.trim();
      var lg = $id('lclog');
      if (!why || !dd || !tt || rr.length < 5) { lg.innerHTML = '<span style="color:#b91c1c">Pick a reason, the date and time, and write what happened.</span>'; return; }
      if (dd > CT().date) { lg.innerHTML = '<span style="color:#b91c1c">The date can\'t be in the future.</span>'; return; }
      // look for attendance after the discharge date
      lg.textContent = 'Checking attendance after ' + mdy(dd) + '...';
      var later = [];
      var dt = new Date(dd + 'T12:00:00');
      for (var i = 1; i <= 7; i++) {
        var x = new Date(dt.getTime() + i * 864e5); var iso = x.toISOString().slice(0, 10); if (iso > CT().date) break;
        try { var a = await post('/attendance/attendeddatatable', { location_id: 21, datechosen: iso, service_id: 0, draw: 1, start: 0, length: 1000 }); if ((a.data || []).some(function (r) { return strip(r.name).toLowerCase().indexOf(P.name.split(' ').slice(-1)[0].toLowerCase()) > -1 && strip(r.name).toLowerCase().indexOf(P.name.split(' ')[0].toLowerCase()) > -1; })) later.push(mdy(iso)); } catch (e) {}
      }
      if (later.length) { lg.innerHTML = '<span style="color:#b91c1c">Stopped. ' + esc(P.short) + ' has attendance marked after the discharge date (' + later.join(', ') + '). Check the date, or ask Mendy.</span>'; return; }
      if (!confirm('Discharge ' + P.name + ' on ' + mdy(dd) + ' at ' + tt + ' (' + DC[why].label + ')? This can\'t be undone here.')) return;
      state.running = true; $id('lc_go').disabled = true;
      var fr;
      try {
        lg.textContent = 'Opening the admission...';
        fr = await frame('/admission/' + P.adm + '/edit');
        var w = fr.contentWindow, $ = w.jQuery, d = w.document; await idle(w);
        var radio = d.querySelector('input[name=status][value=Discharged]');
        if (!radio) throw new Error('No Discharged option on the admission page.');
        radio.click(); $(radio).trigger('change'); await sleep(600);
        var S = function (n, v) { var el = d.querySelector('[name="' + n + '"]'); if (!el) throw new Error('missing ' + n); $(el).val(v).trigger('input').trigger('change'); };
        S('datedischarge', dd); S('timedischarge', tt); S('dctype_id', DC[why].type); await sleep(300); S('dcstatuscode', DC[why].code); S('dcreason', rr + ' (entered with LS Forms)');
        var fu = d.getElementById('isfollowup'); if (fu && fu.checked) { var fl = d.querySelector('label[for="isfollowup"]'); if (fl) fl.click(); if (fu.checked) { fu.checked = false; $(fu).trigger('change'); } }
        await sleep(400);
        var chk = { date: d.querySelector('[name=datedischarge]').value, time: d.querySelector('[name=timedischarge]').value, type: d.querySelector('[name=dctype_id]').value, code: d.querySelector('[name=dcstatuscode]').value, follow: fu ? fu.checked : false };
        if (chk.date !== dd || chk.time.slice(0, 5) !== tt || chk.type !== DC[why].type || chk.code !== DC[why].code || chk.follow) throw new Error('The form didn\'t take the values (' + JSON.stringify(chk) + '). Nothing was saved.');
        lg.textContent = 'Saving the discharge...';
        d.getElementById('btnsubmit').click();
        await approveFlow(w, 'Discharge entered with LS Forms: ' + rr, function (m) { lg.textContent = m + '...'; });
        fr.remove();
        var adm = await post('/admission/getadmissionsdatatable', { inquiry_id: P.inq, client_id: P.cid, profile: 'client', draw: 1 });
        var row = (adm.data || []).find(function (a) { return String(a.id) === P.adm; });
        var ok = row && /discharg/i.test(strip(row.status));
        state.running = false;
        lg.innerHTML = ok ? '<div style="background:#f0fdf4;border:2px solid #166534;border-radius:8px;padding:10px"><b>Done.</b> ' + esc(P.short) + ' is discharged as of ' + mdy(dd) + ' at ' + esc(tt) + ' (' + esc(DC[why].label) + ').<br>Unity Living: tell Jason the discharge date if they live there.</div>'
          : '<span style="color:#b91c1c">Saved, but the Admissions tab doesn\'t show Discharged yet. Check it before doing anything else.</span>';
      } catch (e) {
        state.running = false; try { if (fr) fr.remove(); } catch (x) {}
        lg.innerHTML = '<div style="background:#fef2f2;border:2px solid #b91c1c;border-radius:8px;padding:10px"><b>Stopped.</b> ' + esc(e.message) + '<br>Check the Admissions tab before trying again.</div>';
      }
    };
  }

  // ---------------- Level of care ----------------
  async function openLoc() {
    var P;
    box('Change level of care', '<p>Loading...</p>');
    try { P = await patient(); } catch (e) { box('Change level of care', '<p style="color:#b91c1c">' + esc(e.message) + '</p>'); return; }
    var curSub = { 'IOP': '2', 'PHP': '7', 'IOP. 1': '10', 'PHP.': '13' }[P.program.trim()];
    if (!curSub) { box('Change level of care', '<p>' + esc(P.short) + ' is in ' + esc(P.program) + '. This button only moves patients between IOP and PHP. Ask Mendy.</p>'); return; }
    var target = curSub === '2' ? '7' : curSub === '7' ? '2' : curSub === '10' ? '13' : '10';
    var medicaid = /medicaid/i.test(P.payer) && !/passe|caresource|total care|ambetter|blue/i.test(P.payer);
    var ruleWarn = (PROG[target].name.indexOf('PHP') === 0 && medicaid) ? 'CrossHope rule: Arkansas Medicaid patients go to IOP. Moving to PHP needs Mendy\'s OK.' : (PROG[target].name.indexOf('IOP') === 0 && P.payer && !medicaid) ? 'CrossHope rule: payers other than Arkansas Medicaid go to PHP. Moving to IOP needs Mendy\'s OK.' : '';
    box('Change level of care: ' + P.short, '<p style="margin:6px 0">' + esc(P.name) + ' is in <b>' + esc(P.program) + '</b>' + (P.payer ? ' with ' + esc(P.payer) : '') + '. Move to <b>' + esc(PROG[target].name) + '</b>.</p>' +
      (ruleWarn ? '<div style="background:#fffbeb;border:1px solid #d97706;border-radius:8px;padding:8px;margin:8px 0">' + esc(ruleWarn) + '</div>' : '') +
      '<p style="margin:10px 0 4px"><b>Which is it?</b></p>' +
      '<label style="display:block;margin:6px 0"><input type="radio" name="lckind" value="step"> <b>Step ' + (/PHP/.test(PROG[target].name) ? 'up' : 'down') + '</b>: their care is changing now. Past days stay as they are; ' + esc(PROG[target].name) + ' starts tomorrow.</label>' +
      '<label style="display:block;margin:6px 0"><input type="radio" name="lckind" value="fix"> <b>Entered in the wrong program</b>: it was a mistake from the first day. The whole stay moves to ' + esc(PROG[target].name) + ', and past attendance and group notes are fixed too.</label>' +
      '<p style="margin:10px 0 4px"><b>Why</b> (one line, it goes in the record)</p><input id="lc_r" style="width:100%;box-sizing:border-box;padding:6px" placeholder="e.g. Stepped down to IOP per treatment team 11/23">' +
      '<div id="lclog" style="margin:10px 0;font-size:14px"></div><button type="button" id="lc_go" style="' + btnCss + '">Check and change</button>', '#b45309');
    $id('lc_go').onclick = async function () {
      var kind = (document.querySelector('input[name=lckind]:checked') || {}).value, rr = $id('lc_r').value.trim(), lg = $id('lclog');
      if (!kind || rr.length < 5) { lg.innerHTML = '<span style="color:#b91c1c">Pick which it is and write why.</span>'; return; }
      if (ruleWarn && !confirm(ruleWarn + '\n\nDid Mendy OK this?')) return;
      if (kind === 'step') return stepChange(P, target, rr, lg);
      return wrongFix(P, curSub, target, rr, lg);
    };
  }

  async function stepChange(P, target, rr, lg) {
    var today = CT().date, tomorrow = CT(1).date;
    if (!confirm('Move ' + P.name + ' to ' + PROG[target].name + ' starting ' + mdy(tomorrow) + '? Today stays in ' + P.program + '.')) return;
    state.running = true; $id('lc_go').disabled = true;
    var fr;
    try {
      lg.textContent = 'Opening the admission...';
      fr = await frame('/admission/' + P.adm + '/edit');
      var w = fr.contentWindow, $ = w.jQuery, d = w.document; await idle(w);
      var radio = d.querySelector('input[name=status][value=Transfer]');
      radio.click(); $(radio).trigger('change'); await sleep(600);
      var S = function (n, v) { var el = d.querySelector('[name="' + n + '"]'); if (!el) throw new Error('missing ' + n); $(el).val(v).trigger('input').trigger('change'); };
      S('datetransfer', today); S('timetransfer', '09:00'); S('dateadmit_transfer', tomorrow);
      S('transferlocation_id', '21');
      for (var k = 0; k < 30 && !d.querySelector('[name=transferto_id] option[value="' + target + '"]'); k++) await sleep(300);
      if (!d.querySelector('[name=transferto_id] option[value="' + target + '"]')) throw new Error(PROG[target].name + ' isn\'t offered as a transfer program. Nothing was changed.');
      S('transferto_id', target); S('jfc', rr + ' (entered with LS Forms)');
      var ne = d.querySelector('[name=isnewepisode]'); if (ne && ne.checked) { ne.checked = false; $(ne).trigger('change'); }
      await sleep(400);
      var chk = { dt: d.querySelector('[name=datetransfer]').value, na: d.querySelector('[name=dateadmit_transfer]').value, to: d.querySelector('[name=transferto_id]').value };
      if (chk.dt !== today || chk.na !== tomorrow || chk.to !== target) throw new Error('The form didn\'t take the values. Nothing was saved.');
      lg.textContent = 'Saving the transfer...';
      d.getElementById('btnsubmit').click();
      await approveFlow(w, 'Level of care change entered with LS Forms: ' + rr, function (m) { lg.textContent = m + '...'; });
      fr.remove();
      var adm = await post('/admission/getadmissionsdatatable', { inquiry_id: P.inq, client_id: P.cid, profile: 'client', draw: 1 });
      var newRow = (adm.data || []).find(function (a) { return /admitted/i.test(strip(a.status)); });
      var ch = await chargesFrom(P.inq, tomorrow);
      state.running = false;
      lg.innerHTML = '<div style="background:#f0fdf4;border:2px solid #166534;border-radius:8px;padding:10px"><b>Done.</b> ' + esc(P.short) + ' moves to ' + esc(PROG[target].name) + ' on ' + mdy(tomorrow) + (newRow ? ' (new admission ' + newRow.id + ', ' + esc(strip(newRow.subprogram)) + ')' : '') + '.' +
        '<br>LightningStep adds them to the new program\'s group on the calendar by itself. <b>Tell Mendy</b> to check they\'re off the old program\'s group from tomorrow.' +
        (ch.length ? '<br><b>Charges already built from ' + mdy(tomorrow) + ':</b><br>' + ch.map(esc).join('<br>') : '') + '</div>' + billingNote(P, 'stepped to ' + PROG[target].name + ' from ' + mdy(tomorrow), rr, ch);
      wireCopy();
    } catch (e) {
      state.running = false; try { if (fr) fr.remove(); } catch (x) {}
      lg.innerHTML = '<div style="background:#fef2f2;border:2px solid #b91c1c;border-radius:8px;padding:10px"><b>Stopped.</b> ' + esc(e.message) + '<br>Check the Admissions tab before trying again.</div>';
    }
  }

  function billingNote(P, what, rr, ch) {
    var txt = 'Hi Malky, ' + P.short + ' ' + what + '. Reason: ' + rr + '.' + (ch.length ? ' Charges already built for those dates: ' + ch.join('; ') + '.' : ' No charges were built for those dates yet.') + ' Entered by LS Forms.';
    return '<div style="margin-top:10px"><b>Message for billing</b> (paste it to Malky in Teams):<textarea id="lcbill" rows="4" style="width:100%;box-sizing:border-box;margin-top:4px">' + esc(txt) + '</textarea><button type="button" id="lccopy" style="' + btnCss + ';background:#475569;margin-top:4px">Copy message</button></div>';
  }
  function wireCopy() { var b = $id('lccopy'); if (b) b.onclick = function () { var t = $id('lcbill'); t.select(); try { navigator.clipboard.writeText(t.value); } catch (e) { document.execCommand('copy'); } b.textContent = 'Copied'; }; }

  async function wrongFix(P, curSub, target, rr, lg) {
    var admIso = (P.admDate.match(/(\d{2})\/(\d{2})\/(\d{4})/) || []);
    var startIso = admIso.length ? admIso[3] + '-' + admIso[1] + '-' + admIso[2] : '';
    if (!confirm('Change ' + P.name + '\'s whole stay' + (startIso ? ' (from ' + mdy(startIso) + ')' : '') + ' from ' + P.program + ' to ' + PROG[target].name + '? Past attendance gets relabeled and group notes will need redoing.')) return;
    state.running = true; $id('lc_go').disabled = true;
    var fr;
    try {
      // 1. Power Edit: send the request
      lg.textContent = 'Opening Power Edit...';
      fr = await frame('/admission/' + P.adm + '/poweredit');
      var w = fr.contentWindow, $ = w.jQuery, d = w.document; await idle(w);
      var sp = d.querySelector('[name=subprogram_id]');
      if (!sp || !sp.querySelector('option[value="' + target + '"]')) throw new Error(PROG[target].name + ' isn\'t offered in Power Edit. Nothing was changed.');
      $(sp).val(target).trigger('change'); await sleep(500);
      var saveBtn = d.getElementById('admissionDetailsBtn');
      if (!saveBtn) throw new Error('No Save button on Power Edit.');
      saveBtn.click();
      // reason prompt then Send Request (or it applies at once)
      var sent = false;
      for (var k = 0; k < 40 && !sent; k++) {
        await sleep(500);
        var reason = [].find.call(d.querySelectorAll('textarea,input[type=text]'), function (x) { return vis(w, x) && ['approvalReason', 'dcreason', 'jfc'].indexOf(x.id) < 0 && x.name !== 'jfc' && x.name !== 'dcreason' && /reason for change/i.test(strip((x.closest('div') || {}).textContent || '') + ' ' + strip(((x.closest('form') || d.body).textContent || '').slice(0, 0))); }) || [].find.call(d.querySelectorAll('textarea'), function (x) { return vis(w, x) && ['approvalReason', 'dcreason', 'jfc'].indexOf(x.id) < 0 && x.name !== 'jfc'; });
        if (reason && !reason.value) { reason.value = rr + ' (entered with LS Forms)'; reason.dispatchEvent(new w.Event('input', { bubbles: true })); $(reason).trigger('change'); }
        var send = [].find.call(d.querySelectorAll('button'), function (b) { return /send request/i.test(b.textContent) && vis(w, b); });
        if (send && reason && reason.value) { send.click(); sent = true; for (var k9 = 0; k9 < 10; k9++) { await sleep(400); var ok9 = d.querySelector('.swal2-confirm'); if (ok9 && vis(w, ok9)) { ok9.click(); break; } } }
        if (/\/inquiry\/show|\/client\/index/.test(href(w))) sent = true;
      }
      if (!sent) throw new Error('The Power Edit request didn\'t send. Nothing was changed.');
      await sleep(2500); fr.remove();
      // 2. reload and approve (Connie approves her own if LightningStep lets her)
      lg.textContent = 'Approving the change...';
      fr = await frame('/admission/' + P.adm + '/poweredit');
      w = fr.contentWindow; $ = w.jQuery; d = w.document; await idle(w);
      var apb = [].find.call(d.querySelectorAll('button,a'), function (b) { return /^approve$/i.test(strip(b.textContent)) && vis(w, b); });
      var applied = false;
      if (apb) {
        apb.click();
        await approveFlow(w, 'Wrong program from the first day, fixed with LS Forms: ' + rr, function (m) { lg.textContent = m + '...'; }).catch(function (e) { if (!/didn.t finish/.test(e.message)) throw e; });
        applied = true;
      }
      fr.remove();
      var adm = await post('/admission/getadmissionsdatatable', { inquiry_id: P.inq, client_id: P.cid, profile: 'client', draw: 1 });
      var row = (adm.data || []).find(function (a) { return String(a.id) === P.adm; });
      var nowProg = row ? strip(row.subprogram) : '';
      var landed = new RegExp(PROG[target].name.split(' ')[0].replace('.', '\\.'), 'i').test(nowProg) && !new RegExp('^' + P.program.replace(/[.()]/g, '\\$&') + '$', 'i').test(nowProg);
      if (!landed) {
        state.running = false;
        lg.innerHTML = '<div style="background:#fffbeb;border:2px solid #d97706;border-radius:8px;padding:10px"><b>Request sent, waiting for approval.</b> LightningStep didn\'t let this login approve its own Power Edit' + (applied ? ' (or it hasn\'t applied yet)' : '') + '. Ask Mendy or Abigail to approve it. Then open this chart and click <b>Change level of care</b> again: it will pick up at fixing attendance and notes.</div>';
        sessionStorage.setItem('lslc-pending-' + P.cid, JSON.stringify({ target: target, from: curSub, rr: rr }));
        return;
      }
      await fixPast(P, curSub, target, rr, lg, startIso);
    } catch (e) {
      state.running = false; try { if (fr) fr.remove(); } catch (x) {}
      lg.innerHTML = '<div style="background:#fef2f2;border:2px solid #b91c1c;border-radius:8px;padding:10px"><b>Stopped.</b> ' + esc(e.message) + '<br>Check the Admissions tab before trying again.</div>';
    }
  }

  // After the program is fixed: relabel attendance, list group notes to redo, billing note
  async function fixPast(P, fromSub, target, rr, lg, startIso) {
    state.running = true;
    var today = CT().date;
    var start = startIso || today;
    var last = P.name.split(' ').slice(-1)[0].toLowerCase(), first = P.name.split(' ')[0].toLowerCase();
    var mine = function (r) { var n = strip(r.name).toLowerCase(); return n.indexOf(last) > -1 && n.indexOf(first) > -1; };
    lg.textContent = 'Relabeling attendance...';
    var relabeled = [], skipped = [];
    for (var t = new Date(start + 'T12:00:00'); t.toISOString().slice(0, 10) <= today; t = new Date(t.getTime() + 864e5)) {
      var iso = t.toISOString().slice(0, 10);
      var att;
      try { att = await post('/attendance/attendeddatatable', { location_id: 21, datechosen: iso, service_id: 0, draw: 1, start: 0, length: 1000 }); } catch (e) { continue; }
      var r = (att.data || []).find(mine);
      if (!r) continue;
      if (new RegExp(PROG[target].name.split(' ')[0], 'i').test(strip(r.programname))) continue;
      var cid = r.census_id || r.id;
      if (!cid) { skipped.push(mdy(iso)); continue; }
      var t1 = tok();
      var u = await fetch('/attendance/unmarkattended/' + cid, { method: 'DELETE', credentials: 'same-origin', headers: Object.assign({ 'X-CSRF-TOKEN': t1 }, H) });
      if (!u.ok) { skipped.push(mdy(iso)); continue; }
      await post('/attendance/markattended', { census_id: cid, modifierkey: '', service_id: PROG[target].svc });
      relabeled.push(mdy(iso));
    }
    // group sessions the patient is in under the old program
    lg.textContent = 'Looking for group notes to redo...';
    var redo = [];
    try {
      var gs = await post('/groupsession/listsupervisor', { datestart: start, dateend: today, provider_id: 0, location_id: 21, subprogram_id: 0 });
      for (var g of (gs.groupsessions || [])) {
        var h = await (await fetch('/groupsession/' + g.id + '/edit', { credentials: 'same-origin', headers: { 'X-Requested-With': 'XMLHttpRequest' } })).text();
        var re = /var at = "(\d+)";\s*var cid = "(\d+)";/g, mm, present = false;
        while ((mm = re.exec(h))) if (mm[2] === String(P.cid) && mm[1] === '1') present = true;
        if (!present) continue;
        var isTarget = new RegExp(PROG[target].name.split(' ')[0], 'i').test(g.groupname);
        if (isTarget) continue;
        var note = (h.match(new RegExp('btngoedit' + P.cid + '[^>]*data-note_id="(\\d+)"')) || h.match(new RegExp('data-note_id="(\\d+)"[^>]*btngoedit' + P.cid)) || [])[1];
        redo.push({ date: g.datedoc, session: g.id, group: g.groupname, note: note });
      }
    } catch (e) { redo.push({ error: true }); }
    var ch = await chargesFrom(P.inq, start);
    state.running = false;
    sessionStorage.removeItem('lslc-pending-' + P.cid);
    sessionStorage.setItem('lslc-redo-' + P.cid, JSON.stringify({ target: target, rr: rr, redo: redo }));
    lg.innerHTML = '<div style="background:#f0fdf4;border:2px solid #166534;border-radius:8px;padding:10px"><b>Program fixed:</b> ' + esc(P.short) + ' is now in ' + esc(PROG[target].name) + ' for the whole stay.' +
      '<br><b>Attendance relabeled:</b> ' + (relabeled.length ? relabeled.join(', ') : 'none needed') + (skipped.length ? '<br><span style="color:#b91c1c">Couldn\'t relabel: ' + skipped.join(', ') + '. Ask Mendy.</span>' : '') +
      '<br><b>Calendar:</b> Power Edit doesn\'t move them between the program groups. <b>Tell Mendy</b> to move them to the ' + esc(PROG[target].name) + ' group from the next group day.</div>' +
      (redo.length ? '<div style="margin-top:10px"><b>Group notes to redo (' + redo.filter(function (x) { return !x.error; }).length + '):</b> these were written in the old program\'s group. Do them one at a time:' +
        '<ol style="margin:6px 0 0 18px;padding:0">' + redo.map(function (x, i) { return x.error ? '<li style="color:#b91c1c">Couldn\'t read the group sessions. Ask Mendy.</li>' : '<li style="margin:6px 0">' + esc(x.date) + ' ' + esc(x.group) + ' <button type="button" class="lcredo" data-i="' + i + '" style="' + btnCss + ';padding:4px 10px;font-size:13px">Redo this note</button></li>'; }).join('') + '</ol></div>' : '<p>No group notes to redo.</p>') +
      billingNote(P, 'was entered in the wrong program and is now ' + PROG[target].name + ' for the whole stay from ' + mdy(start), rr, ch);
    wireCopy();
    [].forEach.call(document.querySelectorAll('.lcredo'), function (b) { b.onclick = function () { redoNote(P, target, redo[+b.getAttribute('data-i')], b); }; });
  }

  // One note: copy old text, put her in the right session, mark Present, open the new note
  async function redoNote(P, target, x, b) {
    b.disabled = true; b.textContent = 'Working...';
    try {
      var oldText = '';
      if (x.note) { var nh = await (await fetch('/groupnote/' + x.note + '/edit', { credentials: 'same-origin' })).text(); var nd = new DOMParser().parseFromString(nh, 'text/html'); var ta = nd.querySelector('textarea[name=data]'); oldText = ta ? ta.value || ta.textContent : ''; }
      var gs = await post('/groupsession/listsupervisor', { datestart: x.date.replace(/(\d{2})\/(\d{2})\/(\d{4})/, '$3-$1-$2'), dateend: x.date.replace(/(\d{2})\/(\d{2})\/(\d{4})/, '$3-$1-$2'), provider_id: 0, location_id: 21, subprogram_id: 0 });
      var right = (gs.groupsessions || []).find(function (g) { return new RegExp(PROG[target].name.split(' ')[0], 'i').test(g.groupname); });
      if (!right) throw new Error('No ' + PROG[target].name + ' group session on ' + x.date + '. Ask Mendy.');
      var eh = await (await fetch('/groupsession/enrollment/' + right.id, { credentials: 'same-origin' })).text();
      if (eh.indexOf('rowenr' + P.cid) < 0 && !new RegExp('"client_id"\\s*:\\s*"?' + P.cid).test(eh)) {
        var defid = (eh.match(/name="groupdef_id"[^>]*value="(\d+)"/) || eh.match(/groupdef_id["']?\s*[:=]\s*["']?(\d+)/) || [])[1] || '';
        await post('/groupsession/enrollclient', { groupsession_id: right.id, client_id: P.cid, groupdef_id: defid, groupname: right.groupname });
      }
      var sh = await (await fetch('/groupsession/' + right.id + '/edit', { credentials: 'same-origin', headers: { 'X-Requested-With': 'XMLHttpRequest' } })).text();
      var re = /var at = "(\d+)";\s*var cid = "(\d+)";/g, mm, at = '0';
      while ((mm = re.exec(sh))) if (mm[2] === String(P.cid)) at = mm[1];
      if (at !== '1') {
        var docId = (sh.match(/name="doc_id"[^>]*value="(\d+)"/) || [])[1] || '';
        await post('/groupsession/markattendii', { attendancetype_id: 1, groupsession_id: right.id, secprovider_id: 0, client_id: P.cid, doc_id: docId });
        await sleep(1200);
        sh = await (await fetch('/groupsession/' + right.id + '/edit', { credentials: 'same-origin', headers: { 'X-Requested-With': 'XMLHttpRequest' } })).text();
      }
      var newNote = (sh.match(new RegExp('btngoedit' + P.cid + '[^>]*data-note_id="(\\d+)"')) || sh.match(new RegExp('data-note_id="(\\d+)"[^>]*btngoedit' + P.cid)) || [])[1];
      if (!newNote) throw new Error('Couldn\'t find the new note in the ' + PROG[target].name + ' session. Open the session from the calendar.');
      sessionStorage.setItem('lslc-note', JSON.stringify({ newNote: newNote, oldNote: x.note, text: oldText, date: x.date, who: P.short }));
      b.textContent = 'Opening...';
      location.href = '/groupnote/' + newNote + '/edit';
    } catch (e) { b.disabled = false; b.textContent = 'Redo this note'; alert(e.message); }
  }

  // Runs on every page: helper banners for the note redo
  function notePages() {
    var s; try { s = JSON.parse(sessionStorage.getItem('lslc-note') || 'null'); } catch (e) {}
    if (!s) return;
    var m = location.pathname.match(/\/groupnote\/(\d+)\/edit/);
    if (!m) return;
    var bar = document.createElement('div');
    bar.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:2147483645;padding:10px 16px;font:600 15px system-ui,"Segoe UI",sans-serif;color:#fff;display:flex;gap:10px;align-items:center;flex-wrap:wrap';
    if (m[1] === String(s.newNote)) {
      bar.style.background = '#166534';
      bar.innerHTML = '<span>NEW note for ' + esc(s.who) + ', ' + esc(s.date) + '. 1. Click Copy old text, then click into the note box and press Ctrl+V. 2. Check it, tick Sign and Lock, and Save.</span>';
      var cp = document.createElement('button'); cp.textContent = 'Copy old text'; cp.style.cssText = 'padding:6px 12px;border:0;border-radius:6px;font-weight:700;cursor:pointer';
      cp.onclick = function () { var t = document.createElement('textarea'); t.value = s.text || ''; document.body.appendChild(t); t.select(); try { navigator.clipboard.writeText(s.text || ''); } catch (e) {} document.execCommand('copy'); t.remove(); cp.textContent = 'Copied'; };
      bar.appendChild(cp);
      if (s.oldNote) { var go = document.createElement('a'); go.href = '/groupnote/' + s.oldNote + '/edit'; go.textContent = 'After saving: open the OLD note'; go.style.cssText = 'color:#fff;text-decoration:underline;margin-left:8px'; bar.appendChild(go); }
    } else if (s.oldNote && m[1] === String(s.oldNote)) {
      bar.style.background = '#b91c1c';
      bar.innerHTML = '<span>OLD note (wrong program) for ' + esc(s.who) + ', ' + esc(s.date) + '. Only do this after the new note is signed: scroll to the bottom, tick Entered in Error, type "Moved to the right program\'s group note", and click Mark as Error.</span>';
      var dn = document.createElement('button'); dn.textContent = 'Done with this note'; dn.style.cssText = 'padding:6px 12px;border:0;border-radius:6px;font-weight:700;cursor:pointer';
      dn.onclick = function () { sessionStorage.removeItem('lslc-note'); bar.remove(); history.back(); };
      bar.appendChild(dn);
    } else return;
    document.body.appendChild(bar);
  }

  async function openLocOrResume() {
    var cid = (location.pathname.match(/\/client\/index\/(\d+)/) || [])[1];
    var pend = cid && sessionStorage.getItem('lslc-pending-' + cid);
    if (pend) {
      var P; try { P = await patient(); } catch (e) { return openLoc(); }
      var pp = JSON.parse(pend);
      if (new RegExp(PROG[pp.target].name.split(' ')[0], 'i').test(P.program)) {
        box('Change level of care: ' + P.short, '<p>The Power Edit was approved. Finishing attendance and notes...</p><div id="lclog"></div><button id="lc_go" style="display:none"></button>');
        return fixPast(P, pp.from, pp.target, pp.rr, $id('lclog'), '');
      }
    }
    var redo = cid && sessionStorage.getItem('lslc-redo-' + cid);
    if (redo) {
      var R = JSON.parse(redo), P2;
      try { P2 = await patient(); } catch (e) { return openLoc(); }
      if (R.redo && R.redo.length) {
        box('Group notes to redo: ' + P2.short, '<ol style="margin:6px 0 0 18px;padding:0">' + R.redo.map(function (x, i) { return x.error ? '' : '<li style="margin:6px 0">' + esc(x.date) + ' ' + esc(x.group) + ' <button type="button" class="lcredo" data-i="' + i + '" style="' + btnCss + ';padding:4px 10px;font-size:13px">Redo this note</button></li>'; }).join('') + '</ol><p style="font-size:13px;color:#555">Done with all of them? <a href="#" id="lcclear">Clear this list</a>. Want a different change? <a href="#" id="lcnew">Change level of care</a>.</p>', '#b45309');
        [].forEach.call(document.querySelectorAll('.lcredo'), function (b) { b.onclick = function () { redoNote(P2, R.target, R.redo[+b.getAttribute('data-i')], b); }; });
        $id('lcclear').onclick = function (e) { e.preventDefault(); sessionStorage.removeItem('lslc-redo-' + cid); $id('lslc').remove(); };
        $id('lcnew').onclick = function (e) { e.preventDefault(); openLoc(); };
        return;
      }
    }
    return openLoc();
  }

  return { openDischarge: openDischarge, openLoc: openLocOrResume, notePages: notePages };
})();

// ---------- Treatment plans: who needs a plan or a review ----------
var TP = (function () {
  'use strict';
  // House rules (change here). PHP: DHS Partial Hospitalization rule says review
  // every 20 days of service; counted here as 20 weekdays. IOP: every 30 days
  // (CrossHope's usual "Date Next Review" pattern).
  var RULE = { PHP: { workdays: 20, label: 'every 20 service days' }, IOP: { days: 30, label: 'every 30 days' } };
  var SOON = 7; // "Due this week" window, in days
  var CACHE = 'lstp-cache-v1', TTL = 60 * 60 * 1000;
  var strip = function (s) { return String(s == null ? '' : s).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(); };
  var tok = function () { return document.querySelector('meta[name="csrf-token"]').content; };
  async function post(u, body) {
    var t = tok();
    var r = await fetch(u, { method: 'POST', credentials: 'same-origin', headers: { 'X-CSRF-TOKEN': t, 'X-Requested-With': 'XMLHttpRequest', 'Content-Type': 'application/x-www-form-urlencoded', 'Accept': 'application/json' }, body: '_token=' + encodeURIComponent(t) + '&' + body });
    if (!r.ok) throw new Error('LightningStep answered ' + r.status);
    return r.json();
  }
  function day(s) { var m = String(s || '').match(/(\d{4})-(\d{2})-(\d{2})/) || []; if (m[1]) return new Date(+m[1], m[2] - 1, +m[3]); m = String(s || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? new Date(+m[3], m[1] - 1, +m[2]) : null; }
  function today() { var n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); }
  function addDays(d, n) { var x = new Date(d); x.setDate(x.getDate() + n); return x; }
  function addWorkdays(d, n) { var x = new Date(d); while (n > 0) { x.setDate(x.getDate() + 1); if (x.getDay() % 6) n--; } return x; }
  function diff(a, b) { return Math.round((a - b) / 86400000); }
  function fmt(d) { return d ? (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear() : ''; }
  function shortName(n) { var p = n.split(/,\s*/); return p.length > 1 ? (p[1] + ' ' + p[0].charAt(0) + '.') : n; }

  async function load(force) {
    if (!force) { try { var c = JSON.parse(localStorage.getItem(CACHE) || 'null'); if (c && Date.now() - c.t < TTL) return c.list; } catch (e) {} }
    var cen = await post('/home/census', 'draw=1&start=0&length=2000&programfilter%5B%5D=0');
    var pts = (cen.data || []).filter(function (r) {
      var p = strip(r.pgmname);
      return /cross/i.test(r.location || '') && /^(PHP|IOP)$/.test(p) && String(r.isdc) !== '1';
    }).map(function (r) {
      var a = String(r.name || '').match(/<a\b[^>]*>([^<]*)<\/a>/i);
      return { c: r.client_id, n: strip(a ? a[1] : r.lnamefname), p: strip(r.pgmname), adm: day(r.datepgmadmission_for_sort || r.dateagencyadmission_for_sort), ep: day(r.current_episode_for_sort) };
    }).filter(function (r) { return !/^demo\b|,\s*demo\b/i.test(r.n); });
    var out = [], i = 0;
    async function one(r) {
      var j = await post('/txplan/datatable', 'client_id=' + r.c + '&draw=1&start=0&length=100');
      var since = r.ep || r.adm;
      var plans = (j.data || []).map(function (x) { return { id: x.id, d: day(x.datedoc), tpr: String(x.istpr) === '1', st: strip(x.status) }; })
        .filter(function (x) { return x.d && (!since || x.d >= addDays(since, -3)); })
        .sort(function (a, b) { return b.d - a.d || b.id - a.id; });
      var last = plans[0] || null, t = today(), o = { c: r.c, n: r.n, p: r.p, adm: r.adm && r.adm.getTime() };
      if (!last) { o.g = 'none'; o.days = r.adm ? diff(t, r.adm) : null; }
      else {
        var due = r.p === 'PHP' ? addWorkdays(last.d, RULE.PHP.workdays) : addDays(last.d, RULE.IOP.days);
        o.last = last.d.getTime(); o.kind = last.tpr ? 'Review' : 'Plan'; o.unsigned = last.st === 'Active';
        o.due = due.getTime(); o.days = diff(t, last.d); o.left = diff(due, t);
        o.g = o.left < 0 ? 'over' : (o.left <= SOON ? 'soon' : 'ok');
      }
      out.push(o);
    }
    var queue = pts.slice();
    async function worker() { while (queue.length) { var r = queue.shift(); try { await one(r); } catch (e) { out.push({ c: r.c, n: r.n, p: r.p, g: 'err' }); } } }
    await Promise.all([worker(), worker(), worker(), worker()]);
    try { localStorage.setItem(CACHE, JSON.stringify({ t: Date.now(), list: out })); } catch (e) {}
    return out;
  }
  function count(list) { return list.filter(function (x) { return x.g === 'none' || x.g === 'over'; }).length; }

  async function badge(btn) {
    try { var l = await load(false); var n = count(l); btn.textContent = n ? 'Treatment plans (' + n + ')' : 'Treatment plans'; btn.style.background = n ? '#dc2626' : '#0e7490'; } catch (e) {}
  }

  async function startReview(o, msg) {
    msg.textContent = 'Opening ' + shortName(o.n) + '...';
    var h = await (await fetch('/client/index/' + o.c, { credentials: 'same-origin' })).text();
    var url = (h.match(/\/txplan\/create\/\d+\/1\?episodeId=\d+/) || [])[0];
    if (!url) { location.href = '/client/index/' + o.c + '?section=txplan'; return; }
    try { localStorage.removeItem(CACHE); } catch (e) {}
    location.href = url.replace(/&amp;/g, '&');
  }

  function open() {
    var old = document.getElementById('lstp'); if (old) { old.remove(); return; }
    var b = document.createElement('div'); b.id = 'lstp';
    b.style.cssText = 'position:fixed;right:18px;bottom:74px;z-index:2147483647;background:#fff;border:2px solid #0e7490;border-radius:12px;padding:16px;width:640px;max-width:calc(100vw - 40px);max-height:78vh;overflow:auto;font:15px system-ui,"Segoe UI",sans-serif;box-shadow:0 8px 30px rgba(0,0,0,.3)';
    b.innerHTML = '<div style="display:flex;justify-content:space-between;align-items:center"><b style="font-size:18px">Treatment plans</b><span><button type="button" id="lstpr" style="border:1px solid #ccc;background:#fff;border-radius:6px;padding:4px 10px;cursor:pointer;margin-right:6px">Refresh</button><button type="button" id="lstpx" style="font-size:22px;border:0;background:none;cursor:pointer" title="Close">&times;</button></span></div>' +
      '<div id="lstpm" style="margin:6px 0 10px;color:#555;font-size:13px">Checking every PHP and IOP patient on the census...</div><div id="lstpl"></div>';
    document.body.appendChild(b);
    b.querySelector('#lstpx').onclick = function () { b.remove(); };
    b.querySelector('#lstpr').onclick = function () { draw(true); };
    var msg = b.querySelector('#lstpm'), L = b.querySelector('#lstpl');
    async function draw(force) {
      L.innerHTML = ''; msg.textContent = 'Checking every PHP and IOP patient on the census...';
      var list;
      try { list = await load(force); } catch (e) { msg.textContent = 'Couldn\'t load the list (' + e.message + '). Make sure you\'re signed in, then click Refresh.'; return; }
      msg.textContent = list.length + ' PHP and IOP patients checked. PHP plans are reviewed ' + RULE.PHP.label + ', IOP ' + RULE.IOP.label + '. Oldest first.';
      var groups = [
        ['none', 'No plan yet', '#b91c1c', function (a, b) { return (a.adm || 0) - (b.adm || 0); }],
        ['over', 'Review overdue', '#dc2626', function (a, b) { return a.left - b.left; }],
        ['soon', 'Due this week', '#b45309', function (a, b) { return a.left - b.left; }],
        ['err', 'Couldn\'t check', '#555', null]
      ];
      var any = false;
      groups.forEach(function (g) {
        var rows = list.filter(function (x) { return x.g === g[0]; }); if (!rows.length) return; any = true;
        if (g[3]) rows.sort(g[3]);
        var h = document.createElement('div'); h.style.cssText = 'font-weight:700;color:' + g[2] + ';margin:12px 0 4px'; h.textContent = g[1] + ' (' + rows.length + ')'; L.appendChild(h);
        rows.forEach(function (o) {
          var d = document.createElement('div'); d.style.cssText = 'display:flex;justify-content:space-between;align-items:center;gap:10px;padding:8px 10px;border:1px solid #e5e7eb;border-left:4px solid ' + g[2] + ';border-radius:8px;margin:4px 0';
          var info = document.createElement('div');
          var nm = document.createElement('b'); nm.textContent = o.n; info.appendChild(nm);
          var sub = document.createElement('div'); sub.style.cssText = 'color:#555;font-size:13px';
          if (o.g === 'none') sub.textContent = o.p + ' · admitted ' + fmt(o.adm ? new Date(o.adm) : null) + (o.days != null ? ' (' + o.days + ' days ago)' : '') + ' · no treatment plan';
          else if (o.g === 'err') sub.textContent = o.p + ' · open the chart to check';
          else sub.textContent = o.p + ' · last ' + o.kind.toLowerCase() + ' ' + fmt(new Date(o.last)) + ' (' + o.days + ' days ago) · ' + (o.left < 0 ? 'was due ' + fmt(new Date(o.due)) : 'due ' + fmt(new Date(o.due))) + (o.unsigned ? ' · not signed yet' : '');
          info.appendChild(sub); d.appendChild(info);
          var btn = document.createElement('button'); btn.type = 'button';
          btn.style.cssText = 'white-space:nowrap;background:' + g[2] + ';color:#fff;border:0;border-radius:8px;padding:8px 12px;font-weight:600;cursor:pointer';
          if (o.g === 'none' || o.g === 'err' || o.unsigned) { btn.textContent = o.unsigned ? 'Finish and sign' : (o.g === 'none' ? 'Open plans' : 'Open chart'); btn.onclick = function () { location.href = '/client/index/' + o.c + '?section=txplan'; }; }
          else { btn.textContent = 'Update plan now'; btn.onclick = function () { startReview(o, msg).catch(function () { location.href = '/client/index/' + o.c + '?section=txplan'; }); }; }
          d.appendChild(btn); L.appendChild(d);
        });
      });
      if (!any) L.innerHTML = '<div style="padding:12px;background:#ecfdf5;border-radius:8px">Every PHP and IOP patient has a current plan. Nothing is due this week.</div>';
    }
    draw(false);
  }
  return { open: open, badge: badge, load: load };
})();

if (location.hostname === 'chtc.lightningstep.com') { try { LC.notePages(); } catch (e) {} }
})();
