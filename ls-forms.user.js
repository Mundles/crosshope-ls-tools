// ==UserScript==
// @name         CrossHope LS Forms
// @namespace    https://crosshopetreatment.com/ls-tools
// @version      1.1.0
// @description  Adds an "Open a form" button to LightningStep. Pick a form, type the patient's name, and the form opens on their chart.
// @author       CrossHope
// @match        https://chtc.lightningstep.com/*
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
  st.textContent = '@media print{#lsfp-btn,#lsfp{display:none!important}}';
  document.head.appendChild(st);

  var btn = document.createElement('button');
  btn.id = 'lsfp-btn';
  btn.type = 'button';
  btn.textContent = 'Open a form';
  btn.title = 'Pick a form, type the patient\'s name, and the form opens on their chart.';
  btn.style.cssText = 'position:fixed;right:18px;bottom:18px;z-index:2147483646;background:#1a73e8;color:#fff;border:0;border-radius:24px;padding:11px 18px;font:600 15px system-ui,"Segoe UI",sans-serif;cursor:pointer;box-shadow:0 4px 14px rgba(0,0,0,.25)';
  btn.onclick = function () { togglePicker(); };
  document.body.appendChild(btn);

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
})();
