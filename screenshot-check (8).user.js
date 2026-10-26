// ==UserScript==
// @name         Tickets: сверка скриншота (сумма и кошелёк агента)
// @namespace    local.tickets.screenshots
// @version      1.9
// @description  Сверяет скриншот чека в тикете с данными тикета: сумма и кошелёк агента. Данные распознавания сайта подтверждаются распознаванием самой картинки (Google Apps Script с распознаванием Google Drive; без него — слабое распознавание в браузере); обход всех страниц, выгрузка в Excel
// @match        https://REPLACE_ME/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        GM_setClipboard
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      *
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  const CFG = {
    ticketHeader: 'ticket id',               // колонка Ticket ID в списке
    historyHeader: 'ticket history',         // колонка Ticket history (если есть; помогает найти главную таблицу)
    histMarker: 'date of ticket update',     // по этому заголовку узнаём таблицу истории
    supportFilesHeader: /^support.*files?$/,   // колонка Support team's files: скрин может лежать и там
    userFilesHeader: /^user'?s? files?$/,    // колонка User files: оттуда берём ссылки на скриншоты
    historyApi: '/admin/backoffice/paymentsupporthistory',
    concurrency: 2,                  // сколько запросов истории идёт одновременно (сервер отвечает 529 при перегрузке)
    apiGapMs: 250,                   // пауза после каждого запроса в каждом потоке
    apiRetries: 5,                   // повторы при 429/5xx
    apiBackoffMs: 1500,              // первая пауза перед повтором, дальше ×2
    pagerSelector: null,             // например '.my-pager' — если автопоиск пагинации не сработает
    nextSelector: null,              // например '.pagination li:last-child a' — если автопоиск кнопки "далее" не сработает
    pageWaitMs: 20000,               // ожидание загрузки страницы (растёт с каждой повторной попыткой)
    retries: 3,                      // сколько раз пробовать перейти на следующую страницу
  };

  // страница сайта (скрипт работает в изолированной среде Tampermonkey, поэтому перехваты делаем через unsafeWindow)
  const PW = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;

  // запоминаем заголовки, с которыми сам сайт обращается за историей (токены и т.п.), чтобы повторять их в фоновых запросах
  const learned = {};
  (function () {
    try {
      const XP = PW.XMLHttpRequest && PW.XMLHttpRequest.prototype;
      if (!XP) return;
      const o = XP.open, sh = XP.setRequestHeader, sd = XP.send;
      XP.open = function (m, u) { this.__tsU = String(u); this.__tsH = {}; return o.apply(this, arguments); };
      XP.setRequestHeader = function (k, v) { if (this.__tsH) this.__tsH[k] = v; return sh.apply(this, arguments); };
      XP.send = function () {
        if (this.__tsU && this.__tsU.indexOf('paymentsupporthistory') >= 0) Object.assign(learned, this.__tsH);
        return sd.apply(this, arguments);
      };
    } catch (e) { /* перехват необязателен */ }
  })();

  const COLORS = { bad: '#ff7875', review: '#ffe58f', ok: '#d9f7be', error: '#d9d9d9' };

  const norm = (s) => (s || '').replace(/[’`]/g, "'").replace(/\s+/g, ' ').trim().toLowerCase();
  const txt = (el) => (el ? (el.innerText !== undefined ? el.innerText : el.textContent) || '' : '');
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const visible = (el) => !!el && !!(el.offsetWidth || el.offsetHeight || (el.getClientRects && el.getClientRects().length));
  const statusEq = (a, b) => norm(a) === norm(b);
  const clickEl = (el) => (typeof el.click === 'function' ? el.click() : el.dispatchEvent(new MouseEvent('click', { bubbles: true })));
  async function waitFor(fn, ms, step = 150) {
    const t = Date.now();
    for (;;) {
      let v = null;
      try { v = fn(); } catch (e) { v = null; }
      if (v) return v;
      if (Date.now() - t > ms) return null;
      await sleep(step);
    }
  }

  // ---------- настройки (сохраняются между сессиями) ----------
  const SKEY = 'ts_settings_v1';
  const DEFAULTS = { gurl: '', gsecret: '', gkey: '', ocr: true, first: true, hide: false, badges: true, open: false };
  function loadSettings() {
    let raw = null;
    try { raw = typeof GM_getValue === 'function' ? GM_getValue(SKEY, null) : localStorage.getItem(SKEY); } catch (e) { /* ignore */ }
    try { return Object.assign({}, DEFAULTS, raw ? JSON.parse(raw) : {}); } catch (e) { return Object.assign({}, DEFAULTS); }
  }
  const S = loadSettings();
  function saveSettings() {
    try {
      const raw = JSON.stringify(S);
      if (typeof GM_setValue === 'function') GM_setValue(SKEY, raw); else localStorage.setItem(SKEY, raw);
    } catch (e) { /* ignore */ }
  }

  // ---------- таблица истории (чтобы не принять её за список) ----------
  function isHistoryTable(t) {
    return !!t && [...t.querySelectorAll('th')].some((x) => norm(txt(x)).includes(CFG.histMarker)) &&
      ![...t.querySelectorAll('th')].some((x) => norm(txt(x)).includes(CFG.historyHeader));
  }

  // ---------- список жалоб ----------
  function analyze() {
    const ths = [...document.querySelectorAll('th')];
    const cands = ths.filter((th) => norm(txt(th)).includes(CFG.ticketHeader));
    const hasHist = (t) => t && [...t.querySelectorAll('th')].some((x) => norm(txt(x)).includes(CFG.historyHeader));
    const tTh = cands.find((th) => hasHist(th.closest('table'))) || cands.find((th) => !isHistoryTable(th.closest('table')));
    if (!tTh) return null;
    const table = tTh.closest('table');
    const heads = [...tTh.parentElement.children].map((c) => norm(txt(c)));
    const idx = {
      ticket: heads.findIndex((h) => h.includes(CFG.ticketHeader)),
      user: heads.findIndex((h) => CFG.userFilesHeader.test(h)),
      sup: heads.findIndex((h) => CFG.supportFilesHeader.test(h)),
    };
    let trs = [...table.querySelectorAll(':scope > tbody > tr')];
    if (!trs.length) trs = [...table.querySelectorAll('tbody tr')];
    const res = [];
    trs.forEach((tr) => {
      const cells = tr.children;
      if (cells.length <= Math.max(idx.ticket, idx.user, idx.sup)) return;
      const id = (cells[idx.ticket].innerText || '').trim();
      if (!id) return;
      const grab = (i) => (i >= 0 && cells[i]) ? [...cells[i].querySelectorAll('a[href],img[src]')].map((e) => e.getAttribute('href') || e.getAttribute('src')).filter((h) => h && !/^\s*javascript:/i.test(h)) : [];
      const links = (idx.user >= 0 || idx.sup >= 0) ? [...new Set([...grab(idx.user), ...grab(idx.sup)])].slice(0, 5) : null;
      res.push({ tr, id, tCell: cells[idx.ticket], userLinks: links });
    });
    res.table = table; res.idx = idx;
    return res;
  }

  // ---------- сверка скриншотов: данные распознавания сайта (socratResults) против суммы и кошелька тикета ----------
  function parseSoc(v) {
    if (v == null || v === '') return [];
    let x = v;
    if (typeof x === 'string') { try { x = JSON.parse(x); } catch (e) { return []; } }
    if (!Array.isArray(x)) x = x && typeof x === 'object' ? [x] : [];
    return x.filter((o) => o && typeof o === 'object');
  }
  // арабско-индийские (٠-٩) и персидские (۰-۹) цифры → обычные; арабские разделители → . и ,
  function asciiDigits(t) {
    return String(t == null ? '' : t)
      .replace(/[٠-٩]/g, (c) => String(c.charCodeAt(0) - 0x0660))
      .replace(/[۰-۹]/g, (c) => String(c.charCodeAt(0) - 0x06F0))
      .replace(/[\u200E\u200F\u202A-\u202E\u2066-\u2069\u061C]/g, '')
      .replace(/٫/g, '.').replace(/٬/g, ',').replace(/[   ]/g, ' ');
  }
  // все числа, которые можно прочитать в строке (соседний текст на любом языке игнорируется);
  // если разделитель неоднозначен («1.200» — 1,2 или 1200), возвращаются оба варианта
  function sumCands(v, strict) {
    const base = asciiDigits(v);
    const out = new Set();
    const add = (n) => { if (isFinite(n)) out.add(Math.round(n * 1000) / 1000); };
    const read = (str) => {
      [...str.matchAll(/\d[\d.,]*/g)].forEach((m) => {
        const tok = m[0].replace(/[.,]+$/, '');
        if (!tok) return;
        if (strict) {                                                  // «отдельное число»: не часть даты, времени, телефона, номера через - / :
          const prev = str[m.index - 1] || '', next = str[m.index + tok.length] || '', after = str[m.index + tok.length + 1] || '';
          if (/[\-\/:@#_]/.test(prev) && /\d/.test(str[m.index - 2] || '')) return;
          if (/[\-\/:@#_]/.test(next) && /\d/.test(after)) return;
          if (/[A-Za-z_]/.test(prev) && /\d/.test(tok[0]) && /[A-Za-z0-9]/.test(prev)) return;    // «ref20», «TX20» — часть кода
        }
        const lc = tok.lastIndexOf(','), ld = tok.lastIndexOf('.');
        if (lc >= 0 && ld >= 0) {                                    // оба разделителя: последний — десятичный
          add(lc > ld ? parseFloat(tok.replace(/\./g, '').replace(',', '.')) : parseFloat(tok.replace(/,/g, '')));
        } else if (lc >= 0 || ld >= 0) {
          const sep = lc >= 0 ? ',' : '.';
          const parts = tok.split(sep);
          const last = parts[parts.length - 1];
          if (parts.length > 2) add(parseFloat(parts.join('')));       // 1,200,000 — тысячи
          else if (last.length === 3) { add(parseFloat(parts[0] + '.' + last)); add(parseFloat(parts[0] + last)); }   // «5.000»: 5 или 5000
          else add(parseFloat(parts[0] + '.' + last));
        } else {
          add(parseFloat(tok));
          // OCR часто теряет десятичный разделитель: «94.00» → «9400», «94.0» → «940»
          if (/^\d{3,7}$/.test(tok)) {
            if (/00$/.test(tok)) add(parseFloat(tok) / 100);
            if (/0$/.test(tok)) add(parseFloat(tok) / 10);
          }
        }
      });
    };
    read(base);
    read(base.replace(/(?<=\d) (?=\d{3}(?!\d))/g, ''));              // «1 200» — тысячи через пробел
    return [...out];
  }
  const parseSum = (v) => { const c = sumCands(v); return c.length ? c[0] : NaN; };
  const alnum = (x) => String(x == null ? '' : x).toLowerCase().replace(/[^a-z0-9]/g, '');
  // receiver с маской («**** 5021») → видимая часть; не похоже на номер (например название банка) → '' (не сравниваем)
  function visibleTail(recv) {
    const parts = String(recv || '').split(/[*xX•]+/).map(alnum).filter(Boolean);
    const vis = parts.sort((a, b) => b.length - a.length)[0] || '';
    return /\d{4,}/.test(vis) ? vis : '';
  }
  const walletHit = (vis, w) => { w = alnum(w); return !!vis && !!w && (w.endsWith(vis) || w.startsWith(vis) || (vis.length >= 6 && w.includes(vis))); };
  function analyzeShots(entries) {
    const newest = entries.slice().sort((a, b) => String(b.date).localeCompare(String(a.date)))[0] || {};
    const amount = parseSum(newest.amount);
    const seen = new Map();
    entries.forEach((e) => (e.soc || []).forEach((o) => {
      const key = o.documents_user || o.documents_support || JSON.stringify(o).slice(0, 120);
      if (!seen.has(key)) seen.set(key, o);
    }));
    const all = [...seen.values()];
    // чеки с ошибкой распознавания считаем нечитаемыми и не учитываем
    const recs = all.filter((o) => o.classification === 'receipt' && o.extracted_data && typeof o.extracted_data === 'object' && !o.error_message);
    const sh = { n: all.length, m: recs.length, amount: newest.amount == null ? '' : String(newest.amount), agentW: newest.agentW == null ? '' : String(newest.agentW),
      sums: recs.map((o) => String(o.extracted_data.sum == null ? '' : o.extracted_data.sum)), recv: recs.map((o) => String(o.extracted_data.receiver || '')),
      sumRes: 'none', walletRes: 'none', reasons: all.map((o) => o.rejection_reason).filter(Boolean).join(' | ') };
    if (!recs.length) return sh;
    // осторожное правило: «не сошлось» ставим только при однозначном расхождении, любые сомнения → 'unknown' (тикет не отмечается)
    const sums = recs.map((o) => sumCands(o.extracted_data.sum));
    // поиск суммы тикета как отдельного числа в остальном тексте чека (кроме даты, времени, номера операции, получателя, банка)
    const SKIP = /^(date|time|utn|receiver|bank|status|period_start|period_end|currency)$/i;
    const textOf = (o) => [...Object.entries(o.extracted_data).filter(([k]) => !SKIP.test(k) && k !== 'sum').map(([, v]) => v),
      ...Object.entries(o).filter(([k, v]) => /(ocr|raw|text|content)/i.test(k) && typeof v === 'string').map(([, v]) => v)].filter((v) => typeof v === 'string' || typeof v === 'number').join(' \n ');
    const standalone = recs.map((o) => sumCands(textOf(o), true));
    sh.found = '';
    const sumHit = recs.some((o, k) => {
      if (!isNaN(amount) && sums[k].some((x) => Math.abs(x - amount) < 0.005)) { sh.found = 'поле суммы'; return true; }
      if (!isNaN(amount) && standalone[k].some((x) => Math.abs(x - amount) < 0.005)) { sh.found = 'отдельное число в тексте чека'; return true; }
      if (o.field_matches && o.field_matches.sum_match === '1') { sh.found = 'сайт (sum_match)'; return true; }
      return false;
    });
    if (sumHit) sh.sumRes = 'ok';
    else if (isNaN(amount) || amount <= 0 || sums.some((x) => !x.length)) sh.sumRes = 'unknown';
    else sh.sumRes = 'bad';
    const vis = recs.map((o) => visibleTail(o.extracted_data.receiver));
    const agentW = alnum(newest.agentW), userW = alnum(newest.userW);
    if (vis.some((v) => v && walletHit(v, agentW))) sh.walletRes = 'ok';
    else if (!agentW || vis.some((v) => !v)) sh.walletRes = 'unknown';        // нет кошелька в тикете или получатель на каком-то чеке не читается
    else if (userW && vis.some((v) => walletHit(v, userW))) sh.walletRes = 'user';
    else sh.walletRes = 'bad';
    return sh;
  }
  // ---------- распознавание самой картинки (второй проход, только для спорных тикетов) ----------
  // данные сайта по чекам ненадёжны (берёт баланс вместо суммы, номер операции вместо кошелька), поэтому «не сошлось» ставим,
  // только если распознавание картинки тоже не нашло сумму тикета / кошелёк агента
  const TESS_SRC = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
  const ocrWorkers = {};   // lang -> Promise<worker>
  const loadScript = (src) => new Promise((ok, no) => {
    const sc = document.createElement('script');
    sc.src = src; sc.onload = ok;
    sc.onerror = () => no(new Error('не удалось загрузить библиотеку распознавания (возможно, запрещено политикой безопасности сайта)'));
    document.head.appendChild(sc);
  });
  function getOcr(lang) {
    if (!ocrWorkers[lang]) {
      ocrWorkers[lang] = (async () => {
        if (!PW.Tesseract) await loadScript(TESS_SRC);
        if (!PW.Tesseract) throw new Error('библиотека распознавания не загрузилась');
        const w = await PW.Tesseract.createWorker(lang);
        await w.setParameters({ tessedit_pageseg_mode: '11' });   // разреженный текст: ищем все надписи на картинке
        return w;
      })().catch((e) => { delete ocrWorkers[lang]; if (lang === 'eng') e.fatal = true; throw e; });
    }
    return ocrWorkers[lang];
  }
  function gmBlob(url) {
    return new Promise((ok, no) => {
      if (typeof GM_xmlhttpRequest !== 'function') { fetch(url, { credentials: 'include' }).then((r) => (r.ok ? r.blob() : Promise.reject(new Error('HTTP ' + r.status)))).then(ok, no); return; }
      GM_xmlhttpRequest({
        method: 'GET', url, responseType: 'blob', timeout: 30000,
        onload: (r) => (r.status >= 200 && r.status < 300 ? ok(r.response) : no(new Error('HTTP ' + r.status))),
        onerror: () => no(new Error('картинка не скачалась')), ontimeout: () => no(new Error('картинка не скачалась (таймаут)')),
      });
    });
  }
  // увеличиваем, переводим в градации серого, тёмные скрины инвертируем (так распознаётся заметно лучше)
  async function prepImage(blob) {
    const bmp = await createImageBitmap(blob);
    const scale = Math.max(1, Math.min(4, 1200 / bmp.width, 6000 / bmp.height));
    const c = document.createElement('canvas');
    c.width = Math.round(bmp.width * scale); c.height = Math.round(bmp.height * scale);
    const g = c.getContext('2d');
    g.imageSmoothingQuality = 'high';
    g.drawImage(bmp, 0, 0, c.width, c.height);
    const d = g.getImageData(0, 0, c.width, c.height), px = d.data;
    let total = 0;
    for (let k = 0; k < px.length; k += 4) { const l = 0.299 * px[k] + 0.587 * px[k + 1] + 0.114 * px[k + 2]; px[k] = px[k + 1] = px[k + 2] = l; total += l; }
    if (total / (px.length / 4) < 110) for (let k = 0; k < px.length; k += 4) px[k] = px[k + 1] = px[k + 2] = 255 - px[k];
    g.putImageData(d, 0, 0);
    return c;
  }
  async function loadCanvas(url) {
    const blob = await gmBlob(url);
    if (!blob || !/^image\//.test(blob.type || '')) throw new Error('по ссылке не картинка (' + ((blob && blob.type) || 'нет типа') + ')');
    return prepImage(blob);
  }
  async function ocrCanvas(canvas, lang) {
    const w = await getOcr(lang);
    const r = await w.recognize(canvas);
    return { text: (r && r.data && r.data.text) || '', conf: (r && r.data && r.data.confidence) || 0 };
  }
  // ---------- Google Cloud Vision (если задан ключ): читает арабские и персидские цифры, длинные скриншоты — целиком ----------
  const GVISION = 'https://vision.googleapis.com/v1/images:annotate';
  const blobB64 = (blob) => new Promise((ok, no) => {
    const fr = new FileReader();
    fr.onload = () => ok(String(fr.result).split(',')[1] || '');
    fr.onerror = () => no(new Error('не удалось прочитать файл картинки'));
    fr.readAsDataURL(blob);
  });
  async function visionText(blob) {
    if (blob.size > 9.5 * 1024 * 1024) throw new Error('картинка больше 10 МБ — Google Vision её не примет');
    const content = await blobB64(blob);
    const body = JSON.stringify({ requests: [{ image: { content }, features: [{ type: 'DOCUMENT_TEXT_DETECTION' }], imageContext: { languageHints: ['ar', 'fa', 'en'] } }] });
    const url = GVISION + '?key=' + encodeURIComponent(S.gkey);
    const resp = await new Promise((ok, no) => {
      const done = (status, text) => { try { ok({ status, json: JSON.parse(text) }); } catch (e) { no(new Error('ответ Google не JSON (HTTP ' + status + ')')); } };
      if (typeof GM_xmlhttpRequest === 'function') {
        GM_xmlhttpRequest({ method: 'POST', url, headers: { 'Content-Type': 'application/json' }, data: body, timeout: 40000,
          onload: (r) => done(r.status, r.responseText), onerror: () => no(new Error('нет связи с Google Vision')), ontimeout: () => no(new Error('Google Vision: таймаут')) });
      } else {
        fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body }).then((r) => r.text().then((t) => done(r.status, t)), () => no(new Error('нет связи с Google Vision')));
      }
    });
    const j = resp.json || {};
    if (j.error) { const e = new Error('Google Vision: ' + (j.error.message || ('HTTP ' + resp.status))); e.fatal = true; throw e; }   // неверный ключ, не включён API, нет квоты — дальше пробовать бессмысленно
    const r0 = (j.responses || [])[0];
    if (!r0) throw new Error('Google Vision: пустой ответ');
    if (r0.error) throw new Error('Google Vision: ' + (r0.error.message || 'ошибка'));
    const fta = r0.fullTextAnnotation;
    const page = fta && fta.pages && fta.pages[0];
    return { text: (fta && fta.text) || '', conf: page && typeof page.confidence === 'number' ? page.confidence * 100 : (fta && fta.text ? 80 : 0) };
  }
  // ---------- Google Apps Script (распознавание через Google Drive): картинка уходит на ваш скрипт, обратно приходит текст ----------
  async function appsScriptText(blob) {
    const image = await blobB64(blob);
    const body = JSON.stringify({ secret: S.gsecret, mime: blob.type || 'image/jpeg', image });
    const resp = await new Promise((ok, no) => {
      const done = (status, text) => {
        try { ok(JSON.parse(text)); } catch (e) { no(Object.assign(new Error('Apps Script вернул не JSON (HTTP ' + status + '). Проверьте, что приложение развёрнуто как «Веб-приложение» с доступом «Все» и адрес заканчивается на /exec'), { fatal: true })); }
      };
      if (typeof GM_xmlhttpRequest === 'function') {
        GM_xmlhttpRequest({ method: 'POST', url: S.gurl, headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, data: body, timeout: 90000,
          onload: (r) => done(r.status, r.responseText), onerror: () => no(new Error('нет связи с Apps Script')), ontimeout: () => no(new Error('Apps Script: таймаут')) });
      } else {
        fetch(S.gurl, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body }).then((r) => r.text().then((t) => done(r.status, t)), () => no(new Error('нет связи с Apps Script')));
      }
    });
    if (!resp.ok) throw Object.assign(new Error('Apps Script: ' + (resp.error || 'ошибка')), { fatal: !!resp.fatal || /Drive|permission|авториз|authoriz|scope|not found.*service/i.test(resp.error || '') });
    return { text: resp.text || '', conf: resp.text ? 70 : 0 };     // Google Drive не отдаёт уверенность — считаем текст по наличию чисел
  }
  // единая точка распознавания: Google Vision, если задан ключ, иначе английская модель в браузере
  async function recognizeUrl(url) {
    const blob = await gmBlob(url);
    if (!blob || !/^image\//.test(blob.type || '')) throw new Error('по ссылке не картинка (' + ((blob && blob.type) || 'нет типа') + ')');
    if (S.gurl) { const o = await appsScriptText(blob); o.engine = 'apps-script'; o.bytes = blob.size; return o; }
    if (S.gkey) { const o = await visionText(blob); o.engine = 'google'; o.bytes = blob.size; return o; }
    const o = await ocrCanvas(await prepImage(blob), 'eng');
    o.engine = 'tesseract'; o.bytes = blob.size;
    return o;
  }
  // текст считаем прочитанным, только если уверенность распознавания не низкая и в нём есть хотя бы два числа из 2+ цифр
  const readable = (o) => o.conf >= 50 && (asciiDigits(o.text).match(/\d{2,}/g) || []).length >= 2;
  async function ocrCheckRow(res, links) {
    const out = { ok: false, sumFound: false, walletFound: false, siteSeen: false, longNum: false, n: 0, err: '', text: '', fatal: false };
    const amount = parseSum(res.shot.amount);
    const w = alnum(res.shot.agentW);
    // сумма с комиссией: на чеке есть итог T и небольшая комиссия F (до 15%), и T − F или T + F равно сумме тикета (24.50 − 0.50 = 24)
    const withFee = (nums, amt) => {
      if (!(amt > 0)) return false;
      const u = [...new Set(nums.filter((x) => x > 0 && x < 1e7))].slice(0, 60);
      return u.some((t) => u.some((f) => f !== t && f <= 0.15 * Math.max(t, amt) && (Math.abs(t - f - amt) < 0.005 || Math.abs(t + f - amt) < 0.005)));
    };
    const siteC = (res.shot.sums || []).flatMap((x) => sumCands(x));
    const fuzzyHas = (hay, needle) => {                              // допускаем потерянные/неверные символы (до 2 для длинных кошельков): HHHOU856 ≈ HHH856
      const m = needle.length, maxD = m >= 8 ? 2 : 1;
      let prev = Array.from({ length: m + 1 }, (_, k) => k);
      for (let t = 0; t < hay.length; t++) {
        const cur = [0];
        for (let k = 1; k <= m; k++) cur[k] = Math.min(prev[k] + 1, cur[k - 1] + 1, prev[k - 1] + (needle[k - 1] === hay[t] ? 0 : 1));
        if (cur[m] <= maxD) return true;
        prev = cur;
      }
      return false;
    };
    // trusted — текст прочитан уверенно. Найденное совпадение (сумма тикета, кошелёк агента) засчитываем из ЛЮБОГО прохода,
    // а «подтверждающие» признаки для отметки (сумма сайта на картинке, длинный номер) — только из уверенно прочитанного текста
    const scan = (text, trusted) => {
      const nums = sumCands(text, true);
      if (!isNaN(amount) && (nums.some((x) => Math.abs(x - amount) < 0.005) || withFee(nums, amount))) out.sumFound = true;
      if (trusted && siteC.some((c) => nums.some((x) => Math.abs(x - c) < 0.005))) out.siteSeen = true;     // картинка подтверждает ту сумму, которую назвал сайт
      const dg = asciiDigits(text);
      if (trusted && /\d{9,}/.test(dg)) out.longNum = true;                                                 // на картинке прочитан длинный номер (кошелёк / операция)
      if (w.length < 8 || fuzzyHas(alnum(dg), w.slice(-10))) out.walletFound = true;
    };
    for (const href of links.slice(0, 3)) {
      try {
        const p1 = await recognizeUrl(new URL(href, location.origin).href);
        let any = readable(p1);
        scan(p1.text, any); out.text += ' ' + p1.text.replace(/\s+/g, ' ').trim().slice(0, 200);
        if (any) { out.ok = true; out.n++; } else out.err = out.err || 'текст на картинке не распознан (низкая уверенность)';
      } catch (e) { out.err = e.message; if (e.fatal) { out.fatal = true; break; } }
    }
    return out;
  }
  const CHECK_WALLET = false;   // сверка кошелька агента пока отключена — проверяем только сумму
  const walSus = (r) => CHECK_WALLET && !!r.shot && (r.shot.walletRes === 'bad' || r.shot.walletRes === 'user');
  const siteSuspect = (r) => !!r.shot && (r.shot.sumRes === 'bad' || walSus(r));
  const ocrConfirmed = (r) => !!(S.ocr && r.ocr && r.ocr.ok);
  // сумму отмечаем, только если картинка подтверждает ДРУГУЮ сумму (ту же, что назвал сайт), а суммы тикета на ней нет; если число просто не прочиталось — не отмечаем
  const sumBad = (r) => !r.error && !!r.shot && r.shot.sumRes === 'bad' && ocrConfirmed(r) && !r.ocr.sumFound && !!r.ocr.siteSeen;
  const walletBad = (r) => !r.error && walSus(r) && ocrConfirmed(r) && !r.ocr.walletFound && !!r.ocr.longNum;
  const isShot = (r) => sumBad(r) || walletBad(r);
  const SHOT_TXT = { bad: 'не совпала', user: 'совпала с кошельком ПОЛЬЗОВАТЕЛЯ', ok: 'совпала', unknown: 'не определить', none: 'нет чека' };

  // ---------- фоновый режим: запрос истории без открытия окна ----------
  // ищет в ответе массив объектов (строки истории)
  function findRows(node, keys, depth) {
    if (!node || typeof node !== 'object' || depth > 6) return null;
    if (Array.isArray(node)) {
      if (node.length && node[0] && typeof node[0] === 'object' && keys.some((k) => k in node[0])) return node;
      for (const x of node) { const r = findRows(x, keys, depth + 1); if (r) return r; }
      return null;
    }
    for (const k of Object.keys(node)) { const r = findRows(node[k], keys, depth + 1); if (r) return r; }
    return null;
  }
  // строки истории: нам нужны данные распознавания (socratResults), сумма и кошельки
  function parseApiHistory(j) {
    const arr = findRows(j, ['dateEdit', 'socratResults'], 0);
    if (!arr) throw new Error('в ответе нет строк истории');
    return arr.map((o) => ({ date: String(o.dateEdit == null ? '' : o.dateEdit), soc: parseSoc(o.socratResults), amount: o.amount, agentW: o.walletAccountNumber, userW: o.numberWallet }));
  }
  async function fetchHistoryRaw(id) {
    if (!/^\d+$/.test(id)) throw new Error('Ticket ID не число');
    const headers = Object.assign({ 'Content-Type': 'application/json;charset=UTF-8', Accept: 'application/json, text/plain, */*', 'X-Requested-With': 'XMLHttpRequest' }, learned);
    let r;
    for (let a = 0; ; a++) {
      r = await fetch(location.origin + CFG.historyApi, { method: 'POST', credentials: 'include', headers, body: JSON.stringify({ ticketId: Number(id), is_iframe: 1 }) });
      if (r.ok) break;
      const retryable = r.status === 429 || r.status >= 500;
      if (!retryable || a >= CFG.apiRetries) throw new Error('HTTP ' + r.status);
      const ra = Number(r.headers.get('Retry-After'));
      await sleep(ra > 0 ? Math.min(ra * 1000, 30000) : CFG.apiBackoffMs * Math.pow(2, a));
    }
    try { return await r.json(); } catch (e) { throw new Error('ответ не JSON'); }
  }
  const fetchHistoryApi = async (id) => parseApiHistory(await fetchHistoryRaw(id));

  // ---------- меню ----------
  const style = document.createElement('style');
  style.textContent = `
    #ts-panel, #ts-panel * { box-sizing: border-box; }
    #ts-panel .ts-sec { font-weight: 700; margin: 8px 0 4px; color: #23909f; font-size: 11px; text-transform: uppercase; letter-spacing: .4px; }
    #ts-panel label { display: block; margin-bottom: 4px; cursor: pointer; }
    #ts-panel button { width: 100%; margin-bottom: 4px; }
    #ts-panel .ts-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 4px; margin-bottom: 4px; }
    #ts-panel .ts-grid button { margin: 0; padding: 3px 2px; font-size: 12px; }
    #ts-panel pre { max-height: 160px; overflow: auto; font-size: 11px; white-space: pre-wrap; background: #f5f5f5; padding: 4px; margin: 4px 0 0; }
    td[data-tsb]::after { content: attr(data-tsb); display: block; margin-top: 2px; font: 600 10px/1.2 Arial, sans-serif; color: inherit; white-space: normal; }
  `;
  document.head.appendChild(style);

  const launcher = document.createElement('button');
  launcher.type = 'button';
  launcher.title = 'Сверка скриншота с тикетом';
  launcher.style.cssText =
    'position:fixed;right:188px;bottom:12px;z-index:999999;background:#c41d7f;color:#fff;border:none;border-radius:20px;' +
    'box-sizing:border-box;width:80px;padding:8px 0;text-align:center;white-space:nowrap;font:700 13px Arial,sans-serif;cursor:pointer;box-shadow:0 2px 10px rgba(0,0,0,.35)';
  launcher.textContent = 'Скрины';
  document.body.appendChild(launcher);

  const panel = document.createElement('div');
  panel.id = 'ts-panel';
  panel.style.cssText =
    'position:fixed;right:188px;bottom:56px;z-index:999999;background:#fff;border:2px solid #c41d7f;border-radius:8px;' +
    'padding:10px;font:13px/1.4 Arial,sans-serif;width:300px;max-height:calc(100vh - 80px);overflow:auto;' +
    'box-shadow:0 4px 16px rgba(0,0,0,.3);color:#222;display:none';
  panel.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center"><b>Сверка скриншота с тикетом</b><span id="ts-x" style="cursor:pointer;font-size:16px;padding:0 4px" title="Закрыть меню">✕</span></div>
    <div id="ts-sets">
      <div class="ts-sec">Проверка</div>
      <div style="font-size:11px;color:#666;margin-bottom:4px">Сумма тикета и кошелёк агента сверяются со скриншотом чека. Сначала данные распознавания сайта; если они не сошлись — распознаётся сама картинка. Отмечается, только если расхождение подтверждено; всё сомнительное идёт в «проверить вручную».</div>
      <label><input type="checkbox" id="ts-ocr"> Распознавать картинки для спорных тикетов (рекомендуется)</label>
      <div style="font-size:11px;color:#666;margin:4px 0 2px">Распознавание через Google Apps Script (читает арабские и персидские цифры). Адрес приложения (оканчивается на /exec):</div>
      <input type="text" id="ts-gurl" placeholder="https://script.google.com/macros/s/…/exec" style="width:100%;margin-bottom:4px" autocomplete="off">
      <input type="password" id="ts-gsecret" placeholder="секретное слово из скрипта" style="width:100%;margin-bottom:4px" autocomplete="off">
      <div style="font-size:11px;color:#666;margin-bottom:4px">Пока адрес не задан — слабое распознавание в браузере (арабские цифры не читает).</div>
      <div class="ts-sec">Отображение</div>
      <label><input type="checkbox" id="ts-hide"> Скрыть тикеты без проблем</label>
      <label><input type="checkbox" id="ts-badges"> Подписи в ячейке</label>
      <div style="font-size:11px;color:#666;margin-bottom:4px">Красный — не сошлось (подтверждено), жёлтый — проверить вручную, зелёный — сошлось.</div>
    </div>
    <div class="ts-sec">Результат</div>
    <div id="ts-stat" style="margin-bottom:6px">—</div>
    <button id="ts-run">Проверить страницу</button>
    <button id="ts-retry">Повторить ошибки</button>
    <div class="ts-sec">Копировать ID</div>
    <div class="ts-grid"><button id="ts-c-bad">Не сошлось</button><button id="ts-c-review">Вручную</button><button id="ts-c-error">Ошибки</button></div>
    <div class="ts-sec">Все страницы</div>
    <label><input type="checkbox" id="ts-first"> Начать с 1-й страницы</label>
    <button id="ts-crawl">Все страницы → проверить</button>
    <button id="ts-resume" disabled>Продолжить с текущей</button>
    <button id="ts-stop" disabled>Стоп</button>
    <button id="ts-xlsx" disabled>Скачать Excel</button>
    <button id="ts-diag">Диагностика</button>
    <button id="ts-ocrtest">Тест распознавания скрина (1 тикет)</button>
    <div id="ts-msg" style="font-size:12px;color:#555"></div>
    <pre id="ts-pre" style="display:none"></pre>`;
  document.body.appendChild(panel);
  const $ = (q) => panel.querySelector(q);
  const msg = (t) => ($('#ts-msg').textContent = t);
  const setStat = (t) => ($('#ts-stat').textContent = t);

  const showPanel = (v) => { S.open = v; panel.style.display = v ? 'block' : 'none'; saveSettings(); };
  launcher.onclick = () => showPanel(panel.style.display === 'none');
  $('#ts-x').onclick = () => showPanel(false);
  showPanel(!!S.open);

  function syncUI() {
    $('#ts-gurl').value = S.gurl || ''; $('#ts-gsecret').value = S.gsecret || ''; $('#ts-ocr').checked = !!S.ocr; $('#ts-hide').checked = !!S.hide; $('#ts-badges').checked = !!S.badges; $('#ts-first').checked = !!S.first;
  }
  const changed = () => { saveSettings(); run(); };
  $('#ts-gurl').onchange = (e) => {
    const u = e.target.value.replace(/^[\s"'`]+|[\s"'`]+$/g, '');
    e.target.value = u; S.gurl = u; saveSettings();
    if (!u) msg('Адрес удалён. Распознавание — в браузере.');
    else if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(u)) msg('⚠ Адрес выглядит неверно: должен начинаться с https://script.google.com/ и заканчиваться на /exec (это адрес «Веб-приложения», а не редактора).');
    else msg('Адрес сохранён. Введите секретное слово и нажмите «Тест распознавания скрина».');
  };
  $('#ts-gsecret').onchange = (e) => { S.gsecret = e.target.value.trim(); saveSettings(); msg('Секретное слово сохранено.'); };
  $('#ts-ocr').onchange = (e) => { S.ocr = e.target.checked; changed(); };
  $('#ts-hide').onchange = (e) => { S.hide = e.target.checked; changed(); };
  $('#ts-badges').onchange = (e) => { S.badges = e.target.checked; changed(); };
  $('#ts-first').onchange = (e) => { S.first = e.target.checked; saveSettings(); };

  const results = new Map();   // id -> результат проверки
  let busy = false;            // идёт проверка (не перерисовываем список в это время)
  let running = false;
  let aborted = false;
  let curPageNo = 0;

  const newRes = (row) => ({ id: row.id, error: true, shot: null, ocr: null, n: 0, note: '', page: '', via: '' });

  // итог по тикету: error / none (чеков нет) / ok / review (проверить вручную) / bad (не сошлось, подтверждено)
  const st = (r) => {
    if (r.error) return 'error';
    if (!r.shot || !r.shot.m) return 'none';
    if (!siteSuspect(r)) return 'ok';
    if (!S.ocr || !r.ocr || !r.ocr.ok) return 'review';
    if (isShot(r)) return 'bad';
    const sumOK = r.shot.sumRes !== 'bad' || r.ocr.sumFound;
    const walOK = !walSus(r) || r.ocr.walletFound;
    return sumOK && walOK ? 'ok' : 'review';
  };
  function reviewReason(r) {
    if (!S.ocr) return 'распознавание картинок выключено';
    if (!r.ocr) return 'ждёт распознавания картинки';
    if (!r.ocr.ok) return r.ocr.err || 'картинку не удалось прочитать';
    const p = [];
    if (r.shot && r.shot.sumRes === 'bad' && !r.ocr.sumFound) p.push('суммы тикета нет среди прочитанных чисел' + (r.ocr.siteSeen ? '' : ', сумма сайта на картинке тоже не найдена'));
    if (walSus(r) && !r.ocr.walletFound) p.push('кошелёк агента не найден' + (r.ocr.longNum ? '' : ' (длинный номер не прочитан)'));
    return 'расхождение не подтверждено картинкой: ' + (p.join('; ') || 'числа не прочитаны уверенно');
  }

  function paint(rows) {
    rows.forEach((r) => {
      const res = results.get(r.id);
      let color = '', badge = '', title = '';
      if (res) {
        const s = st(res);
        color = COLORS[s] || '';
        if (s === 'error') { badge = 'ошибка'; title = res.note; }
        else if (s !== 'none') {
          if (res.shot) title = `Чеков распознано сайтом: ${res.shot.m} из ${res.shot.n}; сумма тикета ${res.shot.amount}, на чеке ${res.shot.sums.join(' / ') || '—'} (${SHOT_TXT[res.shot.sumRes]}); кошелёк агента ${res.shot.agentW || '—'}, получатель на чеке ${res.shot.recv.join(' / ') || '—'} (${SHOT_TXT[res.shot.walletRes]})`;
          if (res.ocr) title += '\nРаспознавание картинки: ' + (res.ocr.ok ? `сумма тикета ${res.ocr.sumFound ? 'найдена' : 'не найдена'}${res.ocr.siteSeen ? ' (на картинке видна сумма сайта)' : ''}, кошелёк агента ${res.ocr.walletFound ? 'найден' : 'не найден'}${res.ocr.longNum ? '' : ' (длинных номеров не прочитано)'}` : 'не удалось (' + res.ocr.err + ')');
          if (s === 'bad') badge = [sumBad(res) ? 'суммы тикета нет на скрине' : '', walletBad(res) ? 'кошелька агента нет на скрине' : ''].filter(Boolean).join(' · ');
          else if (s === 'review') { badge = 'проверить вручную'; title += '\n' + reviewReason(res); }
        }
      }
      r.tCell.style.background = color;
      r.tCell.style.fontWeight = color && res && (st(res) === 'bad' || st(res) === 'review') ? '700' : '';
      r.tCell.title = title;
      if (S.badges && badge) r.tCell.setAttribute('data-tsb', badge); else r.tCell.removeAttribute('data-tsb');
      r.tr.style.display = S.hide && res && ['ok', 'none'].includes(st(res)) ? 'none' : '';
    });
  }
  function counts() {
    const v = [...results.values()].map(st);
    const c = (k) => v.filter((x) => x === k).length;
    return { bad: c('bad'), review: c('review'), ok: c('ok'), none: c('none'), error: c('error') };
  }
  function run() {
    const rows = analyze();
    if (!rows) { setStat('Таблица не найдена'); return; }
    paint(rows);
    const c = counts();
    const w = (rows.idx.user < 0 && rows.idx.sup < 0) ? '<div style="color:#d4380d;font-size:12px">⚠ Нет колонки User files (включите её в Column settings) — без неё картинки не распознать</div>' : '';
    $('#ts-stat').innerHTML = `Строк: <b>${rows.length}</b> · к проверке: <b>${rows.filter((r) => !results.has(r.id)).length}</b><br>` +
      `всего проверено: сошлось <b style="color:#389e0d">${c.ok}</b> · не сошлось <b style="color:#cf1322">${c.bad}</b> · вручную <b style="color:#ad8b00">${c.review}</b> · без чека ${c.none} · ошибок <b>${c.error}</b>${w}`;
  }

  let apiAllFailed = false;
  async function apiPass() {
    const L = analyze();
    if (!L) return;
    const rows = L.filter((r) => !results.has(r.id));
    if (!rows.length) return;
    busy = true;
    let ok = 0, fail = 0, firstErr = '', next = 0;
    const worker = async () => {
      while (!aborted) {
        const row = rows[next++];
        if (!row) return;
        const res = newRes(row);
        res.page = curPageNo || ''; res.via = 'фон';
        try {
          const entries = await fetchHistoryApi(row.id);
          res.shot = analyzeShots(entries); res.n = entries.length; res.error = false;
          ok++;
        } catch (e) { fail++; res.note = e.message; if (!firstErr) firstErr = e.message; }
        results.set(row.id, res);
        setStat(`Запрашиваю историю: ${ok + fail} из ${rows.length}…`);
        await sleep(CFG.apiGapMs);
      }
    };
    await Promise.all(Array.from({ length: CFG.concurrency }, worker));
    busy = false;
    apiAllFailed = ok === 0 && fail > 0;
    run();
    if (apiAllFailed) msg(`История не читается (${firstErr}). Нажмите «Диагностика».`);
    else if (fail > 0) msg(`История получена для ${ok}, не удалось ${fail} (${firstErr}) — см. «Повторить ошибки».`);
  }

  async function ocrPass() {
    const L = analyze();
    if (!L) return;
    const rows = L.filter((r) => { const x = results.get(r.id); return x && !x.error && !x.ocr && siteSuspect(x); });
    if (!rows.length) return;
    busy = true;
    let n = 0, okc = 0, errMsg = '', next = 0, stop = false;
    const threads = S.gurl ? 3 : 1;                                // Apps Script медленнее (несколько секунд на картинку) — берём 3 потока
    const worker = async () => {
      while (!aborted && !stop) {
        const r = rows[next++];
        if (!r) return;
        const res = results.get(r.id);
        if (!r.userLinks || !r.userLinks.length) { res.ocr = { ok: false, err: 'в списке нет ссылок на файлы пользователя' }; n++; continue; }
        res.ocr = await ocrCheckRow(res, r.userLinks);
        n++;
        setStat(`Распознаю скриншоты: ${n} из ${rows.length}…`);
        if (res.ocr.ok) okc++; else errMsg = res.ocr.err;
        if (res.ocr.fatal) stop = true;
      }
    };
    await Promise.all(Array.from({ length: threads }, worker));
    busy = false;
    run();
    if (!okc && rows.length) msg(`Распознавание картинок не сработало: ${errMsg}. Скрины не отмечены.`);
  }

  // проверка всех непроверенных тикетов текущей страницы; 'ok' | 'errors' | 'aborted'
  async function checkPage() {
    const L = analyze();
    if (!L) return 'ok';
    await apiPass();
    if (aborted) return 'aborted';
    if (apiAllFailed) return 'errors';
    if (S.ocr) { await ocrPass(); if (aborted) return 'aborted'; }
    return 'ok';
  }

  function setRunning(v) {
    running = v;
    ['#ts-run', '#ts-retry', '#ts-diag', '#ts-ocrtest', '#ts-crawl'].forEach((q) => { $(q).disabled = v; });
    $('#ts-resume').disabled = v || pagesDone === 0;
    $('#ts-stop').disabled = !v;
    $('#ts-xlsx').disabled = v || results.size === 0;
    panel.querySelectorAll('#ts-sets input').forEach((el) => { el.disabled = v; });
  }

  $('#ts-run').onclick = async () => {
    if (running) return;
    aborted = false; curPageNo = 0; setRunning(true);
    const r = await checkPage();
    run(); setRunning(false);
    if (r === 'errors') return;   // сообщение об ошибке уже в строке статуса
    msg(r === 'aborted' ? 'Остановлено.' : 'Страница проверена.');
  };
  $('#ts-retry').onclick = () => {
    [...results.entries()].forEach(([id, r]) => { if (r.error) results.delete(id); });
    run(); $('#ts-run').onclick();
  };

  function copy(text) {
    if (typeof GM_setClipboard === 'function') GM_setClipboard(text);
    else navigator.clipboard.writeText(text);
  }
  const LBL = { bad: 'не сошлось', review: 'проверить вручную', error: 'ошибки' };
  ['bad', 'review', 'error'].forEach((k) => {
    $('#ts-c-' + k).onclick = () => {
      const ids = [...results.values()].filter((r) => st(r) === k).map((r) => r.id);
      copy(ids.join('\n'));
      msg(`Скопировано (${LBL[k]}): ${ids.length}`);
    };
  });

  // ---------- обход всех страниц ----------
  const getTotals = () => {
    const m = document.body.innerText.match(/from\s+(\d+)\s+to\s+(\d+)\s+of\s+(\d+)/i);
    return m ? { from: +m[1], to: +m[2], total: +m[3] } : null;
  };
  const signature = () => {
    const t = getTotals();
    const a = analyze();
    return (t ? t.from + '-' + t.to : '') + '|' + (a && a[0] ? a[0].id : '');
  };

  // пагинация ищется по содержимому: ряд элементов вида «‹ 1 … 4 5 6 … 39 ›»
  const PAGER_TEXT = /^(\d+|…|\.{2,3}|[‹›«»<>])?$/;
  const isNum = (el) => /^\d+$/.test((el.innerText || '').trim());
  let pageSize = 0;

  function findPagerBox() {
    if (CFG.pagerSelector) return document.querySelector(CFG.pagerSelector);
    const cands = [...document.querySelectorAll('ul,ol,nav,div')].filter((el) => {
      if (panel.contains(el) || el.closest('#ts-panel,#tn-panel,#th-panel') || el.closest('table')) return false;
      const kids = [...el.children];
      if (kids.length < 4 || kids.length > 25) return false;
      if (!kids.every((k) => PAGER_TEXT.test((k.innerText || '').trim()))) return false;
      return kids.filter(isNum).length >= 3;
    });
    cands.sort((a, b) => a.querySelectorAll('*').length - b.querySelectorAll('*').length);
    return cands[0] || null;
  }

  function getPager() {
    const box = findPagerBox();
    if (!box) return null;
    const items = [...box.children];
    let active = items.find(
      (k) => /active|current|selected/i.test(k.className || '') ||
        k.getAttribute('aria-current') ||
        k.querySelector('.active,.current,.selected,[aria-current]')
    );
    const t = getTotals();
    if (t) pageSize = Math.max(pageSize, t.to - t.from + 1);
    if (!active && t && pageSize) {
      const cur = Math.floor((t.from - 1) / pageSize) + 1;
      active = items.find((k) => (k.innerText || '').trim() === String(cur));
    }
    return { box, items, active };
  }

  // запасной поиск кнопки «далее» по тексту / aria-label / title / rel / классу
  function findNextFallback() {
    const re = /^(›|»|>|next|вперед|вперёд|далее|следующая)$/i;
    const cands = [...document.querySelectorAll('a,button,li,span,[role=button]')].filter((el) => {
      if (panel.contains(el) || el.closest('#ts-panel,#tn-panel,#th-panel') || el.closest('table') || !visible(el)) return false;
      const attrs = [el.getAttribute('aria-label'), el.getAttribute('title'), el.getAttribute('rel')].filter(Boolean).join(' ');
      return re.test((el.innerText || '').trim()) || /(^|\s)next(\s|$)/i.test(attrs) || /(^|[\s_-])next([\s_-]|$)/i.test(el.getAttribute('class') || '');
    });
    return cands.find((el) => !isDisabled(el)) || cands[0] || null;
  }
  function findNext() {
    if (CFG.nextSelector) return document.querySelector(CFG.nextSelector);
    const p = getPager();
    if (!p) return findNextFallback();
    const lastItem = p.items[p.items.length - 1];
    if (!isNum(lastItem)) return lastItem;            // стрелка «›»
    const i = p.items.indexOf(p.active);
    return (i >= 0 && i < p.items.length - 1 ? p.items[i + 1] : null) || findNextFallback();
  }
  function findFirst() {
    const p = getPager();
    if (!p) return null;
    return p.items.find((el) => (el.innerText || '').trim() === '1') || null;
  }
  const isDisabled = (el) =>
    !el || [el, el.querySelector('a,button')].some(
      (n) => n && (n.classList.contains('disabled') || n.getAttribute('aria-disabled') === 'true' || n.hasAttribute('disabled'))
    );
  const clickItem = (el) => (el.querySelector('a,button') || el).click();

  // ждём смены страницы и пока таблица "устаканится"
  async function settle(t0, limit) {
    let prev = signature(), same = 0;
    while (same < 2 && Date.now() - t0 < limit) {
      await sleep(300);
      const cur = signature();
      const n = (analyze() || []).length;
      if (cur === prev && n > 0) same++; else { same = 0; prev = cur; }
    }
    return same >= 2;
  }
  async function waitChange(before, limit) {
    const t0 = Date.now();
    while (Date.now() - t0 < limit) {
      await sleep(300);
      if (signature() !== before) return settle(t0, limit);
    }
    return false;
  }

  // клик с повторами; если страница сменилась с опозданием — повторно не кликаем (чтобы не перескочить страницу)
  async function clickAndWait(getEl) {
    const before = signature();
    for (let attempt = 1; attempt <= CFG.retries; attempt++) {
      if (signature() !== before) { if (await settle(Date.now(), CFG.pageWaitMs)) return 'ok'; }
      const el = getEl();
      if (!el) return 'missing';
      if (isDisabled(el)) return 'disabled';
      clickItem(el);
      if (await waitChange(before, CFG.pageWaitMs * attempt)) return 'ok';
      if (aborted) return 'aborted';
      msg(`Страница долго грузится, повтор ${attempt}/${CFG.retries}…`);
      await sleep(1500);
    }
    return signature() !== before ? 'ok' : 'timeout';
  }


  let visited = new Set();     // «from» каждой пройденной страницы
  let pagesDone = 0;

  function coverage() {
    const t = getTotals();
    if (!t || !pageSize) return '';
    const expected = Math.ceil(t.total / pageSize);
    const missing = [];
    for (let i = 1; i <= expected; i++) if (!visited.has((i - 1) * pageSize + 1)) missing.push(i);
    return missing.length
      ? ` ⚠ Пройдено ${expected - missing.length} из ${expected} страниц. Не проверены: ${missing.slice(0, 15).join(', ')}${missing.length > 15 ? '…' : ''}.`
      : ` ✔ Пройдены все ${expected} страниц.`;
  }

  const WHY = {
    missing: 'не найдена кнопка «далее» / блок пагинации (нажмите «Диагностика» и пришлите результат)',
    disabled: 'кнопка «далее» неактивна',
    timeout: `страница не загрузилась после ${CFG.retries} попыток`,
    aborted: 'остановлено кнопкой',
  };


  async function allPages(resume) {
    if (running) return;
    aborted = false;
    if (!resume) { visited = new Set(); pagesDone = 0; }
    setRunning(true);
    let reason = '';
    let note = '';
    try {
      // 1) перейти на первую страницу (не при продолжении)
      if (!resume && S.first) {
        const t = getTotals();
        if (t && t.from > 1) {
          const first = findFirst();
          if (!first) note = 'кнопка «1» не найдена, старт с текущей. ';
          else if (!isDisabled(first)) {
            const r = await clickAndWait(findFirst);
            if (r !== 'ok') note = 'не удалось перейти на 1-ю страницу, старт с текущей. ';
          }
        }
      }

      // 2) обход: на каждой странице проверяем историю всех тикетов нужных статусов
      while (!aborted && pagesDone < 1000) {
        run();
        const t = getTotals();
        if (t) { pageSize = Math.max(pageSize, t.to - t.from + 1); visited.add(t.from); }
        pagesDone++;
        curPageNo = pagesDone;
        msg(`Страница ${pagesDone}${t ? ` (${t.to} из ${t.total})` : ''}…`);
        const r = await checkPage();
        if (r === 'errors') { reason = 'история не читается (3 ошибки подряд) — нажмите «Диагностика (1 тикет)»'; break; }
        if (r === 'aborted') break;
        msg(`Страница ${pagesDone} готова · проблемных: ${counts().bad}, ошибок: ${counts().error}`);
        if (t && t.to >= t.total) { reason = 'последняя страница'; break; }
        const nr = await clickAndWait(findNext);
        if (nr !== 'ok') { reason = WHY[nr] || nr; break; }
      }
      if (aborted) reason = WHY.aborted;
    } catch (e) {
      reason = 'ошибка: ' + e.message;
      console.error('[ts]', e);
    }
    busy = false;
    run();
    const c = counts();
    msg(`${note}Конец: ${reason || '—'}. Страниц: ${pagesDone}. Не сошлось: ${c.bad}, проверить вручную: ${c.review}, ошибок: ${c.error}.${coverage()}`);
    setRunning(false);
  }

  $('#ts-resume').onclick = () => allPages(true);
  $('#ts-crawl').onclick = () => allPages(false);
  $('#ts-stop').onclick = () => { aborted = true; };

  // Excel: «Не сошлось» / «Проверить вручную» / «Ошибки»
  $('#ts-xlsx').onclick = () => {
    if (typeof XLSX === 'undefined') { msg('Библиотека XLSX не загрузилась.'); return; }
    const numId = (id) => (/^\d{1,15}$/.test(id) ? Number(id) : id);
    const yn = (b) => (b ? 'да' : 'нет');
    const all = [...results.values()];
    const wb = XLSX.utils.book_new();
    const add = (name, aoa, wch) => {
      const sh = XLSX.utils.aoa_to_sheet(aoa);
      sh['!cols'] = wch.map((w) => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, sh, name);
    };
    const head = ['ticket_id', 'сумма тикета', 'сумма на чеке (данные сайта)', 'сумма', 'кошелёк агента (тикет)', 'получатель на чеке (данные сайта)', 'кошелёк', 'чеков распознано', 'скринов всего', 'картинка: сумма найдена', 'картинка: кошелёк агента найден', 'текст с картинки (фрагмент)', 'причина', 'page'];
    const line = (r) => [numId(r.id), r.shot.amount, r.shot.sums.join(' / '), SHOT_TXT[r.shot.sumRes], r.shot.agentW, r.shot.recv.join(' / '), SHOT_TXT[r.shot.walletRes], r.shot.m, r.shot.n,
      r.ocr && r.ocr.ok ? yn(r.ocr.sumFound) : '—', r.ocr && r.ocr.ok ? yn(r.ocr.walletFound) : '—', r.ocr ? (r.ocr.text || r.ocr.err || '').slice(0, 300) : '',
      st(r) === 'bad' ? [sumBad(r) ? 'суммы тикета нет на скрине' : '', walletBad(r) ? 'кошелька агента нет на скрине' : ''].filter(Boolean).join('; ') : reviewReason(r), r.page];
    const W = [14, 12, 22, 14, 26, 26, 30, 12, 12, 14, 18, 60, 40, 6];
    const bad = all.filter((r) => st(r) === 'bad'), rev = all.filter((r) => st(r) === 'review'), err = all.filter((r) => st(r) === 'error');
    add('Не сошлось', [head, ...bad.map(line)], W);
    add('Проверить вручную', [head, ...rev.map(line)], W);
    add('Ошибки', [['ticket_id', 'причина', 'page'], ...err.map((r) => [numId(r.id), r.note, r.page])], [14, 70, 6]);
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'screenshot_check_' + new Date().toISOString().slice(0, 10) + '.xlsx';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    msg(`Excel: не сошлось — ${bad.length}; проверить вручную — ${rev.length}; ошибок — ${err.length}`);
  };

  // ---------- диагностика: пагинация и данные распознавания сайта (копируется в буфер — пришлите разработчику) ----------
  $('#ts-diag').onclick = async () => {
    if (running) return;
    const L = analyze();
    if (!L) { msg('Таблица не найдена'); return; }
    running = true; busy = true;
    const pg = getPager();
    const nx = findNext();
    const info = {
      pager: { totals: getTotals(), box_found: !!pg, items: pg ? pg.items.map((k) => (k.innerText || '').trim().slice(0, 12)) : null, active: pg && pg.active ? (pg.active.innerText || '').trim() : null,
        next_found: !!nx, next_html: nx ? nx.outerHTML.slice(0, 250) : null, next_disabled: nx ? isDisabled(nx) : null },
      list_columns_index: L.idx, rows_in_list: L.length, rows_with_links: L.filter((r) => r.userLinks && r.userLinks.length).length,
      tickets: [],
    };
    for (const r of L.slice(0, 20)) {
      if (info.tickets.length >= 3) break;
      try {
        const j = await fetchHistoryRaw(r.id);
        const arr = findRows(j, ['dateEdit', 'socratResults'], 0) || [];
        if (!arr.some((o) => /receipt/.test(JSON.stringify(o.socratResults || '')))) continue;
        const entries = parseApiHistory(j);
        const sh = analyzeShots(entries);
        const o = arr.find((x) => x.socratResults && JSON.stringify(x.socratResults).length > 4) || arr[0];
        info.tickets.push({ ticket: r.id, links: r.userLinks, analysis: sh, site_suspect: siteSuspect({ shot: sh }), socrat_full: JSON.stringify(o.socratResults).slice(0, 3000) });
      } catch (e) { info.tickets.push({ ticket: r.id, error: e.message }); }
      await sleep(CFG.apiGapMs);
    }
    busy = false; running = false;
    const json = JSON.stringify(info, null, 1);
    console.log('[ts diag]', info);
    copy(json);
    const pre = $('#ts-pre');
    pre.style.display = 'block'; pre.textContent = json;
    msg('Результат диагностики скопирован в буфер — пришлите его.');
    run();
  };

  // ---------- тест распознавания: качаем картинку первого тикета со ссылкой в User files и показываем, что прочитано ----------
  $('#ts-ocrtest').onclick = async () => {
    if (running) return;
    const L = analyze();
    if (!L) { msg('Таблица не найдена'); return; }
    const row = L.find((r) => r.userLinks && r.userLinks.length);
    if (!row) { msg('В списке нет ни одной ссылки на файл пользователя (колонка User files)'); return; }
    running = true; busy = true;
    const info = { ticket: row.id, link: row.userLinks[0] };
    const t0 = Date.now();
    try {
      const url = new URL(row.userLinks[0], location.origin).href;
      const t1 = Date.now();
      const o = await recognizeUrl(url);
      info.engine = o.engine; info.bytes = o.bytes;
      info.seconds = Math.round((Date.now() - t1) / 100) / 10; info.confidence = Math.round(o.conf); info.readable = readable(o);
      info.text = o.text.replace(/\s+/g, ' ').trim().slice(0, 1500);
      info.standalone_numbers = sumCands(o.text, true);
    } catch (e) { info.error = e.message; info.fatal = !!e.fatal; }
    busy = false; running = false;
    const json = JSON.stringify(info, null, 1);
    console.log('[tn ocr test]', info);
    copy(json);
    const pre = $('#ts-pre');
    pre.style.display = 'block'; pre.textContent = json;
    msg(info.error ? 'Распознавание не сработало — результат скопирован в буфер, пришлите его.' : 'Готово. Результат скопирован в буфер — пришлите его.');
    run();
  };

  // ---------- автозапуск при изменении таблицы ----------
  let timer = null;
  new MutationObserver((muts) => {
    if (busy || running || muts.every((m) => panel.contains(m.target) || launcher.contains(m.target))) return;
    clearTimeout(timer);
    timer = setTimeout(run, 400);
  }).observe(document.body, { childList: true, subtree: true });

  syncUI();
  setTimeout(run, 800);
})();
