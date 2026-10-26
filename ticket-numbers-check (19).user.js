// ==UserScript==
// @name         Tickets: проверка номера (11 цифр) и файла агента
// @namespace    local.tickets.numbers
// @version      1.20
// @description  Ищет тикеты, где нет номера из 11 цифр (Transaction ID / Inside comment) или нет Agent's file — в самом списке и в истории тикета (фоновым запросом, без открытия окон); обход всех страниц, выгрузка в Excel
// @match        https://REPLACE_ME/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        GM_setClipboard
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        unsafeWindow
// @grant        GM_xmlhttpRequest
// @connect      *
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  const CFG = {
    ticketHeader: 'ticket id',               // колонка Ticket ID в списке
    historyHeader: 'ticket history',         // колонка с кнопкой Show (нужна только для запасного способа через окно)
    histMarker: 'date of ticket update',     // по этому заголовку узнаём таблицу истории
    // колонки (в списке и в истории ищутся по заголовку)
    txHeader: /^transaction id$/,
    insideHeader: /^inside comment$/,
    fileHeader: /^agent'?s? files?$/,
    // номер: ровно 11 цифр подряд (буквы рядом допустимы, цифры рядом — нет)
    // фраза в Inside comment (список + история) и «Sent for processing» N и более раз в истории — одна дополнительная проверка
    commentPhrase: 'Sir, please attach the approved transaction related to the payment from the ticket or set the right status.',
    sentStatus: 'Sent for processing (M)',
    sentMin: 3,
    extHeader: /^external status$/,
    numberRe: /(?<!\d)\d{11}(?!\d)/,
    historyApi: '/admin/backoffice/paymentsupporthistory',
    concurrency: 2,                  // сколько запросов истории идёт одновременно (сервер отвечает 529 при перегрузке)
    apiGapMs: 250,                   // пауза после каждого запроса в каждом потоке
    apiRetries: 5,                   // повторы при 429/5xx
    apiBackoffMs: 1500,              // первая пауза перед повтором, дальше ×2
    openWaitMs: 15000,               // запасной способ: сколько ждать появления истории после клика Show
    pauseMs: 400,                    // запасной способ: пауза между тикетами
    maxConsecutiveErrors: 3,         // столько ошибок подряд — останавливаемся
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
      XP.open = function (m, u) { this.__tnU = String(u); this.__tnH = {}; return o.apply(this, arguments); };
      XP.setRequestHeader = function (k, v) { if (this.__tnH) this.__tnH[k] = v; return sh.apply(this, arguments); };
      XP.send = function () {
        if (this.__tnU && this.__tnU.indexOf('paymentsupporthistory') >= 0) Object.assign(learned, this.__tnH);
        return sd.apply(this, arguments);
      };
    } catch (e) { /* перехват необязателен */ }
  })();

  const COLORS = { shot: '#ffadd2', rule: '#d3adf7', both: '#ff7875', num: '#ffd591', file: '#fff566', ok: '#d9f7be', error: '#d9d9d9' };
  const NUM = (t) => CFG.numberRe.test(t || '');

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
  const SKEY = 'tn_settings_v1';
  const DEFAULTS = { mode: 'both', chkNum: true, chkRule: true, shots: false, first: true, hide: false, badges: true, open: false, api: true };
  function loadSettings() {
    let raw = null;
    try { raw = typeof GM_getValue === 'function' ? GM_getValue(SKEY, null) : localStorage.getItem(SKEY); } catch (e) { /* ignore */ }
    try { return Object.assign({}, DEFAULTS, raw ? JSON.parse(raw) : {}); } catch (e) { return Object.assign({}, DEFAULTS); }
  }
  const S = loadSettings();
  S.shots = false;   // проверка скриншотов вынесена в отдельный плагин «Скрины»
  function saveSettings() {
    try {
      const raw = JSON.stringify(S);
      if (typeof GM_setValue === 'function') GM_setValue(SKEY, raw); else localStorage.setItem(SKEY, raw);
    } catch (e) { /* ignore */ }
  }

  // ---------- список жалоб ----------
  function hasLink(cell) {
    if (!cell) return false;
    if ([...cell.querySelectorAll('a[href]')].some((a) => !/^\s*javascript:/i.test(a.getAttribute('href')))) return true;
    if (cell.querySelector('img[src],video[src],source[src]')) return true;
    return /https?:\/\/\S+|\.(jpe?g|png|gif|pdf|docx?|xlsx?|zip|mp4|webp)\b/i.test(cell.textContent || '');
  }
  const phraseNorm = (t) => norm(String(t || '')).replace(/[.\s]+$/, '');
  const hasPhrase = (t) => phraseNorm(t).includes(phraseNorm(CFG.commentPhrase));
  const cellText = (c) => (c ? (c.textContent || '').replace(/\s+/g, ' ').trim() : '');

  // главная таблица — та, где есть Ticket ID (и обычно Ticket history; таблица истории — другая)
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
      hist: heads.findIndex((h) => h.includes(CFG.historyHeader)),
      tx: heads.findIndex((h) => CFG.txHeader.test(h)),
      inside: heads.findIndex((h) => CFG.insideHeader.test(h)),
      file: heads.findIndex((h) => CFG.fileHeader.test(h)),
      user: heads.findIndex((h) => /^user'?s? files?$/.test(h)),
    };
    let trs = [...table.querySelectorAll(':scope > tbody > tr')];
    if (!trs.length) trs = [...table.querySelectorAll('tbody tr')];
    const res = [];
    trs.forEach((tr) => {
      const cells = tr.children;
      if (cells.length <= Math.max(idx.ticket, idx.tx, idx.inside, idx.file, idx.hist, idx.user)) return;
      const id = (cells[idx.ticket].innerText || '').trim();
      if (!id) return;
      let btn = null;
      if (idx.hist >= 0) {
        const cell = cells[idx.hist];
        const els = [...cell.querySelectorAll('a,button,span,[role=button]')];
        btn = els.find((e) => /^(show|history|история)$/i.test(txt(e).trim())) || els[0] || (txt(cell).trim() ? cell : null);
      }
      const tx = idx.tx >= 0 ? cellText(cells[idx.tx]) : '';
      const inside = idx.inside >= 0 ? cellText(cells[idx.inside]) : '';
      res.push({
        tr, id, tCell: cells[idx.ticket], btn, tx, inside,
        numMain: NUM(tx) || NUM(inside),
        cmtMain: hasPhrase(inside),
        fileMain: idx.file >= 0 && hasLink(cells[idx.file]),
        userLinks: idx.user >= 0 ? [...cells[idx.user].querySelectorAll('a[href],img[src]')].slice(0, 3).map((e) => e.getAttribute('href') || e.getAttribute('src')) : null,
        fileHtml: idx.file >= 0 ? (cells[idx.file].innerHTML || '').slice(0, 200) : null,
      });
    });
    res.table = table; res.idx = idx;
    return res;
  }
  // истории запрашиваем только если в самом списке чего-то не хватает
  const needHist = () => true;   // статусы (Sent for processing) видны только в истории — запрашиваем её для каждого тикета

  // ---------- таблица истории ----------
  function isHistoryTable(t) {
    return !!t && [...t.querySelectorAll('th')].some((x) => norm(txt(x)).includes(CFG.histMarker)) &&
      ![...t.querySelectorAll('th')].some((x) => norm(txt(x)).includes(CFG.historyHeader));
  }
  function findHistoryTable(root, mainTable) {
    for (const th of root.querySelectorAll('th')) {
      if (!norm(txt(th)).includes(CFG.histMarker)) continue;
      const t = th.closest('table');
      if (!t || t === mainTable) continue;
      if (root === document && !visible(t)) continue;
      return t;
    }
    return null;
  }
  function historyFacts(table) {
    const hth = [...table.querySelectorAll('th')].find((x) => norm(txt(x)).includes(CFG.histMarker));
    if (!hth) return { entries: [], missing: [] };
    const heads = [...hth.parentElement.children].map((c) => norm(txt(c)));
    const iTx = heads.findIndex((h) => CFG.txHeader.test(h));
    const iIn = heads.findIndex((h) => CFG.insideHeader.test(h));
    const iFile = heads.findIndex((h) => CFG.fileHeader.test(h));
    const iExt = heads.findIndex((h) => CFG.extHeader.test(h));
    const missing = [];
    if (iTx < 0) missing.push('Transaction ID');
    if (iIn < 0) missing.push('Inside comment');
    if (iFile < 0) missing.push("Agent's file");
    if (iExt < 0) missing.push('External status');
    let trs = [...table.querySelectorAll('tbody tr')];
    if (!trs.length) trs = [...table.querySelectorAll('tr')].filter((tr) => tr !== hth.parentElement);
    const entries = trs.filter((tr) => !tr.querySelector('th') && tr.children.length > 1).map((tr) => ({
      tx: iTx >= 0 ? cellText(tr.children[iTx]) : '',
      inside: iIn >= 0 ? cellText(tr.children[iIn]) : '',
      file: iFile >= 0 && hasLink(tr.children[iFile]),
      ext: iExt >= 0 ? cellText(tr.children[iExt]) : '',
    }));
    return { entries, missing };
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
        } else add(parseFloat(tok));
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
    const fuzzyHas = (hay, needle) => {                              // допускаем одну неверно прочитанную цифру
      for (let a = 0; a + needle.length <= hay.length; a++) { let d = 0; for (let k = 0; k < needle.length && d <= 1; k++) if (hay[a + k] !== needle[k]) d++; if (d <= 1) return true; }
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
      if (w.length < 8 || fuzzyHas(alnum(dg), w.slice(-8))) out.walletFound = true;
    };
    for (const href of links.slice(0, 3)) {
      try {
        const canvas = await loadCanvas(new URL(href, location.origin).href);
        const p1 = await ocrCanvas(canvas, 'eng');                    // латинские цифры
        let any = readable(p1);
        scan(p1.text, any); out.text += ' ' + p1.text.replace(/\s+/g, ' ').trim().slice(0, 200);
        if (any) { out.ok = true; out.n++; } else out.err = out.err || 'текст на картинке не распознан (низкая уверенность)';
      } catch (e) { out.err = e.message; if (e.fatal) { out.fatal = true; break; } }
    }
    return out;
  }
  const siteSuspect = (r) => !!r.shot && (r.shot.sumRes === 'bad' || r.shot.walletRes === 'bad' || r.shot.walletRes === 'user');
  const ocrConfirmed = (r) => !!(S.shots && r.ocr && r.ocr.ok);
  // сумму отмечаем, только если картинка подтверждает ДРУГУЮ сумму (ту же, что назвал сайт), а суммы тикета на ней нет; если число просто не прочиталось — не отмечаем
  const sumBad = (r) => !r.error && !!r.shot && r.shot.sumRes === 'bad' && ocrConfirmed(r) && !r.ocr.sumFound && !!r.ocr.siteSeen;
  const walletBad = (r) => !r.error && !!r.shot && (r.shot.walletRes === 'bad' || r.shot.walletRes === 'user') && ocrConfirmed(r) && !r.ocr.walletFound && !!r.ocr.longNum;
  const isShot = (r) => sumBad(r) || walletBad(r);
  const SHOT_TXT = { bad: 'не совпала', user: 'совпала с кошельком ПОЛЬЗОВАТЕЛЯ', ok: 'совпала', unknown: 'не определить', none: 'нет чека' };

  const histMissing = new Set();   // колонки, которых не оказалось в истории
  function fillResult(res, entries, missing) {
    (missing || []).forEach((m) => histMissing.add(m));
    const numHist = entries.some((e) => NUM(e.tx) || NUM(e.inside));
    const fileHist = entries.some((e) => e.file);
    res.numSrc = res.numMain ? 'список' : numHist ? 'история' : '';
    res.fileSrc = res.fileMain ? 'список' : fileHist ? 'история' : '';
    res.cmtSrc = res.cmtMain ? 'список' : entries.some((e) => hasPhrase(e.inside)) ? 'история' : '';
    res.sent = entries.filter((e) => norm(e.ext) === norm(CFG.sentStatus)).length;
    res.shot = entries.some((e) => e.soc) ? analyzeShots(entries) : null;
    res.n = entries.length;
    res.error = false;
  }
  const newRes = (row) => ({ id: row.id, error: true, numMain: row.numMain, fileMain: row.fileMain, numSrc: '', fileSrc: '', cmtMain: row.cmtMain, cmtSrc: '', sent: 0, shot: null, n: 0, note: '', stuck: false, page: '', via: '' });

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
  // имена полей берём из словаря columns ответа: «поле → заголовок»
  function resolveKeys(cols) {
    const byLabel = (re) => Object.keys(cols).find((k) => re.test(norm(String(cols[k]))));
    const byKey = (re) => Object.keys(cols).find((k) => re.test(k));
    return {
      date: byLabel(/^date of ticket update$/) || 'dateEdit',
      tx: byLabel(CFG.txHeader) || byKey(/transaction/i) || '',
      inside: byLabel(CFG.insideHeader) || byKey(/inside/i) || '',
      ext: byLabel(CFG.extHeader) || byKey(/external.?status/i) || '',
      file: byLabel(CFG.fileHeader) || byKey(/agent.*file|file.*agent/i) || '',
    };
  }
  const valText = (v) => (v == null ? '' : typeof v === 'object' ? JSON.stringify(v) : String(v));
  // есть ли в значении файл: непустой массив/объект, либо строка со ссылкой / <a> / расширением файла
  function valHasFile(v) {
    if (v == null || v === '' || v === false) return false;
    if (Array.isArray(v)) return v.some((x) => valHasFile(x));
    if (typeof v === 'object') return Object.keys(v).length > 0 && Object.values(v).some((x) => valHasFile(x) || (typeof x === 'string' && x.trim() !== ''));
    const s = String(v).trim();
    if (!s || /^(-|null|n\/?a|none|undefined|unknown|0)$/i.test(s)) return false;
    return /https?:\/\/|<a\s|<img\s|href=|^\/|\.(jpe?g|png|gif|pdf|docx?|xlsx?|zip|mp4|webp)\b/i.test(s);
  }
  function parseApiHistory(j) {
    const cols = (j && (j.columns || (j.params && j.params.columns))) || null;
    if (!cols) throw new Error('в ответе нет описания колонок');
    const k = resolveKeys(cols);
    const found = [k.tx, k.inside, k.file].filter(Boolean);
    if (!found.length) throw new Error('в ответе нет ни одной из нужных колонок');
    const arr = findRows(j, [k.date, ...found, k.ext].filter(Boolean), 0);
    if (!arr) throw new Error('в ответе нет строк истории');
    const missing = [];
    if (!k.tx) missing.push('Transaction ID');
    if (!k.inside) missing.push('Inside comment');
    if (!k.file) missing.push("Agent's file");
    if (!k.ext) missing.push('External status');
    const entries = arr.map((o) => ({
      tx: k.tx ? valText(o[k.tx]) : '',
      inside: k.inside ? valText(o[k.inside]) : '',
      file: k.file ? valHasFile(o[k.file]) : false,
      ext: k.ext ? valText(o[k.ext]).replace(/\s+/g, ' ').trim() : '',
      date: String(o[k.date] == null ? '' : o[k.date]),
      soc: parseSoc(o.socratResults),
      amount: o.amount, agentW: o.walletAccountNumber, userW: o.numberWallet,
    }));
    return { entries, missing };
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

  // ---------- закрытие окна истории ----------
  function modalRoot(el) {
    const m = el.closest('[role=dialog],[class*=modal],[class*=popup],[class*=dialog],[class*=overlay],[class*=drawer]');
    if (m && m !== document.body && m !== document.documentElement) return m;
    let r = el;
    for (let i = 0; i < 6 && r.parentElement && r.parentElement !== document.body; i++) r = r.parentElement;
    return r;
  }
  function findCloseBtn(root) {
    const isClose = (e) => {
      const t = (txt(e) || '').trim();
      const c = [e.getAttribute('class'), e.getAttribute('aria-label'), e.getAttribute('title')].filter(Boolean).join(' ');
      return /^(×|✕|✖|x|close|закрыть)$/i.test(t) || /(^|[\s_-])close([\s_-]|$)/i.test(c);
    };
    return [...root.querySelectorAll('button,a,span,i,div,svg,[role=button]')].filter((e) => visible(e) && !e.closest('table')).find(isClose) || null;
  }
  async function closeHistory(table, btn) {
    const gone = () => !document.contains(table) || !visible(table);
    if (gone()) return true;
    const x = findCloseBtn(modalRoot(table));            // 1) крестик окна
    if (x) { clickEl(x); if (await waitFor(gone, 3000)) return true; }
    ['keydown', 'keyup'].forEach((t) =>                  // 2) Esc
      document.dispatchEvent(new KeyboardEvent(t, { key: 'Escape', keyCode: 27, which: 27, bubbles: true })));
    if (await waitFor(gone, 2000)) return true;
    if (btn) { clickEl(btn); if (await waitFor(gone, 3000)) return true; }   // 3) история раскрывается на месте — повторный клик по Show
    return false;
  }

  async function fetchDoc(url) {
    try {
      const r = await fetch(url, { credentials: 'include' });
      if (!r.ok) return null;
      return new DOMParser().parseFromString(await r.text(), 'text/html');
    } catch (e) { return null; }
  }

  // ---------- проверка одного тикета ----------
  // res.error — не удалось получить историю
  async function checkRow(row, mainTable) {
    const res = newRes(row);
    if (!row.btn) { res.note = 'нет кнопки Show в колонке Ticket history'; return res; }
    const calls = [];
    const origOpen = PW.open;
    PW.open = function (u) { calls.push(String(u)); return null; };   // если история открывается в новой вкладке — берём адрес и читаем сами
    let table = null, fromDoc = false;
    try {
      const stale = findHistoryTable(document, mainTable);               // остатки истории от прошлого тикета
      if (stale) await closeHistory(stale, row.btn);

      const a = row.btn.closest('a[href]');
      const href = a && a.getAttribute('href');
      const real = href && !/^(#|javascript:)/i.test(href);
      if (real) {                                                         // у Show настоящий адрес: пробуем прочитать страницу напрямую
        const doc = await fetchDoc(a.href);
        const t = doc && findHistoryTable(doc, null);
        if (t) { table = t; fromDoc = true; }
      }
      if (!table) {
        const urlBefore = location.href;
        clickEl(row.btn);
        const got = await waitFor(() => findHistoryTable(document, mainTable) || (calls.length ? 'OPEN' : null), CFG.openWaitMs);
        if (got === 'OPEN') {
          const doc = await fetchDoc(calls[0]);
          const t = doc && findHistoryTable(doc, null);
          if (t) { table = t; fromDoc = true; } else { res.note = 'история открылась в новой вкладке, но таблицу из неё прочитать не удалось: ' + calls[0]; return res; }
        } else if (got) { table = got; }
        else if (location.href !== urlBefore || !analyze()) {
          res.note = 'история открывается отдельной страницей — запустите «Диагностика»'; res.stuck = true;
          try { history.back(); } catch (e) { /* ignore */ }
          await waitFor(() => analyze(), 8000);
          return res;
        } else { res.note = 'история не открылась за ' + CFG.openWaitMs / 1000 + ' сек'; return res; }
      }
      if (!fromDoc) {                                                     // дождаться, пока строки истории догрузятся
        let prev = -1, same = 0;
        const t0 = Date.now();
        while (same < 2 && Date.now() - t0 < 8000) {
          await sleep(300);
          const n = historyFacts(table).entries.length;
          if (n > 0 && n === prev) same++; else { same = 0; prev = n; }
        }
      }
      const hf = historyFacts(table);
      if (!hf.entries.length) { res.note = 'в таблице истории нет строк'; }
      else fillResult(res, hf.entries, hf.missing);
      if (!fromDoc && !(await closeHistory(table, row.btn))) { res.stuck = true; res.note = (res.note ? res.note + '; ' : '') + 'окно истории не закрылось'; }
      return res;
    } finally {
      PW.open = origOpen;
    }
  }

  // ---------- меню ----------
  const style = document.createElement('style');
  style.textContent = `
    #tn-panel, #tn-panel * { box-sizing: border-box; }
    #tn-panel .tn-sec { font-weight: 700; margin: 8px 0 4px; color: #23909f; font-size: 11px; text-transform: uppercase; letter-spacing: .4px; }
    #tn-panel label { display: block; margin-bottom: 4px; cursor: pointer; }
    #tn-panel button { width: 100%; margin-bottom: 4px; }
    #tn-panel .tn-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 4px; margin-bottom: 4px; }
    #tn-panel .tn-grid button { margin: 0; padding: 3px 2px; font-size: 12px; }
    #tn-panel pre { max-height: 160px; overflow: auto; font-size: 11px; white-space: pre-wrap; background: #f5f5f5; padding: 4px; margin: 4px 0 0; }
    td[data-tnb]::after { content: attr(data-tnb); display: block; margin-top: 2px; font: 600 10px/1.2 Arial, sans-serif; color: inherit; white-space: normal; }
  `;
  document.head.appendChild(style);

  const launcher = document.createElement('button');
  launcher.type = 'button';
  launcher.title = 'Проверка номера и файла агента';
  launcher.style.cssText =
    'position:fixed;right:100px;bottom:12px;z-index:999999;background:#08979c;color:#fff;border:none;border-radius:20px;' +
    'box-sizing:border-box;width:80px;padding:8px 0;text-align:center;white-space:nowrap;font:700 13px Arial,sans-serif;cursor:pointer;box-shadow:0 2px 10px rgba(0,0,0,.35)';
  launcher.textContent = 'Номера';
  document.body.appendChild(launcher);

  const panel = document.createElement('div');
  panel.id = 'tn-panel';
  panel.style.cssText =
    'position:fixed;right:100px;bottom:56px;z-index:999999;background:#fff;border:2px solid #08979c;border-radius:8px;' +
    'padding:10px;font:13px/1.4 Arial,sans-serif;width:300px;max-height:calc(100vh - 80px);overflow:auto;' +
    'box-shadow:0 4px 16px rgba(0,0,0,.3);color:#222;display:none';
  panel.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center"><b>Номер (11 цифр) и файл агента</b><span id="tn-x" style="cursor:pointer;font-size:16px;padding:0 4px" title="Закрыть меню">✕</span></div>
    <div id="tn-sets">
      <div class="tn-sec">Какие проверки включить</div>
      <label><input type="checkbox" id="tn-chknum"> Номер (11 цифр) и файл агента</label>
      <label><input type="checkbox" id="tn-chkrule"> Комментарий «attach the approved transaction…» и Sent for processing ×3+</label>
      <div style="font-size:11px;color:#666;margin-bottom:4px">Выключенная проверка не подсвечивается, не копируется и не попадает в Excel. Историю тикетов плагин читает в любом случае.</div>
      <div class="tn-sec">Номер и файл: что считать проблемой</div>
      <select id="tn-mode" style="width:100%;margin-bottom:4px">
        <option value="both">Нет И номера, И файла агента</option>
        <option value="any">Нет номера ИЛИ нет файла агента</option>
      </select>
      <div style="font-size:11px;color:#666;margin-bottom:4px">Номер — в Transaction ID и Inside comment, файл — в Agent's file; ищем и в списке, и во всей истории. Цвета: красный — нет номера и файла, оранжевый — нет номера, жёлтый — нет файла, фиолетовый — комментарий / Sent.</div>
      <label><input type="checkbox" id="tn-api"> Фоновый режим (без открытия окон)</label>
      <div class="tn-sec">Отображение</div>
      <label><input type="checkbox" id="tn-hide"> Скрыть тикеты без проблем</label>
      <label><input type="checkbox" id="tn-badges"> Подписи в ячейке</label>
    </div>
    <div class="tn-sec">Результат</div>
    <div id="tn-stat" style="margin-bottom:6px">—</div>
    <button id="tn-run">Проверить страницу</button>
    <button id="tn-retry">Повторить ошибки</button>
    <div class="tn-sec">Копировать ID</div>
    <div class="tn-grid"><button id="tn-c-both">Нет обоих</button><button id="tn-c-num">Нет номера</button><button id="tn-c-file">Нет файла</button><button id="tn-c-rule">Комм./Sent</button><button id="tn-c-error">Ошибки</button></div>
    <div class="tn-sec">Все страницы</div>
    <label><input type="checkbox" id="tn-first"> Начать с 1-й страницы</label>
    <button id="tn-crawl">Все страницы → проверить</button>
    <button id="tn-resume" disabled>Продолжить с текущей</button>
    <button id="tn-stop" disabled>Стоп</button>
    <button id="tn-xlsx" disabled>Скачать Excel</button>
    <button id="tn-diag">Диагностика (1 тикет)</button>
    <div id="tn-msg" style="font-size:12px;color:#555"></div>
    <pre id="tn-pre" style="display:none"></pre>`;
  document.body.appendChild(panel);
  const $ = (q) => panel.querySelector(q);
  const msg = (t) => ($('#tn-msg').textContent = t);
  const setStat = (t) => ($('#tn-stat').textContent = t);

  const showPanel = (v) => { S.open = v; panel.style.display = v ? 'block' : 'none'; saveSettings(); };
  launcher.onclick = () => showPanel(panel.style.display === 'none');
  $('#tn-x').onclick = () => showPanel(false);
  showPanel(!!S.open);

  function syncUI() {
    $('#tn-mode').value = S.mode; $('#tn-chknum').checked = !!S.chkNum; $('#tn-chkrule').checked = !!S.chkRule; $('#tn-api').checked = !!S.api;
    $('#tn-hide').checked = !!S.hide; $('#tn-badges').checked = !!S.badges; $('#tn-first').checked = !!S.first;
  }
  const changed = () => { saveSettings(); run(); };
  $('#tn-chknum').onchange = (e) => { S.chkNum = e.target.checked; changed(); };
  $('#tn-chkrule').onchange = (e) => { S.chkRule = e.target.checked; changed(); };
  $('#tn-mode').onchange = (e) => { S.mode = e.target.value; changed(); };
  $('#tn-api').onchange = (e) => { S.api = e.target.checked; if (S.api) { apiOn = true; apiRetried = false; } saveSettings(); };
  $('#tn-hide').onchange = (e) => { S.hide = e.target.checked; changed(); };
  $('#tn-badges').onchange = (e) => { S.badges = e.target.checked; changed(); };
  $('#tn-first').onchange = (e) => { S.first = e.target.checked; saveSettings(); };

  const results = new Map();   // id -> результат проверки
  let busy = false;            // идёт проверка (не перерисовываем список в это время)
  let running = false;
  let aborted = false;
  let curPageNo = 0;

  // итог по тикету считается на лету — переключатели «что проверять» действуют сразу, без повторной проверки
  const noNum = (r) => !r.numSrc;
  const noFile = (r) => !r.fileSrc;
  // cat — что именно не найдено (не зависит от режима); st — как тикет показываем с учётом режима
  const cat = (r) => (r.error ? 'error' : !S.chkNum ? 'ok' : noNum(r) && noFile(r) ? 'both' : noNum(r) ? 'num' : noFile(r) ? 'file' : 'ok');
  const st = (r) => { const c = cat(r); return S.mode === 'both' && (c === 'num' || c === 'file') ? 'ok' : c; };
  const isRule = (r) => S.chkRule && !r.error && (!!r.cmtSrc || r.sent >= CFG.sentMin);
  const LABEL = { both: 'нет номера и файла агента', num: 'нет номера (11 цифр)', file: 'нет файла агента' };

  function paint(rows) {
    rows.forEach((r) => {
      const res = results.get(r.id);
      let color = '', badge = '', title = '';
      if (res) {
        const s = st(res);
        color = COLORS[s];
        title = `Номер: ${res.numSrc ? 'есть (' + res.numSrc + ')' : 'не найден'}; файл агента: ${res.fileSrc ? 'есть (' + res.fileSrc + ')' : 'не найден'}; записей в истории: ${res.n}`;
        const shotTxt = isShot(res) ? [sumBad(res) ? 'суммы тикета нет на скрине' : '', walletBad(res) ? 'кошелька агента нет на скрине' : ''].filter(Boolean) : [];
        if (res.ocr) title += '\nРаспознавание картинки: ' + (res.ocr.ok ? `сумма тикета ${res.ocr.sumFound ? 'найдена' : 'не найдена'}${res.ocr.siteSeen ? ' (на картинке видна сумма сайта)' : ''}, кошелёк агента ${res.ocr.walletFound ? 'найден' : 'не найден'}${res.ocr.longNum ? '' : ' (длинных номеров не прочитано)'}` : 'не удалось (' + res.ocr.err + ')');
        if (res.shot && res.shot.m) title += `\nСкрины: чеков ${res.shot.m}; сумма тикета ${res.shot.amount}, на чеке ${res.shot.sums.join(' / ') || '—'} (${SHOT_TXT[res.shot.sumRes]}); кошелёк агента ${res.shot.agentW || '—'}, получатель на чеке ${res.shot.recv.join(' / ') || '—'} (${SHOT_TXT[res.shot.walletRes]})`;
        const extra = [res.cmtSrc ? 'есть комментарий «attach the approved transaction»' : '', res.sent >= CFG.sentMin ? `Sent for processing ×${res.sent}` : ''].filter(Boolean);
        if (extra.length) title += '\n' + extra.join('; ');
        if (s === 'error') { badge = 'ошибка'; title = res.note; }
        else if (s !== 'ok') badge = [LABEL[s], ...extra.map((x) => (x.startsWith('Sent') ? x : 'комментарий')), ...shotTxt].join(' · ');
        else if (shotTxt.length) { color = COLORS.shot; badge = [...extra.map((x) => (x.startsWith('Sent') ? x : 'комментарий')), ...shotTxt].join(' · '); }
        else if (extra.length) { color = COLORS.rule; badge = extra.map((x) => (x.startsWith('Sent') ? x : 'комментарий')).join(' · '); }
      }
      r.tCell.style.background = color;
      r.tCell.style.fontWeight = color && res && st(res) !== 'ok' ? '700' : '';
      r.tCell.title = title;
      if (S.badges && badge) r.tCell.setAttribute('data-tnb', badge); else r.tCell.removeAttribute('data-tnb');
      r.tr.style.display = S.hide && res && st(res) === 'ok' && !isRule(res) && !isShot(res) ? 'none' : '';
    });
  }
  function counts() {
    const c = [...results.values()].map(cat), v = [...results.values()].map(st);
    return {
      num: c.filter((x) => x === 'num' || x === 'both').length,
      file: c.filter((x) => x === 'file' || x === 'both').length,
      both: c.filter((x) => x === 'both').length,
      bad: v.filter((x) => x === 'num' || x === 'file' || x === 'both').length,
      rule: [...results.values()].filter(isRule).length,
      shot: [...results.values()].filter(isShot).length,
      ok: v.filter((x) => x === 'ok').length,
      error: v.filter((x) => x === 'error').length,
    };
  }
  function run() {
    const rows = analyze();
    if (!rows) { setStat('Таблица не найдена'); return; }
    paint(rows);
    const c = counts();
    const warn = [];
    if (S.chkNum && rows.idx.tx < 0 && rows.idx.inside < 0) warn.push('Transaction ID / Inside comment');
    if (S.chkNum && rows.idx.file < 0) warn.push("Agent's file");
    if (S.shots && rows.idx.user < 0) warn.push('User files (нужна для скриншотов)');
    const w = warn.length ? `<div style="color:#d4380d;font-size:12px">⚠ Нет колонки: ${warn.join(', ')} (включите её в Column settings)</div>` : '';
    const hm = histMissing.size ? `<div style="color:#d4380d;font-size:12px">⚠ В истории не найдены колонки: ${[...histMissing].join(', ')}</div>` : '';
    $('#tn-stat').innerHTML = `Строк: <b>${rows.length}</b> · к проверке: <b>${rows.filter((r) => !results.has(r.id)).length}</b><br>` +
      `всего проверено: без проблем <b style="color:#389e0d">${c.ok}</b> · проблемных <b style="color:#cf1322">${c.bad}</b> (без номера ${c.num}, без файла ${c.file}, без обоих ${c.both}) · комментарий / Sent×${CFG.sentMin}+ <b style="color:#531dab">${c.rule}</b> · скрин не сошёлся <b style="color:#c41d7f">${c.shot}</b> · ошибок <b>${c.error}</b>${w}${hm}`;
  }

  // тикеты, у которых всё есть уже в списке, в историю не ходим
  function prepare(L) {
    L.forEach((r) => {
      if (results.has(r.id) || needHist(r)) return;
      const res = newRes(r);
      res.error = false; res.numSrc = 'список'; res.fileSrc = 'список'; res.via = 'список'; res.page = curPageNo || '';
      results.set(r.id, res);
    });
  }

  let apiOn = true;          // фоновый режим доступен (выключается, если запросы не работают)
  let apiRetried = false;
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
        try {
          const h = await fetchHistoryApi(row.id);
          const res = newRes(row);
          fillResult(res, h.entries, h.missing);
          res.page = curPageNo || ''; res.via = 'фон';
          results.set(row.id, res);
          ok++;
        } catch (e) { fail++; if (!firstErr) firstErr = e.message; }
        setStat(`Фоновая проверка: ${ok + fail} из ${rows.length}…`);
        await sleep(CFG.apiGapMs);
      }
    };
    await Promise.all(Array.from({ length: CFG.concurrency }, worker));
    busy = false;
    run();
    if (ok === 0 && fail > 0) { apiOn = false; msg(`Фоновый запрос не сработал (${firstErr}) — проверяю через окно истории.`); }
    else if (fail > 0) msg(`Фоном проверено ${ok}, не удалось ${fail} (${firstErr}) — остальные через окно истории.`);
  }

  async function ocrPass() {
    const L = analyze();
    if (!L) return;
    const rows = L.filter((r) => { const x = results.get(r.id); return x && !x.error && !x.ocr && siteSuspect(x); });
    if (!rows.length) return;
    busy = true;
    let n = 0, okc = 0, errMsg = '';
    for (const r of rows) {
      if (aborted) break;
      n++;
      setStat(`Распознаю скриншоты: ${n} из ${rows.length}…`);
      const res = results.get(r.id);
      if (!r.userLinks || !r.userLinks.length) { res.ocr = { ok: false, err: 'в списке нет ссылок на файлы пользователя' }; continue; }
      res.ocr = await ocrCheckRow(res, r.userLinks);
      if (res.ocr.ok) okc++; else errMsg = res.ocr.err;
      if (res.ocr.fatal) break;
    }
    busy = false;
    run();
    if (!okc && rows.length) msg(`Распознавание картинок не сработало: ${errMsg}. Скрины не отмечены.`);
  }

  // проверка всех непроверенных тикетов текущей страницы; 'ok' | 'errors' | 'aborted'
  async function checkPage() {
    let consec = 0;
    const L0 = analyze();
    if (!L0) return 'ok';
    prepare(L0);
    run();
    if (!apiOn && !apiRetried && Object.keys(learned).length) { apiOn = true; apiRetried = true; }
    if (S.api && apiOn) { await apiPass(); if (aborted) return 'aborted'; }
    if (S.shots) { await ocrPass(); if (aborted) return 'aborted'; }
    for (;;) {
      if (aborted) return 'aborted';
      const L = analyze();
      if (!L) return 'ok';
      const row = L.find((r) => !results.has(r.id));
      if (!row) return 'ok';
      busy = true;
      setStat(`Открываю историю: ${row.id}…`);
      let res;
      try { res = await checkRow(row, L.table); } catch (e) { res = newRes(row); res.note = 'ошибка: ' + e.message; }
      busy = false;
      res.page = curPageNo || ''; res.via = res.via || 'окно';
      results.set(row.id, res);
      run();
      if (res.error) { consec++; if (consec >= CFG.maxConsecutiveErrors || res.stuck) return 'errors'; } else consec = 0;
      await sleep(CFG.pauseMs);
    }
  }

  function setRunning(v) {
    running = v;
    $('#tn-run').disabled = v; $('#tn-retry').disabled = v; $('#tn-diag').disabled = v;
    $('#tn-crawl').disabled = v;
    $('#tn-resume').disabled = v || pagesDone === 0;
    $('#tn-stop').disabled = !v;
    $('#tn-xlsx').disabled = v || results.size === 0;
    panel.querySelectorAll('#tn-sets input').forEach((el) => { el.disabled = v; });
  }

  $('#tn-run').onclick = async () => {
    if (running) return;
    aborted = false; curPageNo = 0; setRunning(true);
    const r = await checkPage();
    run(); setRunning(false);
    msg(r === 'errors' ? 'Остановлено из-за ошибок: история не читается. Нажмите «Диагностика (1 тикет)».' : r === 'aborted' ? 'Остановлено.' : 'Страница проверена.');
  };
  $('#tn-retry').onclick = () => {
    [...results.entries()].forEach(([id, r]) => { if (r.error) results.delete(id); });
    run(); $('#tn-run').onclick();
  };

  function copy(text) {
    if (typeof GM_setClipboard === 'function') GM_setClipboard(text);
    else navigator.clipboard.writeText(text);
  }
  const PICK = {
    num: (r) => cat(r) === 'num' || cat(r) === 'both',
    file: (r) => cat(r) === 'file' || cat(r) === 'both',
    both: (r) => cat(r) === 'both',
    rule: isRule,
    shot: isShot,
    bad: (r) => ['num', 'file', 'both'].includes(st(r)),
    error: (r) => st(r) === 'error',
  };
  const LBL = { shot: 'скрин не сошёлся с суммой/кошельком', rule: 'комментарий / Sent for processing ×3+', both: 'нет ни номера, ни файла', num: 'нет номера (в т.ч. и файла)', file: 'нет файла агента (в т.ч. и номера)', bad: 'проблемные', error: 'ошибки' };
  ['both', 'num', 'file', 'rule', 'error'].forEach((k) => {
    $('#tn-c-' + k).onclick = () => {
      const ids = [...results.values()].filter(PICK[k]).map((r) => r.id);
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
      if (panel.contains(el) || el.closest('#tn-panel,#th-panel,#ts-panel') || el.closest('table')) return false;
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
      if (panel.contains(el) || el.closest('#tn-panel,#th-panel,#ts-panel') || el.closest('table') || !visible(el)) return false;
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
      console.error('[tn]', e);
    }
    busy = false;
    run();
    const c = counts();
    msg(`${note}Конец: ${reason || '—'}. Страниц: ${pagesDone}. Нет ни номера, ни файла: ${c.both}; комментарий / Sent×${CFG.sentMin}+: ${c.rule}; скрин не сошёлся: ${c.shot}; (всего без номера: ${c.num}, без файла: ${c.file}); ошибок: ${c.error}.${coverage()}`);
    setRunning(false);
  }

  $('#tn-resume').onclick = () => allPages(true);
  $('#tn-crawl').onclick = () => allPages(false);
  $('#tn-stop').onclick = () => { aborted = true; };

  // Excel: «Все проблемные» / «Нет номера» / «Нет файла агента» / «Ошибки»
  $('#tn-xlsx').onclick = () => {
    if (typeof XLSX === 'undefined') { msg('Библиотека XLSX не загрузилась.'); return; }
    const numId = (id) => (/^\d{1,15}$/.test(id) ? Number(id) : id);
    const all = [...results.values()];
    const wb = XLSX.utils.book_new();
    const add = (name, aoa, wch) => {
      const sh = XLSX.utils.aoa_to_sheet(aoa);
      sh['!cols'] = wch.map((w) => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, sh, name);
    };
    const head = ['ticket_id', 'номер (11 цифр)', 'файл агента', 'записей в истории', 'проверено', 'page'];
    const yn = (b) => (b ? 'да' : 'нет');
    const src = (v) => v || 'НЕТ';
    const line = (r) => [numId(r.id), src(r.numSrc), src(r.fileSrc), r.n, r.via, r.page];
    const W = [14, 16, 16, 16, 12, 6];
    const bo = all.filter((r) => cat(r) === 'both'), on = all.filter((r) => cat(r) === 'num'), of = all.filter((r) => cat(r) === 'file'), err = all.filter(PICK.error);
    add('Нет номера и файла', [head, ...bo.map(line)], W);
    add('Только нет номера', [head, ...on.map(line)], W);
    add('Только нет файла агента', [head, ...of.map(line)], W);
    const ru = all.filter(isRule);
    add('Комментарий и Sent', [['ticket_id', 'комментарий в Inside comment', `Sent for processing (раз)`, `Sent ≥ ${CFG.sentMin}`, 'нет номера и файла', 'записей в истории', 'page'],
      ...ru.map((r) => [numId(r.id), r.cmtSrc || 'нет', r.sent, yn(r.sent >= CFG.sentMin), yn(cat(r) === 'both'), r.n, r.page])], [14, 26, 22, 12, 18, 16, 6]);
    const sh = all.filter(isShot);
    if (S.shots) add('Скриншот не сошёлся', [['ticket_id', 'сумма тикета', 'сумма на чеке', 'сумма', "кошелёк агента (тикет)", 'получатель на чеке', 'кошелёк', 'чеков распознано', 'скринов всего', 'причина (распознавание сайта)', 'картинка: сумма найдена', 'картинка: кошелёк агента найден', 'текст с картинки (фрагмент)', 'page'],
      ...sh.map((r) => [numId(r.id), r.shot.amount, r.shot.sums.join(' / '), SHOT_TXT[r.shot.sumRes], r.shot.agentW, r.shot.recv.join(' / '), SHOT_TXT[r.shot.walletRes], r.shot.m, r.shot.n, r.shot.reasons, r.ocr && r.ocr.ok ? yn(r.ocr.sumFound) : '—', r.ocr && r.ocr.ok ? yn(r.ocr.walletFound) : '—', r.ocr ? (r.ocr.text || r.ocr.err || '').slice(0, 300) : '', r.page])], [14, 14, 16, 14, 26, 24, 30, 12, 12, 40, 14, 18, 60, 6]);
    add('Ошибки', [['ticket_id', 'причина', 'page'], ...err.map((r) => [numId(r.id), r.note, r.page])], [14, 70, 6]);
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'ticket_numbers_check_' + new Date().toISOString().slice(0, 10) + '.xlsx';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    msg(`Excel: нет номера и файла — ${bo.length}; только нет номера — ${on.length}; только нет файла — ${of.length}; комментарий/Sent — ${ru.length}; скрин не сошёлся — ${sh.length}; ошибок — ${err.length}`);
  };

  // ---------- тест распознавания: качаем картинку первого тикета со ссылкой в User files и показываем, что прочитано ----------
  ($('#tn-ocrtest') || {}).onclick = async () => {
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
      const blob = await gmBlob(url);
      info.downloaded = { type: blob && blob.type, bytes: blob && blob.size };
      const canvas = await prepImage(blob);
      for (const lang of ['eng']) {
        try {
          const t1 = Date.now();
          const o = await ocrCanvas(canvas, lang);
          info[lang] = { seconds: Math.round((Date.now() - t1) / 100) / 10, confidence: Math.round(o.conf), readable: readable(o),
            text: o.text.replace(/\s+/g, ' ').trim().slice(0, 1200), standalone_numbers: sumCands(o.text, true) };
        } catch (e) { info[lang] = { error: e.message }; }
      }
    } catch (e) { info.error = e.message; info.fatal = !!e.fatal; }
    busy = false; running = false;
    const json = JSON.stringify(info, null, 1);
    console.log('[tn ocr test]', info);
    copy(json);
    const pre = $('#tn-pre');
    pre.style.display = 'block'; pre.textContent = json;
    msg(info.error ? 'Распознавание не сработало — результат скопирован в буфер, пришлите его.' : 'Готово. Результат скопирован в буфер — пришлите его.');
    run();
  };

  // ---------- диагностика: что приходит в истории (копируется в буфер — пришлите разработчику) ----------
  $('#tn-diag').onclick = async () => {
    if (running) return;
    const L = analyze();
    if (!L) { msg('Таблица не найдена'); return; }
    const row = L.find((r) => needHist(r)) || L[0];
    if (!row) { msg('Нет строк'); return; }
    running = true; busy = true;
    const pg = getPager();
    const nx = findNext();
    const info = {
      pager: { totals: getTotals(), box_found: !!pg, items: pg ? pg.items.map((k) => (k.innerText || '').trim().slice(0, 12)) : null, active: pg && pg.active ? (pg.active.innerText || '').trim() : null,
        next_found: !!nx, next_html: nx ? nx.outerHTML.slice(0, 250) : null, next_disabled: nx ? isDisabled(nx) : null, box_html: pg ? pg.box.outerHTML.slice(0, 600) : null },
      ticket: row.id, list_columns_index: L.idx,
      list_values: { transaction_id: row.tx.slice(0, 120), inside_comment: row.inside.slice(0, 200), agent_file_cell_html: row.fileHtml, number_in_list: row.numMain, file_in_list: row.fileMain },
    };
    try {
      const j = await fetchHistoryRaw(row.id);
      const cols = (j && (j.columns || (j.params && j.params.columns))) || {};
      info.api_columns = cols;
      const k = resolveKeys(cols);
      info.api_keys_used = k;
      const arr = findRows(j, [k.date, k.tx, k.inside, k.file, k.ext].filter(Boolean), 0) || [];
      info.api_rows = arr.length;
      info.api_first_row_keys = arr[0] ? Object.keys(arr[0]) : [];
      info.api_sample = arr.slice(0, 3).map((o) => {
        const x = {};
        [k.date, k.tx, k.inside, k.file, k.ext].filter(Boolean).forEach((key) => { x[key] = JSON.stringify(o[key] === undefined ? null : o[key]).slice(0, 300); });
        return x;
      });
      try {
        const h = parseApiHistory(j);
        info.parsed = { sent_for_processing_rows: h.entries.filter((e) => norm(e.ext) === norm(CFG.sentStatus)).length, comment_in_history: h.entries.some((e) => hasPhrase(e.inside)), missing_columns: h.missing, number_in_history: h.entries.some((e) => NUM(e.tx) || NUM(e.inside)), file_in_history: h.entries.some((e) => e.file) };
      } catch (e) { info.parse_error = e.message; }
    } catch (e) { info.api_error = e.message; }
    // распознавание документов: что лежит в socratResults и соседних полях (по первым тикетам страницы)
    info.recognition = [];
    for (const r of L.slice(0, 20)) {
      if (info.recognition.length >= 3) break;
      try {
        const j = await fetchHistoryRaw(r.id);
        const cols2 = (j && (j.columns || (j.params && j.params.columns))) || {};
        const arr2 = findRows(j, ['dateEdit', 'nameExternalStatus'], 0) || [];
        const pick = (o, key) => { const v = o[key]; return v === undefined ? undefined : JSON.stringify(v).slice(0, 700); };
        const withData = arr2.find((o) => o.socratResults && JSON.stringify(o.socratResults).length > 4) || arr2[0] || {};
        if (!arr2.some((o) => /receipt/.test(JSON.stringify(o.socratResults || '')))) continue;   // нужны только тикеты, где сайт увидел чек
        info.recognition.push({
          user_files_links_in_list: r.userLinks,
          ticket: r.id, rows: arr2.length, has_socrat_in_any_row: arr2.some((o) => o.socratResults && JSON.stringify(o.socratResults).length > 4),
          socratResults: pick(withData, 'socratResults'), documentsUser: pick(withData, 'documentsUser'), documentsSupport: pick(withData, 'documentsSupport'),
          documentsExternal: pick(withData, 'documentsExternal'), jsonField: pick(withData, 'jsonField'), dopParamMk: pick(withData, 'dopParamMk'),
          amount: pick(withData, 'amount'), walletAccountNumber: pick(withData, 'walletAccountNumber'), numberWallet: pick(withData, 'numberWallet'),
          socrat_full: JSON.stringify(withData.socratResults == null ? null : withData.socratResults).slice(0, 4500),
          top_level_keys: Object.keys(j || {}), params_keys: j && j.params ? Object.keys(j.params) : null,
        });
      } catch (e) { info.recognition.push({ ticket: r.id, error: e.message }); }
      await sleep(CFG.apiGapMs);
    }
    busy = false; running = false;
    const json = JSON.stringify(info, null, 1);
    console.log('[tn diag]', info);
    copy(json);
    const pre = $('#tn-pre');
    pre.style.display = 'block'; pre.textContent = json;
    msg('Результат диагностики скопирован в буфер — пришлите его.');
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
