// ==UserScript==
// @name         Tickets: проверка истории статусов (Received (M))
// @namespace    local.tickets.history
// @version      1.5
// @description  Для тикетов в выбранных статусах (Received (Fraud) (M), Create new transaction (M)) проверяет, был ли в истории статус Received (M): фоновым запросом (без открытия окон) или через окно Show; обход всех страниц, выгрузка в Excel
// @match        https://REPLACE_ME/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        GM_setClipboard
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        unsafeWindow
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  const CFG = {
    ticketHeader: 'ticket id',               // колонка Ticket ID в списке
    statusHeader: /^external status$/,       // колонка ТЕКУЩЕГО статуса в списке
    historyHeader: 'ticket history',         // колонка с кнопкой Show
    histMarker: 'date of ticket update',     // по этому заголовку узнаём таблицу истории
    histStatusHeader: /^external status$/,   // колонка статуса в таблице истории
    // текущие статусы тикетов, которые проверяем (можно отключать в меню), и искомый статус в истории
    currentStatuses: ['Received (Fraud) (M)', 'Create new transaction (M)'],
    targetStatus: 'Received (M)',
    // если Received (M) в истории нет, но есть любой из этих статусов — тикет идёт на отдельный лист
    altStatuses: ['Approved (M)', 'Credited to another account (M)'],
    // фоновый режим: тот же запрос, который сайт делает при нажатии Show (POST, JSON); при сбое автоматически работает проверка через окно
    historyApi: '/admin/backoffice/paymentsupporthistory',
    concurrency: 2,                  // сколько запросов истории идёт одновременно (сервер отвечает 529 при перегрузке)
    apiGapMs: 250,                   // пауза после каждого запроса в каждом потоке
    apiRetries: 5,                   // повторы при 429/5xx
    apiBackoffMs: 1500,              // первая пауза перед повтором, дальше ×2
    openWaitMs: 15000,               // сколько ждать появления истории после клика Show
    pauseMs: 400,                    // пауза между тикетами
    maxConsecutiveErrors: 3,         // столько ошибок подряд — останавливаемся (история не открывается)
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
      XP.open = function (m, u) { this.__thU = String(u); this.__thH = {}; return o.apply(this, arguments); };
      XP.setRequestHeader = function (k, v) { if (this.__thH) this.__thH[k] = v; return sh.apply(this, arguments); };
      XP.send = function () {
        if (this.__thU && this.__thU.indexOf('paymentsupporthistory') >= 0) Object.assign(learned, this.__thH);
        return sd.apply(this, arguments);
      };
    } catch (e) { /* перехват необязателен */ }
  })();

  const COLORS = { yes: '#b7eb8f', no: '#ffd591', alt: '#adc6ff', error: '#d9d9d9' };

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
  const SKEY = 'th_settings_v1';
  const DEFAULTS = { cur: CFG.currentStatuses.map(() => true), target: CFG.targetStatus, first: true, hide: false, badges: true, open: false, api: true };
  function loadSettings() {
    let raw = null;
    try { raw = typeof GM_getValue === 'function' ? GM_getValue(SKEY, null) : localStorage.getItem(SKEY); } catch (e) { /* ignore */ }
    try { return Object.assign({}, DEFAULTS, raw ? JSON.parse(raw) : {}); } catch (e) { return Object.assign({}, DEFAULTS); }
  }
  const S = loadSettings();
  if (!Array.isArray(S.cur) || S.cur.length !== CFG.currentStatuses.length) S.cur = DEFAULTS.cur.slice();
  function saveSettings() {
    try {
      const raw = JSON.stringify(S);
      if (typeof GM_setValue === 'function') GM_setValue(SKEY, raw); else localStorage.setItem(SKEY, raw);
    } catch (e) { /* ignore */ }
  }
  const selectedStatuses = () => CFG.currentStatuses.filter((_, i) => S.cur[i]);

  // ---------- список жалоб ----------
  // главная таблица — та, где есть и Ticket ID, и Ticket history (таблица истории — другая)
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
      status: heads.findIndex((h) => CFG.statusHeader.test(h)),
      hist: heads.findIndex((h) => h.includes(CFG.historyHeader)),
    };
    let trs = [...table.querySelectorAll(':scope > tbody > tr')];
    if (!trs.length) trs = [...table.querySelectorAll('tbody tr')];
    const sel = selectedStatuses();
    const res = [];
    trs.forEach((tr) => {
      const cells = tr.children;
      if (cells.length <= Math.max(idx.ticket, idx.status, idx.hist)) return;
      const id = (cells[idx.ticket].innerText || '').trim();
      if (!id) return;
      const status = idx.status >= 0 ? (cells[idx.status].textContent || '').replace(/\s+/g, ' ').trim() : '';
      let btn = null;
      if (idx.hist >= 0) {
        const cell = cells[idx.hist];
        const els = [...cell.querySelectorAll('a,button,span,[role=button]')];
        btn = els.find((e) => /^(show|history|история)$/i.test(txt(e).trim())) || els[0] || (txt(cell).trim() ? cell : null);
      }
      res.push({ tr, id, tCell: cells[idx.ticket], status, btn, inScope: idx.status >= 0 && sel.some((x) => statusEq(x, status)) });
    });
    res.table = table; res.idx = idx;
    return res;
  }

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
  function historyRows(table) {
    const hth = [...table.querySelectorAll('th')].find((x) => norm(txt(x)).includes(CFG.histMarker));
    if (!hth) return [];
    const heads = [...hth.parentElement.children].map((c) => norm(txt(c)));
    const iExt = heads.findIndex((h) => CFG.histStatusHeader.test(h));
    const iDate = heads.findIndex((h) => h.includes(CFG.histMarker));
    if (iExt < 0) return [];
    let trs = [...table.querySelectorAll('tbody tr')];
    if (!trs.length) trs = [...table.querySelectorAll('tr')].filter((tr) => tr !== hth.parentElement);
    return trs.filter((tr) => tr.children.length > iExt && !tr.querySelector('th'))
      .map((tr) => ({
        date: iDate >= 0 ? (tr.children[iDate].textContent || '').replace(/\s+/g, ' ').trim() : '',
        ext: (tr.children[iExt].textContent || '').replace(/\s+/g, ' ').trim(),
      })).filter((e) => e.ext);
  }
  // цепочка статусов от старых к новым (порядок определяем по датам), подряд идущие дубли схлопываются
  function buildChain(entries) {
    const list = entries.slice();
    if (list.length > 1 && list[0].date && list[list.length - 1].date && list[0].date > list[list.length - 1].date) list.reverse();
    const seq = [];
    list.forEach((e) => { if (!seq.length || seq[seq.length - 1].ext !== e.ext) seq.push(e); });
    return seq;
  }

  function fillResult(res, entries) {
    const seq = buildChain(entries);
    const hit = seq.find((e) => statusEq(e.ext, S.target));
    const alt = CFG.altStatuses.filter((a) => seq.some((e) => statusEq(e.ext, a)));
    res.alt = hit ? [] : alt;
    res.state = hit ? 'yes' : alt.length ? 'alt' : 'no';
    res.when = hit ? hit.date : '';
    res.chain = seq.map((e) => e.ext).join(' → ');
    res.n = entries.length;
  }

  // ---------- фоновый режим: запрос истории без открытия окна ----------
  // ищет в ответе массив объектов, у которых есть поле с датой
  function findRows(node, key, depth) {
    if (!node || typeof node !== 'object' || depth > 6) return null;
    if (Array.isArray(node)) {
      if (node.length && node[0] && typeof node[0] === 'object' && key in node[0]) return node;
      for (const x of node) { const r = findRows(x, key, depth + 1); if (r) return r; }
      return null;
    }
    for (const k of Object.keys(node)) { const r = findRows(node[k], key, depth + 1); if (r) return r; }
    return null;
  }
  // имена полей берём из словаря columns ответа: «поле → заголовок» (External status, Date of ticket update)
  function parseApiHistory(j) {
    const cols = (j && (j.columns || (j.params && j.params.columns))) || null;
    if (!cols) throw new Error('в ответе нет описания колонок');
    const keyByLabel = (re) => Object.keys(cols).find((k) => re.test(norm(String(cols[k]))));
    const extKey = keyByLabel(/^external status$/) || Object.keys(cols).find((k) => /external.?status/i.test(k));
    const dateKey = keyByLabel(/^date of ticket update$/) || 'dateEdit';
    if (!extKey) throw new Error('в ответе нет колонки External status');
    const arr = findRows(j, dateKey, 0);
    if (!arr) throw new Error('в ответе нет строк истории');
    const out = arr.map((o) => {
      const v = o[extKey];
      return { date: String(o[dateKey] == null ? '' : o[dateKey]).trim(), ext: (v && typeof v === 'object' ? '[object]' : String(v == null ? '' : v)).replace(/\s+/g, ' ').trim() };
    }).filter((e) => e.ext);
    if (!out.length) throw new Error('пустая история');
    if (out.some((e) => e.ext === '[object]')) throw new Error('статус в ответе не строка');
    return out;
  }
  async function fetchHistoryApi(id) {
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
    let j;
    try { j = await r.json(); } catch (e) { throw new Error('ответ не JSON'); }
    return parseApiHistory(j);
  }

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
  // res.state: yes — в истории есть искомый статус, no — нет, error — не удалось получить историю
  async function checkRow(row, mainTable) {
    const res = { id: row.id, cur: row.status, state: 'error', chain: '', when: '', n: 0, note: '', stuck: false };
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
          const n = historyRows(table).length;
          if (n > 0 && n === prev) same++; else { same = 0; prev = n; }
        }
      }
      const entries = historyRows(table);
      if (!entries.length) { res.note = 'в таблице истории нет строк (или нет колонки External status)'; }
      else {
        fillResult(res, entries);
      }
      if (!fromDoc && !(await closeHistory(table, row.btn))) { res.stuck = true; res.note = (res.note ? res.note + '; ' : '') + 'окно истории не закрылось'; }
      return res;
    } finally {
      PW.open = origOpen;
    }
  }

  // ---------- меню ----------
  const style = document.createElement('style');
  style.textContent = `
    #th-panel, #th-panel * { box-sizing: border-box; }
    #th-panel .th-sec { font-weight: 700; margin: 8px 0 4px; color: #23909f; font-size: 11px; text-transform: uppercase; letter-spacing: .4px; }
    #th-panel label { display: block; margin-bottom: 4px; cursor: pointer; }
    #th-panel button { width: 100%; margin-bottom: 4px; }
    #th-panel .th-grid { display: grid; grid-template-columns: 1fr 1fr 1fr 1fr; gap: 4px; margin-bottom: 4px; }
    #th-panel .th-grid button { margin: 0; padding: 3px 2px; font-size: 12px; }
    #th-panel input[type=text] { width: 100%; margin: 2px 0 4px; }
    #th-panel pre { max-height: 160px; overflow: auto; font-size: 11px; white-space: pre-wrap; background: #f5f5f5; padding: 4px; margin: 4px 0 0; }
    td[data-thb]::after { content: attr(data-thb); display: block; margin-top: 2px; font: 600 10px/1.2 Arial, sans-serif; color: inherit; white-space: normal; }
  `;
  document.head.appendChild(style);

  const launcher = document.createElement('button');
  launcher.type = 'button';
  launcher.title = 'Проверка истории статусов';
  launcher.style.cssText =
    'position:fixed;right:12px;bottom:12px;z-index:999999;background:#531dab;color:#fff;border:none;border-radius:20px;' +
    'box-sizing:border-box;width:80px;padding:8px 0;text-align:center;white-space:nowrap;font:700 13px Arial,sans-serif;cursor:pointer;box-shadow:0 2px 10px rgba(0,0,0,.35)';
  launcher.textContent = 'История';
  document.body.appendChild(launcher);

  const panel = document.createElement('div');
  panel.id = 'th-panel';
  panel.style.cssText =
    'position:fixed;right:12px;bottom:56px;z-index:999999;background:#fff;border:2px solid #531dab;border-radius:8px;' +
    'padding:10px;font:13px/1.4 Arial,sans-serif;width:300px;max-height:calc(100vh - 80px);overflow:auto;' +
    'box-shadow:0 4px 16px rgba(0,0,0,.3);color:#222;display:none';
  panel.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center"><b>Проверка истории статусов</b><span id="th-x" style="cursor:pointer;font-size:16px;padding:0 4px" title="Закрыть меню">✕</span></div>
    <div id="th-sets">
      <div class="th-sec">Текущий статус тикета</div>
      ${CFG.currentStatuses.map((s, i) => `<label><input type="checkbox" id="th-cur${i}"> ${s}</label>`).join('')}
      <label><input type="checkbox" id="th-api"> Фоновый режим (без открытия окон)</label>
      <div class="th-sec">Искать в истории статус</div>
      <input type="text" id="th-target">
      <div class="th-sec">Отображение</div>
      <label><input type="checkbox" id="th-hide"> Скрыть тикеты других статусов</label>
      <label><input type="checkbox" id="th-badges"> Подписи в ячейке</label>
    </div>
    <div class="th-sec">Результат</div>
    <div id="th-stat" style="margin-bottom:6px">—</div>
    <button id="th-run">Проверить страницу</button>
    <button id="th-retry">Повторить ошибки</button>
    <div class="th-sec">Копировать ID</div>
    <div class="th-grid"><button id="th-c-yes">Был</button><button id="th-c-no">Не было</button><button id="th-c-alt">Appr./Cred.</button><button id="th-c-error">Ошибки</button></div>
    <div class="th-sec">Все страницы</div>
    <label><input type="checkbox" id="th-first"> Начать с 1-й страницы</label>
    <button id="th-crawl">Все страницы → проверить</button>
    <button id="th-resume" disabled>Продолжить с текущей</button>
    <button id="th-stop" disabled>Стоп</button>
    <button id="th-xlsx" disabled>Скачать Excel</button>
    <button id="th-diag">Диагностика (1 тикет)</button>
    <div id="th-msg" style="font-size:12px;color:#555"></div>
    <pre id="th-pre" style="display:none"></pre>`;
  document.body.appendChild(panel);
  const $ = (q) => panel.querySelector(q);
  const msg = (t) => ($('#th-msg').textContent = t);
  const setStat = (t) => ($('#th-stat').textContent = t);

  const showPanel = (v) => { S.open = v; panel.style.display = v ? 'block' : 'none'; saveSettings(); };
  launcher.onclick = () => showPanel(panel.style.display === 'none');
  $('#th-x').onclick = () => showPanel(false);
  showPanel(!!S.open);

  function syncUI() {
    CFG.currentStatuses.forEach((_, i) => { $('#th-cur' + i).checked = !!S.cur[i]; });
    $('#th-target').value = S.target;
    $('#th-api').checked = !!S.api;
    $('#th-hide').checked = !!S.hide; $('#th-badges').checked = !!S.badges; $('#th-first').checked = !!S.first;
  }
  const changed = () => { saveSettings(); run(); };
  CFG.currentStatuses.forEach((_, i) => { $('#th-cur' + i).onchange = (e) => { S.cur[i] = e.target.checked; changed(); }; });
  $('#th-target').onchange = (e) => { S.target = e.target.value.trim() || CFG.targetStatus; changed(); };
  $('#th-api').onchange = (e) => { S.api = e.target.checked; if (S.api) { apiOn = true; apiRetried = false; } saveSettings(); };
  $('#th-hide').onchange = (e) => { S.hide = e.target.checked; changed(); };
  $('#th-badges').onchange = (e) => { S.badges = e.target.checked; changed(); };
  $('#th-first').onchange = (e) => { S.first = e.target.checked; saveSettings(); };

  const results = new Map();   // id -> результат проверки
  let busy = false;            // идёт проверка истории (не перерисовываем список в это время)
  let running = false;
  let aborted = false;
  let curPageNo = 0;

  function paint(rows) {
    rows.forEach((r) => {
      const res = results.get(r.id);
      let color = '', badge = '', title = '';
      if (res) {
        color = COLORS[res.state];
        if (res.state === 'yes') { badge = `был ${S.target}`; title = `В истории есть «${S.target}» (${res.when})\n${res.chain}`; }
        else if (res.state === 'alt') { badge = `нет ${S.target}, есть ${res.alt.join(' + ')}`; title = `В истории нет «${S.target}», но есть: ${res.alt.join(', ')}\n${res.chain}`; }
        else if (res.state === 'no') { badge = `не было ${S.target}`; title = `В истории нет «${S.target}»\n${res.chain}`; }
        else { badge = 'ошибка'; title = res.note; }
      } else if (r.inScope) title = 'К проверке: история ещё не проверена';
      r.tCell.style.background = color;
      r.tCell.style.fontWeight = color ? '700' : '';
      r.tCell.style.boxShadow = !res && r.inScope ? 'inset 0 0 0 3px #531dab' : '';
      r.tCell.title = title;
      if (S.badges && badge) r.tCell.setAttribute('data-thb', badge); else r.tCell.removeAttribute('data-thb');
      r.tr.style.display = S.hide && !r.inScope ? 'none' : '';
    });
  }
  function counts() {
    const v = [...results.values()];
    return { yes: v.filter((x) => x.state === 'yes').length, no: v.filter((x) => x.state === 'no').length, alt: v.filter((x) => x.state === 'alt').length, error: v.filter((x) => x.state === 'error').length };
  }
  function run() {
    const rows = analyze();
    if (!rows) { setStat('Таблица не найдена'); return; }
    paint(rows);
    const c = counts();
    const warn = [];
    if (rows.idx.status < 0) warn.push('External status');
    if (rows.idx.hist < 0) warn.push('Ticket history');
    const w = warn.length ? `<div style="color:#d4380d;font-size:12px">⚠ Нет колонки: ${warn.join(', ')} (включите её в Column settings)</div>` : '';
    $('#th-stat').innerHTML = `Строк: <b>${rows.length}</b> · в выбранных статусах: <b>${rows.filter((r) => r.inScope).length}</b> · ` +
      `к проверке: <b>${rows.filter((r) => r.inScope && !results.has(r.id)).length}</b><br>` +
      `всего проверено: был <b style="color:#389e0d">${c.yes}</b> · не было <b style="color:#d46b08">${c.no}</b> · ошибок <b>${c.error}</b>${w}`;
  }

  let apiOn = true;          // фоновый режим доступен (выключается, если запросы не работают)
  let apiRetried = false;
  async function apiPass() {
    const L = analyze();
    if (!L) return;
    const rows = L.filter((r) => r.inScope && !results.has(r.id));
    if (!rows.length) return;
    busy = true;
    let ok = 0, fail = 0, firstErr = '', next = 0;
    const worker = async () => {
      while (!aborted) {
        const row = rows[next++];
        if (!row) return;
        try {
          const entries = await fetchHistoryApi(row.id);
          const res = { id: row.id, cur: row.status, state: 'error', chain: '', when: '', n: 0, note: '', stuck: false, page: curPageNo || '', via: 'фон' };
          fillResult(res, entries);
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

  // проверка всех непроверенных тикетов текущей страницы; 'ok' | 'errors' | 'aborted'
  async function checkPage() {
    let consec = 0;
    // после проверки через окно сайт сам отправил запрос — мы запомнили его заголовки, можно ещё раз попробовать фоновый режим
    if (!apiOn && !apiRetried && Object.keys(learned).length) { apiOn = true; apiRetried = true; }
    if (S.api && apiOn) { await apiPass(); if (aborted) return 'aborted'; }
    for (;;) {
      if (aborted) return 'aborted';
      const L = analyze();
      if (!L) return 'ok';
      const row = L.find((r) => r.inScope && !results.has(r.id));
      if (!row) return 'ok';
      busy = true;
      setStat(`Проверяю историю: ${row.id}…`);
      let res;
      try { res = await checkRow(row, L.table); } catch (e) { res = { id: row.id, cur: row.status, state: 'error', chain: '', when: '', n: 0, note: 'ошибка: ' + e.message, stuck: false }; }
      busy = false;
      res.page = curPageNo || '';
      results.set(row.id, res);
      run();
      if (res.state === 'error') { consec++; if (consec >= CFG.maxConsecutiveErrors || res.stuck) return 'errors'; } else consec = 0;
      await sleep(CFG.pauseMs);
    }
  }

  function setRunning(v) {
    running = v;
    $('#th-run').disabled = v; $('#th-retry').disabled = v; $('#th-diag').disabled = v;
    $('#th-crawl').disabled = v;
    $('#th-resume').disabled = v || pagesDone === 0;
    $('#th-stop').disabled = !v;
    $('#th-xlsx').disabled = v || results.size === 0;
    panel.querySelectorAll('#th-sets input').forEach((el) => { el.disabled = v; });
  }

  $('#th-run').onclick = async () => {
    if (running) return;
    aborted = false; curPageNo = 0; setRunning(true);
    const r = await checkPage();
    run(); setRunning(false);
    msg(r === 'errors' ? 'Остановлено из-за ошибок: история не открывается. Нажмите «Диагностика (1 тикет)».' : r === 'aborted' ? 'Остановлено.' : 'Страница проверена.');
  };
  $('#th-retry').onclick = () => {
    [...results.entries()].forEach(([id, r]) => { if (r.state === 'error') results.delete(id); });
    run(); $('#th-run').onclick();
  };

  function copy(text) {
    if (typeof GM_setClipboard === 'function') GM_setClipboard(text);
    else navigator.clipboard.writeText(text);
  }
  const LBL = { yes: 'был в истории', no: 'не было в истории', alt: 'нет Received, есть Approved/Credited', error: 'ошибки' };
  ['yes', 'no', 'alt', 'error'].forEach((k) => {
    $('#th-c-' + k).onclick = () => {
      const ids = [...results.values()].filter((r) => r.state === k).map((r) => r.id);
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
      if (panel.contains(el) || el.closest('#th-panel,#tn-panel,#ts-panel') || el.closest('table')) return false;
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

  function findNext() {
    if (CFG.nextSelector) return document.querySelector(CFG.nextSelector);
    const p = getPager();
    if (!p) return null;
    const lastItem = p.items[p.items.length - 1];
    if (!isNum(lastItem)) return lastItem;            // стрелка «›»
    const i = p.items.indexOf(p.active);
    return i >= 0 && i < p.items.length - 1 ? p.items[i + 1] : null;
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
    missing: 'не найдена кнопка «далее» / блок пагинации',
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
        if (r === 'errors') { reason = 'история не открывается (3 ошибки подряд) — нажмите «Диагностика (1 тикет)»'; break; }
        if (r === 'aborted') break;
        msg(`Страница ${pagesDone} готова · был: ${counts().yes}, не было: ${counts().no}, ошибок: ${counts().error}`);
        if (t && t.to >= t.total) { reason = 'последняя страница'; break; }
        const nr = await clickAndWait(findNext);
        if (nr !== 'ok') { reason = WHY[nr] || nr; break; }
      }
      if (aborted) reason = WHY.aborted;
    } catch (e) {
      reason = 'ошибка: ' + e.message;
      console.error('[th]', e);
    }
    busy = false;
    run();
    const c = counts();
    msg(`${note}Конец: ${reason || '—'}. Страниц: ${pagesDone}. Был ${S.target}: ${c.yes}, не было: ${c.no}, нет Received но есть Approved/Credited: ${c.alt}, ошибок: ${c.error}.${coverage()}`);
    setRunning(false);
  }

  $('#th-resume').onclick = () => allPages(true);
  $('#th-crawl').onclick = () => allPages(false);
  $('#th-stop').onclick = () => { aborted = true; };

  // Excel: «Был в истории» / «Не было в истории» / «Ошибки»
  $('#th-xlsx').onclick = () => {
    if (typeof XLSX === 'undefined') { msg('Библиотека XLSX не загрузилась.'); return; }
    const num = (id) => (/^\d{1,15}$/.test(id) ? Number(id) : id);
    const all = [...results.values()];
    const yes = all.filter((r) => r.state === 'yes');
    const no = all.filter((r) => r.state === 'no');
    const alt = all.filter((r) => r.state === 'alt');
    const err = all.filter((r) => r.state === 'error');
    const wb = XLSX.utils.book_new();
    const add = (name, aoa, wch) => {
      const sh = XLSX.utils.aoa_to_sheet(aoa);
      sh['!cols'] = wch.map((w) => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, sh, name);
    };
    add('Был в истории', [['ticket_id', 'текущий статус', `дата первого «${S.target}»`, 'записей в истории', 'цепочка статусов (от старых к новым)', 'page'],
      ...yes.map((r) => [num(r.id), r.cur, r.when, r.n, r.chain, r.page])], [14, 28, 22, 10, 90, 6]);
    add('Не было в истории', [['ticket_id', 'текущий статус', 'записей в истории', 'цепочка статусов (от старых к новым)', 'page'],
      ...no.map((r) => [num(r.id), r.cur, r.n, r.chain, r.page])], [14, 28, 10, 90, 6]);
    add(`Нет Received, есть Appr-Cred`, [['ticket_id', 'текущий статус', 'найденные статусы', 'записей в истории', 'цепочка статусов (от старых к новым)', 'page'],
      ...alt.map((r) => [num(r.id), r.cur, r.alt.join(' + '), r.n, r.chain, r.page])], [14, 28, 44, 10, 90, 6]);
    add('Ошибки', [['ticket_id', 'текущий статус', 'причина', 'page'], ...err.map((r) => [num(r.id), r.cur, r.note, r.page])], [14, 28, 70, 6]);
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'history_check_' + new Date().toISOString().slice(0, 10) + '.xlsx';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    msg(`Excel: «Был в истории» — ${yes.length}, «Не было» — ${no.length}, «Approved/Credited» — ${alt.length}, «Ошибки» — ${err.length}`);
  };

  // ---------- диагностика: что происходит после клика по Show (копируется в буфер — пришлите разработчику) ----------
  $('#th-diag').onclick = async () => {
    if (running) return;
    const L = analyze();
    if (!L) { msg('Таблица не найдена'); return; }
    const row = L.find((r) => r.inScope && r.btn) || L.find((r) => r.btn);
    if (!row) { msg('Нет строки с кнопкой Show'); return; }
    running = true; busy = true;
    const a = row.btn.closest('a[href]');
    const info = {
      ticket: row.id, current_status: row.status, in_scope: row.inScope,
      show_button_html: row.btn.outerHTML.slice(0, 300), show_href: a ? a.getAttribute('href') : null, show_target: a ? a.target : null,
      url_before: location.href, tables_before: document.querySelectorAll('table').length,
    };
    // data-атрибуты кнопки и строки: там часто лежит внутренний id тикета, нужный для запроса истории
    info.show_data_attrs = Object.assign({}, row.btn.dataset || {});
    info.row_data_attrs = Object.assign({}, row.tr.dataset || {});
    info.show_onclick = row.btn.getAttribute('onclick');
    const calls = [];
    const net = [];
    const origOpen = PW.open;
    const origFetch = PW.fetch;
    const XP = PW.XMLHttpRequest && PW.XMLHttpRequest.prototype;
    const origXOpen = XP && XP.open, origXSend = XP && XP.send;
    PW.open = function (u) { calls.push(String(u)); return null; };
    // записываем fetch / XHR, которые уходят после клика (адрес, метод, тело, начало ответа)
    if (origFetch) {
      PW.fetch = async function (...args) {
        const req = args[0], init = args[1] || {};
        const rec = { type: 'fetch', url: String(typeof req === 'string' ? req : (req && req.url)), method: init.method || (req && req.method) || 'GET', body: init.body ? String(init.body).slice(0, 500) : null };
        net.push(rec);
        try {
          const r = await origFetch.apply(this, args);
          rec.status = r.status; rec.content_type = r.headers.get('content-type');
          try { rec.preview = (await r.clone().text()).slice(0, 500); } catch (e) { /* ignore */ }
          return r;
        } catch (e) { rec.error = String(e); throw e; }
      };
    }
    if (XP) {
      XP.open = function (m, u) { this.__thd = { type: 'xhr', method: m, url: String(u) }; return origXOpen.apply(this, arguments); };
      XP.send = function (b) {
        const d = this.__thd;
        if (d) {
          d.body = b ? String(b).slice(0, 500) : null; net.push(d);
          this.addEventListener('load', () => {
            d.status = this.status; d.content_type = this.getResponseHeader('content-type');
            try { d.preview = String(this.responseText).slice(0, 500); } catch (e) { /* ignore */ }
          });
        }
        return origXSend.apply(this, arguments);
      };
    }
    try {
      clickEl(row.btn);
      await sleep(4000);
    } finally {
      PW.open = origOpen;
      if (origFetch) PW.fetch = origFetch;
      if (XP) { XP.open = origXOpen; XP.send = origXSend; }
    }
    info.network_requests = net;
    info.url_after = location.href;
    info.window_open_calls = calls;
    info.tables_after = document.querySelectorAll('table').length;
    info.main_table_still_present = !!analyze();
    const ht = findHistoryTable(document, L.table);
    info.history_table_found = !!ht;
    if (ht) {
      info.history_headers = [...ht.querySelectorAll('th')].map((t) => norm(txt(t))).slice(0, 20);
      const rows = historyRows(ht);
      info.history_rows = rows.length;
      info.history_sample = rows.slice(0, 5);
      const root = modalRoot(ht);
      info.modal_root = (root.outerHTML || '').slice(0, 200);
      const x = findCloseBtn(root);
      info.close_button = x ? x.outerHTML.slice(0, 200) : null;
      info.closed_ok = await closeHistory(ht, row.btn);
    } else {
      info.other_tables_headers = [...document.querySelectorAll('table')].filter((t) => t !== L.table).slice(0, 6)
        .map((t) => [...t.querySelectorAll('th')].map((x) => norm(txt(x))).slice(0, 12));
    }
    busy = false; running = false;
    const json = JSON.stringify(info, null, 1);
    console.log('[th diag]', info);
    copy(json);
    const pre = $('#th-pre');
    pre.style.display = 'block'; pre.textContent = json;
    msg(ht ? 'История найдена. Результат диагностики скопирован в буфер.' : 'Таблица истории не найдена. Результат диагностики скопирован в буфер — пришлите его.');
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
