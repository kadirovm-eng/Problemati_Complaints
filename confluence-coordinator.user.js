// ==UserScript==
// @name         Confluence Координатор
// @namespace    confluence-coordinator
// @version      2.7.1
// @description  Проверка правок в Confluence через бота Reddy — без сервера: заявки, чат проверяющих, уведомления старшим, кабинет
// @match        https://doc.office.lan/*
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @connect      bot.reddy.team
// @connect      raw.githubusercontent.com
// @connect      gist.githubusercontent.com
// @updateURL    https://raw.githubusercontent.com/kadirovm-eng/Problemati_Complaints/main/confluence-coordinator.user.js
// @downloadURL  https://raw.githubusercontent.com/kadirovm-eng/Problemati_Complaints/main/confluence-coordinator.user.js
// @run-at       document-idle
// @noframes
// ==/UserScript==

/*
 * Как это устроено
 * ─────────────────
 * • Данные (заявки, сотрудники, настройки) — в скрытых свойствах служебной страницы Confluence
 *   (REST /rest/api/content/{id}/property). Всем пользователям плагина нужно право редактировать эту страницу.
 * • Reddy BotAPI v2 вызывается прямо из браузера (GM_xmlhttpRequest → bot.reddy.team).
 * • Нажатия кнопок и сообщения боту забирает «дежурная» вкладка: любая открытая вкладка Confluence
 *   с плагином пытается подключиться к Reddy (getupdate), Reddy пускает только одно подключение.
 * • Кабинет открывается кнопкой плагина или при заходе на служебную страницу.
 */
(function () {
  'use strict';

  // ═══════════════ НАСТРОЙКА: ID служебной страницы Confluence ═══════════════
  // Создайте пустую страницу (например «Confluence Координатор — служебная»), откройте её
  // и возьмите число из адреса (…pageId=123456). Это единственное, что нужно вписать в скрипт.
  const STORAGE_PAGE_ID = GM_getValue('storagePageId', '730778260');
  // ═══════════════════════════════════════════════════════════════════════════

  const REDDY_API = 'https://bot.reddy.team';
  const TZ = 'Europe/Moscow';
  const TAB = Math.random().toString(36).slice(2);
  const DEBUG = GM_getValue('debug', false);
  const VERSION = '2.7.1';
  const verNum = (v) => String(v || '0').split('.').reduce((a, x) => a * 1000 + (+x || 0), 0);
  const log = (...a) => DEBUG && console.log('[Координатор]', ...a);

  const meta = (n) => { const el = document.querySelector(`meta[name="${n}"]`); return el ? el.getAttribute('content') : null; };
  if (!(meta('ajs-page-id') || meta('confluence-base-url') || window.AJS || /\/pages\/|\/display\//.test(location.pathname))) return;

  // ─────────────────────────── утилиты ───────────────────────────
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const now = () => Date.now();
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const safe = (s) => String(s || '').replace(/\[/g, '⁅').replace(/\]/g, '⁆'); // текст пользователя не должен становиться BBCode
  const short = (s, n = 400) => safe(s && s.length > n ? s.slice(0, n - 1) + '…' : s || '');
  function dur(ms) {
    if (ms == null || isNaN(ms) || ms < 0) return '—';
    const m = Math.round(ms / 60000);
    if (m < 1) return 'меньше минуты';
    if (m < 60) return `${m} мин`;
    const h = Math.floor(m / 60), mm = m % 60;
    if (h < 24) return mm ? `${h} ч ${mm} мин` : `${h} ч`;
    const d = Math.floor(h / 24), hh = h % 24;
    return hh ? `${d} д ${hh} ч` : `${d} д`;
  }
  const fmtT = (ts) => (ts ? new Intl.DateTimeFormat('ru-RU', { timeZone: TZ, day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' }).format(new Date(ts)) : '—');
  function tzParts(ts) {
    return new Intl.DateTimeFormat('en-CA', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
      .formatToParts(new Date(ts)).reduce((a, p) => ((a[p.type] = p.value), a), {});
  }
  function startOfDay(ts = now()) { const p = tzParts(ts); return ts - (((+p.hour * 60 + +p.minute) * 60 + +p.second) * 1000 + (ts % 1000)); }
  const dayKey = (ts = now()) => { const p = tzParts(ts); return `${p.year}-${p.month}-${p.day}`; };

  // ─────────────────────────── Confluence REST ───────────────────────────
  function confluenceBase() {
    const b = meta('confluence-base-url') || meta('ajs-base-url');
    if (b) return b.replace(/\/$/, '');
    const ctx = meta('ajs-context-path');
    if (ctx !== null) return location.origin + ctx;
    return location.origin + (location.pathname.startsWith('/wiki/') ? '/wiki' : '');
  }
  const CB = confluenceBase();

  class CfError extends Error { constructor(status, text) { super(`Confluence ${status}: ${String(text).slice(0, 200)}`); this.status = status; } }
  async function cf(method, path, body) {
    const res = await fetch(CB + path, {
      method, credentials: 'include',
      headers: { Accept: 'application/json', 'Content-Type': 'application/json', 'X-Atlassian-Token': 'no-check' },
      body: body ? JSON.stringify(body) : undefined,
    });
    const text = await res.text();
    if (!res.ok) throw new CfError(res.status, text);
    return text ? JSON.parse(text) : null;
  }

  // Хранилище: свойства служебной страницы. Каждое значение — JSON, с номером версии для защиты от одновременной записи.
  const P = `/rest/api/content/${encodeURIComponent(STORAGE_PAGE_ID)}/property`;
  async function getProp(key) {
    try { const r = await cf('GET', `${P}/${key}?expand=version`); return { value: r.value, version: r.version.number }; } catch (e) { if (e.status === 404) return null; throw e; }
  }
  async function putProp(key, value, version) {
    if (version == null) return cf('POST', P, { key, value });
    return cf('PUT', `${P}/${key}`, { key, value, version: { number: version + 1, minorEdit: true } });
  }
  /** Атомарное изменение: fn(текущее значение) → новое значение. При одновременной записи — повтор. */
  async function mutate(key, fn, init) {
    for (let i = 0; i < 8; i++) {
      const cur = await getProp(key);
      const value = cur ? cur.value : (typeof init === 'function' ? init() : init);
      const next = await fn(structuredClone(value));
      if (next === undefined) return value; // изменений нет
      try { await putProp(key, next, cur ? cur.version : null); return next; } catch (e) {
        if (![409, 400, 412].includes(e.status)) throw e;
        await sleep(150 + Math.random() * 400);
      }
    }
    throw new Error('Не удалось сохранить данные: слишком много одновременных изменений, попробуйте ещё раз');
  }
  async function listProps(prefix) {
    const out = [];
    for (let start = 0; ; start += 200) {
      const r = await cf('GET', `${P}?expand=version&limit=200&start=${start}`);
      for (const p of r.results || []) if (p.key.startsWith(prefix)) out.push(p.value);
      if (!r.results || r.results.length < 200) break;
    }
    return out;
  }

  // ─────────────────────────── Данные ───────────────────────────
  const K = { config: 'crb-config', meta: 'crb-meta', user: (id) => `crb-u-${id}`, req: (id) => `crb-r-${id}` };
  const META0 = () => ({ nextId: 1, byPage: {}, open: {}, pending: {}, lastDigest: '', lastTick: 0 });

  let cfgCache = null, cfgAt = 0;
  async function getConfig(force) {
    if (!force && cfgCache && now() - cfgAt < 5 * 60000) return cfgCache;
    const p = await getProp(K.config);
    cfgCache = p ? p.value : null; cfgAt = now();
    return cfgCache;
  }
  const getMeta = async () => ((await getProp(K.meta)) || { value: META0() }).value;
  const updateMeta = (fn) => mutate(K.meta, fn, META0);
  const getUser = async (id) => (id ? ((await getProp(K.user(id))) || {}).value || null : null);
  const saveUser = (id, fn) => mutate(K.user(id), fn, () => ({ id: String(id), name: null, key: null, role: 'employee', notify: true, active: true, confluence: null, created: now() }));
  const getReq = async (id) => ((await getProp(K.req(id))) || {}).value || null;
  const isSenior = (u) => !!u && (u.role === 'senior' || u.role === 'admin');
  const uname = (u, fb) => (u && u.name) || fb || (u ? `ID ${u.id}` : 'Неизвестный');

  /** Создать пользователя при первом контакте, дописать имя/номер */
  async function touchUser(id, name, key) {
    if (!id) return null;
    const cfg = await getConfig();
    return saveUser(id, (u) => {
      let changed = false;
      if (name && !u.name) { u.name = name; changed = true; }
      if (key && u.key !== String(key)) { u.key = String(key); changed = true; }
      if (cfg && u.role !== 'admin' && (cfg.admins || []).includes(String(id))) { u.role = 'admin'; changed = true; }
      if (!u.created) { u.created = now(); changed = true; }
      return changed || !u._saved ? Object.assign(u, { _saved: 1 }) : undefined;
    });
  }

  // ─────────────────────────── Reddy BotAPI v2 ───────────────────────────
  function gm(method, url, { token, body, timeout = 20000 } = {}) {
    return new Promise((resolve) => {
      GM_xmlhttpRequest({
        method, url, timeout,
        headers: { 'Content-Type': 'application/json', 'X-Bot-Token': token },
        data: body ? JSON.stringify(body) : undefined,
        onload: (r) => { let j = null; try { j = JSON.parse(r.responseText); } catch (e) { /* */ } resolve({ status: r.status, json: j }); },
        onerror: () => resolve({ status: 0, json: null }),
        ontimeout: () => resolve({ status: -1, json: null }),
      });
    });
  }
  const chatNext = new Map();
  async function throttle(key) { const t = now(), at = Math.max(t, chatNext.get(key) || 0); chatNext.set(key, at + 1100); if (at > t) await sleep(at - t); }
  async function reddy(method, body, query, token) {
    token = token || (await getConfig()).token;
    const url = `${REDDY_API}/v2/${method}${query ? '?' + new URLSearchParams(query) : ''}`;
    for (let i = 0; ; i++) {
      const r = await gm(body ? 'POST' : 'GET', url, { token, body });
      if (r.status === 429 && i < 3) { await sleep(1300 * (i + 1)); continue; }
      if (r.status !== 200 || (r.json && r.json.errorCode)) throw new Error(`Reddy ${r.status}${r.json && r.json.errorMessage ? ': ' + r.json.errorMessage : ''}`);
      return r.json || {};
    }
  }
  const kb = (buttons) => (buttons || []).filter((r) => r && r.length).map((row) => row.slice(0, 8).map((b) => (b.url ? { type: 'link', title: b.text, url: b.url } : { type: 'action', title: b.text, data: b.data })));
  const clip = (s) => (s.length > 4000 ? s.slice(0, 3999) + '…' : s);
  async function send(to, text, buttons, opts = {}) {
    const r = await sendFull(to, text, buttons, opts);
    return r ? r.id : null;
  }
  /** → { id, chat } (chat — ID чата, в том числе личного: он нужен, чтобы потом отредактировать сообщение) или null */
  async function sendFull(to, text, buttons, opts = {}) {
    const tgt = to.chat ? { chat: Number(to.chat) } : { userId: Number(to.user) };
    try {
      await throttle(JSON.stringify(tgt));
      const r = await reddy('send', { ...tgt, msg: clip(text), keyboard: kb(buttons), ...(opts.reply ? { reply: Number(opts.reply) } : {}) });
      return { id: r.messageId != null ? String(r.messageId) : '', chat: r.chatId != null ? String(r.chatId) : to.chat ? String(to.chat) : null };
    } catch (e) { console.warn('[Координатор] send:', e.message); return null; }
  }
  async function edit(chat, messageId, text, buttons) {
    if (!messageId) return false;
    try { await throttle(JSON.stringify({ chat: Number(chat) })); await reddy('update', { chat: Number(chat), messageId: Number(messageId), msg: clip(text), keyboard: kb(buttons) }); return true; } catch (e) { console.warn('[Координатор] update:', e.message); return false; }
  }
  const dm = (userId, text, buttons) => (userId ? send({ user: userId }, text, buttons) : null);
  async function findUser(userKey, token) {
    try {
      const r = await reddy('finduser', null, { userkey: String(userKey).replace(/^\+/, '') }, token);
      return r.user ? { id: String(r.user.id), name: [r.user.firstName, r.user.lastName].filter(Boolean).join(' ') || r.user.userName || null } : null;
    } catch (e) { return null; }
  }

  /** Проверка, что ID — групповой чат с ботом, а не личная переписка. Возвращает текст ошибки или null. */
  async function checkGroupChat(chat, token) {
    if (!/^-?\d+$/.test(String(chat))) return 'ID чата должен состоять из цифр';
    const r = await gm('GET', `${REDDY_API}/v2/getchatinfo?chat=${encodeURIComponent(chat)}`, { token });
    if (r.status === 404) return 'Бот не состоит в этом чате или ID неверный. Добавьте бота в чат и напишите там /show_address';
    if (r.status === 200 && r.json && /private|p2p/i.test(r.json.chatType || '')) {
      return `Это личная переписка с ботом${r.json.chatTitle ? ` (${r.json.chatTitle})` : ''}, а нужен общий групповой чат. Добавьте бота в групповой чат и напишите /show_address там`;
    }
    return null; // 200 и group — ок; другие ответы не блокируют (метод может быть недоступен)
  }

  // ── Установка и обновления: файл плагина лежит на GitHub (ссылка задаётся администратором в настройках) ──
  /** Приводит ссылку GitHub к «сырой» ссылке на последнюю версию файла */
  function normalizeInstallUrl(u) {
    u = String(u || '').trim();
    let m;
    if ((m = u.match(/^https:\/\/github\.com\/([^/]+)\/([^/]+)\/(?:blob|raw)\/(.+\.user\.js)$/))) return `https://raw.githubusercontent.com/${m[1]}/${m[2]}/${m[3]}`;
    if ((m = u.match(/^(https:\/\/gist\.githubusercontent\.com\/[^/]+\/[0-9a-f]+\/raw)\/[0-9a-f]{40}\/(.+\.user\.js)$/))) return `${m[1]}/${m[2]}`; // без хэша — всегда последняя версия
    return u;
  }
  const INSTALL_RE = /^https:\/\/(raw\.githubusercontent\.com|gist\.githubusercontent\.com)\/.+\.user\.js$/;
  /** Читает версию файла плагина по ссылке (без токена бота!) */
  function fetchRemoteVersion(url) {
    return new Promise((resolve) => {
      GM_xmlhttpRequest({ method: 'GET', url: url + (url.includes('?') ? '&' : '?') + 't=' + now(), timeout: 15000,
        onload: (r) => { const m = r.status === 200 && String(r.responseText).match(/\/\/\s*@version\s+([\d.]+)/); resolve(m ? m[1] : null); },
        onerror: () => resolve(null), ontimeout: () => resolve(null) });
    });
  }
  /** Файл плагина в репозитории GitHub — отсюда Tampermonkey сам ставит обновления. В настройках можно указать другую ссылку. */
  const DEFAULT_INSTALL_URL = 'https://raw.githubusercontent.com/kadirovm-eng/Problemati_Complaints/main/confluence-coordinator.user.js';
  const installUrl = () => (cfgCache && cfgCache.installUrl) || DEFAULT_INSTALL_URL;
  const upd = { latest: null, url: null };
  async function checkUpdate(force) {
    let cfg; try { cfg = await getConfig(); } catch (e) { return; }
    if (!cfg) return;
    const url = cfg.installUrl || DEFAULT_INSTALL_URL;
    upd.url = url;
    let cache = null; try { cache = JSON.parse(localStorage.getItem('crb-upd') || 'null'); } catch (e) { /* */ }
    if (!force && cache && cache.url === url && now() - cache.at < 6 * 3600000) { upd.latest = cache.ver; return; }
    const ver = await fetchRemoteVersion(url);
    if (ver) { upd.latest = ver; try { localStorage.setItem('crb-upd', JSON.stringify({ url, at: now(), ver })); } catch (e) { /* */ } }
  }
  const updateAvailable = () => upd.latest && verNum(upd.latest) > verNum(VERSION);

  // ─────────────────────────── Тексты ───────────────────────────
  const STATUS = { new: '🆕 Ожидает проверяющего', in_progress: '🔍 На проверке', returned: '↩️ Возвращено автору с вопросом', approved: '✅ Корректность подтверждена', cancelled: '🚫 Отменено' };
  const title = (r) => `«${short(r.title, 120) || 'Без названия'}»${r.space ? ` (${r.space})` : ''}`;
  function cardText(r) {
    const l = [
      `[b]${r.urgent ? '🔥 ' : ''}Заявка #${r.id} на проверку правок${r.round > 1 ? ` · круг ${r.round}` : ''}[/b]`,
      `📄 ${title(r)}`, `✏️ Автор: ${safe(r.authorName) || '—'}`,
      `🔢 Версия ${r.version ?? '?'}${r.baseVersion ? ` (правки с версии ${r.baseVersion})` : ''}`,
    ];
    if (r.comment) l.push(`💬 ${short(r.comment)}`);
    const lk = L(r); if (lk) l.push(lk.trim());
    l.push('', `Статус: ${STATUS[r.status]}`);
    if (r.status === 'in_progress') l.push(`👤 ${safe(r.reviewerName)} · взял(а) в ${fmtT(r.taken)}`);
    if (r.status === 'returned') l.push(`👤 ${safe(r.reviewerName)}`, `❓ ${short(r.question, 300)}`);
    if (r.status === 'approved') l.push(`👤 ${safe(r.reviewerName)} · проверка ${dur(r.closed - r.taken)}`);
    return l.join('\n');
  }
  // Ссылки — текстом в сообщении (кнопки-ссылки Reddy открываются ненадёжно). Кнопки — только для действий.
  const links = () => [];
  const bbUrl = (u) => String(u || '').replace(/\]/g, '%5D').replace(/\[/g, '%5B');
  /** Ссылка на сообщение в Reddy вида https://lnk.reddy.team/message/<чат>/<сообщение> */
  const APPROVAL_RE = /^https:\/\/lnk\.reddy\.team\/message\/\d+\/\d+\/?$/;
  const L = (r) => { const p = []; if (r.url) p.push(`[url=${bbUrl(r.url)}]📄 Открыть статью[/url]`); if (r.diffUrl) p.push(`[url=${bbUrl(r.diffUrl)}]🔀 Посмотреть изменения[/url]`); if (r.approvalUrl) p.push(`[url=${bbUrl(r.approvalUrl)}]🤝 Согласование[/url]`); return p.length ? `\n${p.join('   ')}` : ''; };
  const reviewButtons = (r) => [[{ text: '✅ Подтвердить', data: `approve:${r.id}` }, { text: '↩️ Вернуть с вопросом', data: `return:${r.id}` }], ...links(r)];
  function cardButtons(r) {
    if (r.status === 'new') return [[{ text: '🙋 Взять в работу', data: `take:${r.id}` }], ...links(r)];
    if (r.status === 'in_progress') return [...reviewButtons(r).slice(0, 1), [{ text: '🔓 Отказаться', data: `release:${r.id}` }], ...links(r)];
    return links(r);
  }

  // ─────────────────────────── Логика заявок ───────────────────────────
  const OPEN = ['new', 'in_progress', 'returned'];
  const fail = (error) => ({ ok: false, error });
  const ev = (r, actor, a, d) => { r.events.push({ a, u: actor ? actor.id : null, n: actor ? uname(actor) : 'Система', t: now(), d: d || null }); r.updated = now(); };

  /** Изменить заявку атомарно. fn возвращает строку-ошибку, чтобы отменить изменение. */
  async function changeReq(id, fn) {
    let err = null;
    const r = await mutate(K.req(id), (r) => { if (!r) { err = 'Заявка не найдена'; return undefined; } const e = fn(r); if (e) { err = e; return undefined; } return r; }, null);
    if (err) return { error: err, req: r };
    await syncOpen(r);
    return { req: r };
  }
  /** Сводка открытых заявок в meta: для очереди, SLA и поиска дублей */
  function syncOpen(r) {
    return updateMeta((m) => {
      if (OPEN.includes(r.status)) {
        m.byPage[r.pageId] = r.id;
        m.open[r.id] = { s: r.status, c: r.created, st: r.stage || null, rv: r.reviewerId || null };
      } else {
        if (m.byPage[r.pageId] === r.id) delete m.byPage[r.pageId];
        delete m.open[r.id];
        for (const [u, p] of Object.entries(m.pending)) if (p.id === r.id) delete m.pending[u];
      }
      return m;
    });
  }
  const setPending = (userId, kind, id) => updateMeta((m) => { m.pending[userId] = { kind, id, at: now() }; return m; });
  const clearPending = (userId) => updateMeta((m) => { if (!m.pending[userId]) return undefined; delete m.pending[userId]; return m; });

  async function refreshCard(r, opts = {}) {
    const cfg = await getConfig();
    const ok = r.chatMsg && (await edit(cfg.chatId, r.chatMsg, cardText(r), cardButtons(r)));
    if (!ok) {
      const id = await send({ chat: cfg.chatId }, cardText(r), cardButtons(r));
      if (id) r = await mutate(K.req(r.id), (x) => { x.chatMsg = id; return x; }, null);
    }
    if (!opts.skipThread && !r.threadId && !r.threadTried) await ensureThread(await getReq(r.id));
  }
  /** Фоновые действия после отправки: автор видит «Готово» сразу, а обсуждение и уведомления уходят следом */
  function background(fn) { const pr = (async () => { try { await fn(); } catch (e) { console.warn('[Координатор] фон:', e.message); } })(); pending.add(pr); pr.finally(() => pending.delete(pr)); return pr; }
  const pending = new Set();
  /** Открыть «Обсуждение» (тред) под карточкой заявки. Нужны права бота на треды — их включает поддержка Reddy. */
  async function ensureThread(r) {
    if (!r || r.threadId || !r.chatMsg) return r;
    if (r.threadTried) return r;
    const cfg = await getConfig();
    let threadId = null, err = null;
    try { const res = await reddy('createthread', { chatId: Number(cfg.chatId), messageId: Number(r.chatMsg) }); threadId = res.threadId != null ? String(res.threadId) : null; if (!threadId) err = 'Reddy не вернул ID обсуждения'; } catch (e) { err = e.message; }
    const x = await mutate(K.req(r.id), (q) => { q.threadId = threadId; q.threadTried = true; return q; }, null);
    await updateMeta((m) => { m.threads = m.threads || {}; if (threadId) m.threads[threadId] = r.id; m.threadError = threadId ? null : err; m.threadCheckedAt = now(); return m; });
    if (err) console.warn('[Координатор] обсуждение не создано:', err);
    if (threadId) await toThread(x, `🧵 Обсуждение заявки #${x.id}. Сюда бот дублирует ход проверки, вопросы проверяющего и ответы автора.\n\nПосле кнопки «Вернуть с вопросом» вопрос автору можно написать прямо здесь.`);
    return x;
  }
  /** Написать в обсуждение заявки (если оно есть) */
  async function toThread(r, text, reply, buttons) {
    if (r && !r.threadId && r.threadTried !== true) r = await getReq(r.id); // обсуждение могло открыться только что
    if (!r || !r.threadId) return null;
    let last = '';
    for (let i = 0; i < 2; i++) {
      try {
        await throttle('thread:' + r.threadId);
        const res = await reddy('thread', { thread: Number(r.threadId), msg: clip(text), ...(reply ? { reply: Number(reply) } : {}), ...(buttons ? { keyboard: kb(buttons) } : {}) });
        if (threadSendFailed) { threadSendFailed = false; await updateMeta((m) => { if (!m.threadSendError) return undefined; m.threadSendError = null; return m; }); }
        return res.messageId != null ? String(res.messageId) : '';
      } catch (e) { last = e.message; if (i === 0) await sleep(1500); } // обсуждение могло ещё не успеть создаться
    }
    console.warn('[Координатор] не удалось написать в обсуждение:', last);
    threadSendFailed = true;
    await updateMeta((m) => { m.threadSendError = `${last} (заявка #${r.id}, ${fmtT(now())})`; return m; }).catch(() => {});
    return null;
  }
  let threadSendFailed = false;

  // ── Старшим: по каждой заявке одно сообщение в личке, которое обновляется на месте.
  //    Новое сообщение (с ❗) — только когда заявку вовремя не взяли или проверка затянулась.
  const hhmm = (ts) => new Intl.DateTimeFormat('ru-RU', { timeZone: TZ, hour: '2-digit', minute: '2-digit' }).format(new Date(ts));
  function seniorStatus(r) {
    const t = now();
    if (r.status === 'new') return r.slaTake ? `❗ [b]Никто не взял[/b] — ждёт ${dur(t - r.created)} (с ${hhmm(r.created)})` : `🆕 Ждёт проверяющего (с ${hhmm(r.created)})`;
    if (r.status === 'in_progress') return `${r.slaCheck ? '❗ [b]Проверка затянулась[/b] — ' : '🔍 '}На проверке у ${safe(r.reviewerName)} (с ${hhmm(r.stage || r.taken)})`;
    if (r.status === 'returned') return `↩️ Вопрос автору от ${safe(r.reviewerName)} — ждём ответ`;
    if (r.status === 'approved') return `✅ Подтверждено — ${safe(r.reviewerName)} · проверка ${dur(r.closed - r.taken)}, всего ${dur(r.closed - r.created)}${r.round > 1 ? `, кругов ${r.round}` : ''}`;
    return `🚫 Отменено`;
  }
  function seniorLine(e) {
    const n = safe(e.n), d = e.d ? short(e.d, 90) : '';
    const L1 = {
      created: `🆕 отправил(а) ${n}`, taken: `🔍 взял(а) ${n}${d ? ` · ${d.replace('ожидала', 'ждала')}` : ''}`, released: `🔓 ${n} отказался(ась)${d ? ` (${d})` : ''}`,
      returned: `↩️ ${n}: вопрос автору — «${d}»`, answered: `💬 ${n}: ответ — «${d}»`, resubmitted: `🔄 ${n}: доработка, ${d}`, updated: `ℹ️ ${n}: ${d}`,
      approved: `✅ подтвердил(а) ${n}${d ? ` · ${d}` : ''}`, cancelled: `🚫 отменил(а) ${n}`, sla_take: `❗ никто не взял (${d})`, sla_check: `❗ проверка затянулась (${d})`,
    }[e.a] || e.a;
    return `${hhmm(e.t)} ${L1}`;
  }
  /** Ссылка на карточку заявки в чате Reddy (формат ссылки из клиента Reddy) */
  const chatLink = (r) => (cfgCache && cfgCache.chatId && r.chatMsg ? `https://lnk.reddy.team/message/${encodeURIComponent(cfgCache.chatId)}/${encodeURIComponent(r.chatMsg)}` : null);
  function seniorText(r) {
    const evs = (r.events || []);
    const shown = evs.length > 12 ? [{ t: evs[0].t, a: 'created', n: evs[0].n }, null, ...evs.slice(-10)] : evs;
    return [
      `[b]${r.urgent ? '🔥 ' : ''}#${r.id} ${title(r)}[/b]`,
      `✏️ ${safe(r.authorName)}${r.comment ? ` · 💬 ${short(r.comment, 150)}` : ''}`,
      `Статус: ${seniorStatus(r)}`,
      '',
      ...shown.map((e) => (e ? seniorLine(e) : '…')),
    ].join('\n') + L(r) + (chatLink(r) ? `\n[url=${chatLink(r)}]💬 Карточка в чате[/url]   ` : '\n') + cabinetLink();
  }
  async function seniorList() {
    const cfg = await getConfig();
    if (cfg.seniorDm === false) return [];
    return (await listProps('crb-u-')).filter((u) => isSenior(u) && u.notify !== false && u.active !== false);
  }
  /** Обновить сводку заявки у всех старших. remind — текст напоминания с ❗ (придёт отдельным сообщением в ответ на сводку). */
  async function seniorsView(rIn, remind) {
    const r = (await getReq(rIn.id)) || rIn; // берём самое свежее состояние
    await getConfig(); // нужен ID чата для ссылки на карточку
    const seniors = await seniorList();
    const text = seniorText(r), views = r.seniorMsgs || {}, fresh = {};
    await Promise.all(seniors.map(async (u) => {
      const v = views[u.id];
      if (v && (await edit(v.chat, v.msg, text))) { if (remind) await send({ user: u.id }, remind, null, { reply: v.msg }); return; }
      if (!OPEN.includes(r.status) && !v) return; // закрытые заявки новым старшим не присылаем
      const sent = await sendFull({ user: u.id }, remind ? `${remind}\n\n${text}` : text);
      if (sent && sent.id) fresh[u.id] = { chat: sent.chat, msg: sent.id };
    }));
    if (Object.keys(fresh).length) await mutate(K.req(r.id), (x) => { x.seniorMsgs = { ...(x.seniorMsgs || {}), ...fresh }; return x; }, null);
  }

  /** Сообщение всем старшим (используется только для ежедневной сводки) */
  async function notifySeniors(text, exclude, buttons) {
    if ((await getConfig()).seniorDm === false) return; // всё видно в обсуждении заявки в чате
    const users = await listProps('crb-u-');
    for (const u of users) if (isSenior(u) && u.notify !== false && u.active !== false && u.id !== String(exclude || '')) await dm(u.id, text, buttons);
  }

  async function submit(author, p) {
    const version = Number(p.version) || null;
    const prev = Number(p.prevVersion) || (version ? version - 1 : null);
    const diff = (from, to) => (from && to && from < to ? `${CB}/pages/diffpagesbyversion.action?pageId=${encodeURIComponent(p.pageId)}&selectedPageVersions=${from}&selectedPageVersions=${to}` : null);
    const m = await getMeta();
    const existing = m.byPage[p.pageId] ? await getReq(m.byPage[p.pageId]) : null;

    if (existing && OPEN.includes(existing.status)) {
      const wasReturned = existing.status === 'returned';
      const { req: r, error } = await changeReq(existing.id, (r) => {
        r.version = version; r.diffUrl = diff(r.baseVersion, version) || r.diffUrl; r.title = p.title || r.title;
        if (p.comment) r.comment = p.comment;
        if (p.approvalUrl) r.approvalUrl = p.approvalUrl;
        if (r.status === 'returned') {
          r.status = r.reviewerId ? 'in_progress' : 'new'; r.round += 1; r.question = null; r.stage = now(); r.slaCheck = false;
          ev(r, author, 'resubmitted', `версия ${version}${p.comment ? `: ${p.comment}` : ''}`);
        } else ev(r, author, 'updated', `версия ${version}`);
      });
      if (error) return fail(error);
      await refreshCard(r, { skipThread: true });
      const after = background(async () => {
        const rr = await ensureThread(await getReq(r.id)); // у заявки, созданной до выдачи прав, обсуждение откроется сейчас
        if (wasReturned) {
          await clearPending(rr.authorId);
          await Promise.all([
            rr.reviewerId && dm(rr.reviewerId, `🔄 ${safe(rr.authorName)} доработал(а) заявку #${rr.id} ${title(rr)} и отправил(а) повторно (версия ${version}).${p.comment ? `\n💬 ${short(p.comment)}` : ''}${L(rr)}`, reviewButtons(rr)),
            toThread(rr, `🔄 ${safe(rr.authorName)} отправил(а) доработку — версия ${version}, круг ${rr.round}.${p.comment ? `\n💬 ${short(p.comment)}` : ''}${L(rr)}`),
            seniorsView(rr),
          ]);
        } else {
          await Promise.all([
            toThread(rr, `ℹ️ ${safe(rr.authorName)} обновил(а) статью — версия ${version}.${p.comment ? `\n💬 ${short(p.comment)}` : ''}${L(rr)}`),
            seniorsView(rr),
            rr.status === 'in_progress' && dm(rr.reviewerId, `ℹ️ В заявке #${rr.id} ${title(rr)} новая версия ${version} — проверьте актуальные изменения.${L(rr)}`, links(rr)),
          ]);
        }
      });
      return { ok: true, req: r, action: wasReturned ? 'resubmitted' : 'updated', after };
    }

    // Новая заявка. Быстрый путь — только то, что нужно, чтобы карточка появилась в чате:
    // 1) номер заявки и отметка «по статье есть открытая заявка» — одной записью в meta;
    const t = now();
    let id;
    await updateMeta((mm) => { id = mm.nextId; mm.nextId += 1; mm.byPage[String(p.pageId)] = id; mm.open[id] = { s: 'new', c: t, st: null, rv: null }; return mm; });
    const r = {
      id, pageId: String(p.pageId), title: p.title, url: p.url, space: p.space, baseVersion: prev, version, diffUrl: diff(prev, version),
      authorId: author.id, authorName: uname(author, p.authorName), authorLogin: p.authorLogin || null, comment: p.comment || null, urgent: !!p.urgent, approvalUrl: p.approvalUrl || null,
      status: 'new', reviewerId: null, reviewerName: null, question: null, round: 1, chatMsg: null,
      created: t, taken: null, stage: null, closed: null, updated: t, slaTake: false, slaCheck: false, events: [],
    };
    ev(r, author, 'created', p.comment);
    // 2) карточка в чат; 3) заявка сохраняется сразу с номером сообщения карточки.
    const cfg = await getConfig();
    r.chatMsg = await send({ chat: cfg.chatId }, cardText(r), cardButtons(r));
    await putProp(K.req(id), r, null);
    // Остальное — в фоне: обсуждение под карточкой и уведомления старшим
    const after = background(() => Promise.all([
      ensureThread(r),
      seniorsView(r),
    ]));
    return { ok: true, req: r, action: 'created', after, cardFailed: !r.chatMsg };
  }

  /** Проверять заявки (брать, подтверждать, возвращать, отказываться) могут только старшие и администраторы */
  const REVIEW_DENIED = 'Брать и проверять заявки могут только старшие и администраторы. Ваши заявки — в меню «📋 Мои заявки».';
  async function take(id, actor) {
    if (!isSenior(actor)) return fail(REVIEW_DENIED);
    const { req: r, error } = await changeReq(id, (r) => {
      if (r.status === 'in_progress') return r.reviewerId === actor.id ? 'Заявка уже у вас' : `Заявку уже взял(а) ${r.reviewerName}`;
      if (r.status !== 'new') return `Заявка в статусе «${STATUS[r.status]}»`;
      Object.assign(r, { status: 'in_progress', reviewerId: actor.id, reviewerName: uname(actor), taken: now(), stage: now() });
      ev(r, actor, 'taken', `ожидала ${dur(now() - r.created)}`);
    });
    if (error) return fail(error);
    await refreshCard(r);
    await toThread(r, `🔍 ${safe(r.reviewerName)} взял(а) в работу · заявка ждала ${dur(r.taken - r.created)}`);
    await dm(actor.id, `Вы взяли в работу заявку #${r.id} ${title(r)}.\nПо итогам проверки нажмите кнопку:${L(r)}`, reviewButtons(r));
    if (r.authorId && r.authorId !== actor.id) await dm(r.authorId, `🔍 Вашу заявку #${r.id} ${title(r)} взял(а) в работу ${safe(r.reviewerName)}.`);
    await seniorsView(r);
    return { ok: true, req: r, message: `Заявка #${r.id} ваша` };
  }
  const checkReviewer = (r, actor) => (!isSenior(actor) ? REVIEW_DENIED : r.status !== 'in_progress' ? `Заявка в статусе «${STATUS[r.status]}»` : r.reviewerId !== actor.id && !isSenior(actor) ? `Заявку проверяет ${r.reviewerName}` : null);

  async function release(id, actor) {
    const { req: r, error } = await changeReq(id, (r) => {
      const e = checkReviewer(r, actor); if (e) return e;
      ev(r, actor, 'released', r.reviewerId !== actor.id ? `снято с ${r.reviewerName}` : null);
      Object.assign(r, { status: 'new', reviewerId: null, reviewerName: null, taken: null, stage: null, slaTake: false, slaCheck: false });
    });
    if (error) return fail(error);
    await refreshCard(r);
    await toThread(r, `🔓 ${safe(uname(actor))} вернул(а) заявку в очередь — нужен новый проверяющий`);
    await seniorsView(r);
    return { ok: true, req: r, message: 'Заявка возвращена в очередь' };
  }
  async function approve(id, actor, comment) {
    const { req: r, error } = await changeReq(id, (r) => {
      const e = checkReviewer(r, actor); if (e) return e;
      r.status = 'approved'; r.closed = now();
      ev(r, actor, 'approved', [`проверка ${dur(r.closed - r.taken)}`, comment].filter(Boolean).join('; '));
    });
    if (error) return fail(error);
    await refreshCard(r);
    await toThread(r, `✅ ${safe(r.reviewerName)} подтвердил(а) корректность правок · проверка ${dur(r.closed - r.taken)}${comment ? `\n💬 ${short(comment)}` : ''}`);
    if (r.authorId) await dm(r.authorId, `✅ Правки в статье ${title(r)} проверены — всё корректно (заявка #${r.id}, проверял(а) ${safe(r.reviewerName)}).${comment ? `\n💬 ${short(comment)}` : ''}${L(r)}`, links(r));
    await seniorsView(r);
    return { ok: true, req: r, message: `Заявка #${r.id} подтверждена` };
  }
  async function startReturn(id, actor) {
    const r = await getReq(id);
    if (!r) return fail('Заявка не найдена');
    const e = checkReviewer(r, actor); if (e) return fail(e);
    await setPending(actor.id, 'return_question', r.id);
    if (r.threadId) {
      await toThread(r, `✍️ ${safe(uname(actor))}, напишите вопрос или замечание автору следующим сообщением здесь, в обсуждении (или боту в личку).`, null, [[{ text: '✖️ Передумал(а)', data: `cancelinput:${r.id}` }]]);
      return { ok: true, message: 'Напишите вопрос в обсуждении заявки' };
    }
    await dm(actor.id, `✍️ Напишите одним сообщением вопрос или замечание автору по заявке #${r.id} ${title(r)}.`, [[{ text: '✖️ Передумал(а)', data: `cancelinput:${r.id}` }]]);
    return { ok: true, message: 'Напишите вопрос боту в личные сообщения' };
  }
  /** src.threadMsg — ID сообщения в обсуждении, если текст написан там (тогда не повторяем его, а отвечаем на него) */
  async function returnWithQuestion(id, actor, question, src = {}) {
    question = String(question || '').trim();
    if (!question) return fail('Вопрос не может быть пустым');
    const { req: r, error } = await changeReq(id, (r) => {
      const e = checkReviewer(r, actor); if (e) return e;
      r.status = 'returned'; r.question = question;
      ev(r, actor, 'returned', question);
    });
    if (error) return fail(error);
    await clearPending(actor.id);
    await refreshCard(r);
    if (src.threadMsg) await toThread(r, `↩️ Заявка возвращена автору (${safe(r.authorName)}) с этим вопросом. Ответ автора появится здесь.`, src.threadMsg);
    else await toThread(r, `↩️ ${safe(r.reviewerName)} вернул(а) заявку автору с вопросом:\n❓ ${short(question, 2000)}`);
    if (r.authorId) {
      await setPending(r.authorId, 'author_answer', r.id);
      await dm(r.authorId, `↩️ ${safe(r.reviewerName)} вернул(а) заявку #${r.id} ${title(r)} с вопросом:\n\n❓ ${short(question)}\n\nОтветьте сообщением сюда — ответ уйдёт проверяющему. Если нужно поправить статью — внесите правки и снова нажмите «На проверку» в Confluence.${L(r)}`,
        [[{ text: '💬 Ответить', data: `answer:${r.id}` }], ...links(r)]);
    }
    await seniorsView(r);
    return { ok: true, req: r, message: 'Вопрос отправлен автору' };
  }
  async function startAnswer(id, actor) {
    const r = await getReq(id);
    if (!r) return fail('Заявка не найдена');
    if (r.authorId !== actor.id) return fail('Отвечать может только автор');
    if (r.status !== 'returned') return fail(`Заявка в статусе «${STATUS[r.status]}»`);
    await setPending(actor.id, 'author_answer', r.id);
    await dm(actor.id, `✍️ Напишите ответ проверяющему по заявке #${r.id} одним сообщением.`, [[{ text: '✖️ Передумал(а)', data: `cancelinput:${r.id}` }]]);
    return { ok: true, message: 'Напишите ответ боту' };
  }
  async function authorAnswer(id, actor, answer, src = {}) {
    answer = String(answer || '').trim();
    if (!answer) return fail('Ответ пустой');
    let question = '';
    const { req: r, error } = await changeReq(id, (r) => {
      if (r.authorId !== actor.id) return 'Отвечать может только автор';
      if (r.status !== 'returned') return `Заявка в статусе «${STATUS[r.status]}»`;
      question = r.question;
      Object.assign(r, { status: 'in_progress', round: r.round + 1, question: null, stage: now(), slaCheck: false });
      ev(r, actor, 'answered', answer);
    });
    if (error) return fail(error);
    await clearPending(actor.id);
    await refreshCard(r);
    if (src.threadMsg) await toThread(r, `💬 Ответ передан ${safe(r.reviewerName)}, заявка снова на проверке (круг ${r.round}).`, src.threadMsg);
    else await toThread(r, `💬 ${safe(r.authorName)} ответил(а):\n${short(answer, 2000)}`);
    await dm(r.reviewerId, `💬 ${safe(r.authorName)} ответил(а) по заявке #${r.id} ${title(r)}:\n\n${short(answer, 2000)}\n\nВаш вопрос был: ${short(question, 300)}${L(r)}`, reviewButtons(r));
    await seniorsView(r);
    return { ok: true, req: r, message: 'Ответ передан проверяющему' };
  }
  async function cancel(id, actor) {
    const { req: r, error } = await changeReq(id, (r) => {
      if (r.authorId !== actor.id && !isSenior(actor)) return 'Отменить может автор или старший';
      if (!OPEN.includes(r.status)) return 'Заявка уже закрыта';
      r.status = 'cancelled'; r.closed = now(); ev(r, actor, 'cancelled');
    });
    if (error) return fail(error);
    await refreshCard(r);
    await toThread(r, `🚫 ${safe(uname(actor))} отменил(а) заявку`);
    if (r.reviewerId && r.reviewerId !== actor.id) await dm(r.reviewerId, `🚫 Заявка #${r.id} ${title(r)} отменена (${safe(uname(actor))}).`);
    await seniorsView(r);
    return { ok: true, req: r, message: 'Заявка отменена' };
  }

  // ─────────────────────────── Статистика ───────────────────────────
  function teamStats(reqs, users, from, to = now()) {
    const s = {};
    const row = (id) => (s[id] = s[id] || { id, name: (users.find((u) => u.id === id) || {}).name || `ID ${id}`, taken: 0, approved: 0, returned: 0, released: 0, inWork: 0, checkSum: 0, checkN: 0, waitSum: 0, waitN: 0, authored: 0 });
    for (const r of reqs) {
      if (r.created >= from && r.created <= to && r.authorId) row(r.authorId).authored++;
      if (r.status === 'in_progress' && r.reviewerId) row(r.reviewerId).inWork++;
      for (const e of r.events || []) {
        if (!e.u || e.t < from || e.t > to) continue;
        if (e.a === 'taken') { const x = row(e.u); x.taken++; }
        if (e.a === 'approved') { const x = row(e.u); x.approved++; if (r.taken) { x.checkSum += r.closed - r.taken; x.checkN++; } }
        if (e.a === 'returned') row(e.u).returned++;
        if (e.a === 'released') row(e.u).released++;
      }
      if (r.taken && r.taken >= from && r.taken <= to && r.reviewerId) { const x = row(r.reviewerId); x.waitSum += r.taken - r.created; x.waitN++; }
    }
    return Object.values(s).map((x) => ({ ...x, avgCheck: x.checkN ? x.checkSum / x.checkN : null, avgWait: x.waitN ? x.waitSum / x.waitN : null }));
  }
  function statsText(reqs, users, from, label) {
    const open = reqs.filter((r) => OPEN.includes(r.status));
    const q = { n: open.filter((r) => r.status === 'new'), p: open.filter((r) => r.status === 'in_progress').length, r: open.filter((r) => r.status === 'returned').length };
    const lines = [`📊 Сводка ${label}`, `Сейчас: в очереди ${q.n.length}, на проверке ${q.p}, у авторов с вопросами ${q.r}`];
    if (q.n.length) lines.push(`Самая старая в очереди ждёт ${dur(now() - Math.min(...q.n.map((r) => r.created)))}`);
    const rows = teamStats(reqs, users, from).filter((x) => x.taken || x.approved || x.returned || x.inWork).sort((a, b) => b.approved - a.approved);
    lines.push('');
    if (!rows.length) lines.push('Проверок за период не было.');
    for (const x of rows) lines.push(`👤 ${safe(x.name)}: взял ${x.taken}, подтвердил ${x.approved}, вернул ${x.returned}${x.released ? `, отказался ${x.released}` : ''}${x.inWork ? `, в работе ${x.inWork}` : ''}${x.avgCheck ? ` · ср. проверка ${dur(x.avgCheck)}` : ''}`);
    return lines.join('\n');
  }

  // ─────────────────────────── Бот: кнопки и сообщения ───────────────────────────

  const listLine = (r) => `#${r.id}${r.urgent ? ' 🔥' : ''} ${title(r)} — ${STATUS[r.status]}${r.status === 'new' ? ` · ждёт ${dur(now() - r.created)}` : ''}${r.status === 'in_progress' ? ` · ${safe(r.reviewerName)}` : ''}`;

  /** Обрабатывает обновление от Reddy и возвращает короткое описание действия — оно попадает в журнал бота */
  async function handleUpdate(u) {
    log('update', u);
    const usr = u.user || {};
    if (!usr.id) return 'пропущено: нет пользователя';
    const thr = u.thread ? (u.threadMessage || u.message) : null; // сообщение в обсуждении: type threadMessage или message с полем thread
    const kind = u.type === 'buttonAction' && u.button ? 'button' : thr && (u.type === 'threadMessage' || u.type === 'message') ? 'thread' : u.type === 'message' && u.message ? 'message' : null;
    if (!kind) return `пропущено: тип ${u.type}`;
    const name = [usr.firstName, usr.lastName].filter(Boolean).join(' ') || usr.userName || null;
    const actor = await touchUser(String(usr.id), name, usr.key);
    if (actor.active === false) { await dm(actor.id, 'Ваш доступ отключён администратором.'); return 'доступ отключён'; }

    if (kind === 'thread') return handleThreadMessage(u, thr, actor);
    if (kind === 'button') {
      const [action, raw] = String(u.button.data || '').split(':');
      const id = Number(raw);
      if (action === 'menu') { await showMenu(raw || 'home', actor, u); return `меню: ${raw}`; }
      if (action === 'cancelinput') {
        const m = await getMeta(); const had = m.pending[actor.id];
        await clearPending(actor.id);
        await dm(actor.id, had ? '👌 Отменено — заявка осталась в прежнем статусе.' : 'Нечего отменять.');
        return had ? 'ввод отменён' : 'нечего отменять';
      }
      const fn = { take, approve, release, return: startReturn, answer: startAnswer, cancel }[action];
      const res = fn ? await fn(id, actor) : fail('Неизвестная кнопка');
      if (!res.ok) await dm(actor.id, `⚠️ ${res.error}`);
      return res.ok ? res.message : `отказ: ${res.error}`;
    }

    const isPrivate = !u.chat || u.chat.p2p !== false;
    const text = String(u.message.msg || '').trim().replace(/^@\S+[\s,]*/, '');
    const cm = text.match(/^(\/[a-zA-Zа-яА-Я_]+)(?:@\S+)?\s*(.*)$/s);
    if (cm) { await command(cm[1].toLowerCase(), cm[2], actor, isPrivate); return `команда ${cm[1]}`; }
    if (!isPrivate) return 'сообщение в общем чате — пропущено';
    const m = await getMeta();
    const pend = m.pending[actor.id];
    if (pend) {
      const res = pend.kind === 'return_question' ? await returnWithQuestion(pend.id, actor, text) : await authorAnswer(pend.id, actor, text);
      if (!res.ok) await clearPending(actor.id);
      await dm(actor.id, res.ok ? `👍 ${res.message}` : `⚠️ ${res.error}`);
      return res.ok ? `личка: ${res.message}` : `личка: отказ — ${res.error}`;
    }
    await showMenu('home', actor);
    return 'личка: бот не ждал ответа — показано меню';
  }

  /** Сообщение в обсуждении заявки: вопрос проверяющего (после «Вернуть с вопросом») или ответ автора. Остальное — обычная переписка, бот молчит. */
  async function handleThreadMessage(u, thr, actor) {
    const text = String(thr.msg || '').trim().replace(/^@\S+[\s,]*/, '');
    if (!text) return 'обсуждение: пустое сообщение (файл/стикер) — пропущено';
    if (text.startsWith('/')) return 'обсуждение: команда — пропущено';
    const m = await getMeta();
    const id = (m.threads || {})[String(u.thread.id)];
    if (!id) return `обсуждение ${u.thread.id}: не относится к заявкам — пропущено`;
    const r = await getReq(id);
    if (!r) return `обсуждение: заявка #${id} не найдена`;
    const src = { threadMsg: thr.id };
    const pend = m.pending[actor.id];
    let res;
    if (pend && pend.kind === 'return_question' && Number(pend.id) === r.id) res = await returnWithQuestion(r.id, actor, text, src);
    else if (r.status === 'returned' && r.authorId === actor.id) res = await authorAnswer(r.id, actor, text, src);
    else return `обсуждение #${r.id}: обычная переписка (статус «${RU_ST[r.status]}», ${r.authorId === actor.id ? 'автор' : 'не автор'}) — бот молчит`;
    if (!res.ok) { await clearPending(actor.id); await toThread(r, `⚠️ ${res.error}`, thr.id); return `обсуждение #${r.id}: отказ — ${res.error}`; }
    return `обсуждение #${r.id}: ${res.message}`;
  }
  const RU_ST = { new: 'ожидает', in_progress: 'на проверке', returned: 'вопрос автору', approved: 'подтверждено', cancelled: 'отменено' };

  // ── Журнал бота: последние 50 входящих событий и что с ними сделано (виден администратору в настройках) ──
  async function journal(u, note) {
    const usr = u.user || {};
    const body = u.button ? `кнопка ${u.button.data}` : (u.threadMessage || u.message || {}).msg || '';
    const item = { t: now(), type: u.type + (u.thread ? ' (обсуждение)' : ''), who: [usr.firstName, usr.lastName].filter(Boolean).join(' ') || String(usr.id || ''), text: String(body).slice(0, 120), note: String(note || '').slice(0, 300), v: VERSION };
    try { await mutate('crb-log', (l) => { l.items.unshift(item); l.items = l.items.slice(0, 50); return l; }, { items: [] }); } catch (e) { /* журнал не критичен */ }
  }

  // ── Меню на кнопках: одно сообщение, которое меняется при нажатии кнопок ──
  const back = [{ text: '⬅️ Меню', data: 'menu:home' }];
  /** Кабинет — это служебная страница Confluence: плагин открывает его, когда на неё заходят */
  const cabinetUrl = () => `${CB}/pages/viewpage.action?pageId=${encodeURIComponent(STORAGE_PAGE_ID)}#crb-cabinet`; // #crb-cabinet — открыть кабинет сразу
  const cabinetLink = () => `[url=${bbUrl(cabinetUrl())}]🗂 Открыть личный кабинет[/url]`;
  /** Для тех, кто ещё не подключил плагин: прямая ссылка на установку (если задана) и страница с инструкцией */
  const installLines = () => `🧩 [b]Плагин ещё не подключён[/b]\n1. Установите расширение Tampermonkey в Chrome\n2. [url=${bbUrl(installUrl())}]Установите плагин[/url] — Tampermonkey предложит «Установить»\n3. [url=${bbUrl(cabinetUrl().replace('#crb-cabinet', ''))}]Откройте страницу Координатора[/url] и привяжите Reddy`;
  function menuButtons(actor) {
    const rows = [isSenior(actor) ? [{ text: '📋 Мои заявки', data: 'menu:my' }, { text: '🆕 Очередь', data: 'menu:queue' }] : [{ text: '📋 Мои заявки', data: 'menu:my' }]];
    if (isSenior(actor)) {
      rows.push([{ text: '📊 Сводка за сегодня', data: 'menu:today' }, { text: '📊 За 7 дней', data: 'menu:week' }]);
      rows.push([{ text: actor.notify !== false ? '🔕 Выключить уведомления' : '🔔 Включить уведомления', data: 'menu:notify' }]);
    }
    rows.push([{ text: '🪪 Мой номер для плагина', data: 'menu:me' }]);
    return rows;
  }
  async function menuView(sub, actor) {
    const me = actor.id;
    await getConfig(true); // меню открывают редко — берём свежие настройки (ссылка на установку и т.п.)
    if (sub === 'notify' && isSenior(actor)) { actor = await saveUser(me, (u) => { u.notify = u.notify === false; return u; }); sub = 'home'; }
    if (sub === 'me') return { text: `🪪 Ваш номер в Reddy: [b]${actor.key || 'смотрите в своём профиле Reddy'}[/b]\n\nЕго вводят в плагине Confluence (кнопка 🗂) при первой привязке — после этого бот присылает код подтверждения сюда.\n\n${cabinetLink()}`, buttons: [back] };
    if (sub === 'today' || sub === 'week') {
      if (!isSenior(actor)) return { text: 'Сводка доступна старшим.', buttons: [back] };
      const reqs = await listProps('crb-r-');
      const week = sub === 'week';
      return { text: statsText(reqs, await listProps('crb-u-'), week ? now() - 7 * 864e5 : startOfDay(), week ? 'за 7 дней' : 'за сегодня'), buttons: [[{ text: '🔄 Обновить', data: `menu:${sub}` }, ...back]] };
    }
    if (sub === 'my') {
      const reqs = await listProps('crb-r-');
      const rev = isSenior(actor) ? reqs.filter((r) => r.reviewerId === me && r.status === 'in_progress') : [];
      const mine = reqs.filter((r) => r.authorId === me && OPEN.includes(r.status));
      const ask = mine.filter((r) => r.status === 'returned');
      return {
        text: (isSenior(actor) ? `[b]🔍 Проверяю[/b]\n${rev.map(listLine).join('\n') || 'ничего'}\n\n` : '')
          + `[b]✏️ Мои заявки[/b]\n${mine.map(listLine).join('\n') || 'нет открытых'}`
          + (ask.length ? `\n\n❓ Ждут вашего ответа: ${ask.map((r) => '#' + r.id).join(', ')}` : ''),
        buttons: [...rev.slice(0, 5).map((r) => [{ text: `✅ Подтвердить #${r.id}`, data: `approve:${r.id}` }, { text: `↩️ Вопрос по #${r.id}`, data: `return:${r.id}` }]),
          ...ask.slice(0, 5).map((r) => [{ text: `💬 Ответить по #${r.id}`, data: `answer:${r.id}` }]), [{ text: '🔄 Обновить', data: 'menu:my' }, ...back]],
      };
    }
    if (sub === 'queue') {
      if (!isSenior(actor)) return { text: 'Очередь заявок видят старшие и администраторы.', buttons: [back] };
      const q = (await listProps('crb-r-')).filter((r) => r.status === 'new').sort((a, b) => b.urgent - a.urgent || a.id - b.id);
      return { text: `[b]🆕 Очередь[/b]\n${q.slice(0, 15).map(listLine).join('\n') || 'пусто 🎉'}`, buttons: [...q.slice(0, 8).map((r) => [{ text: `🙋 Взять #${r.id}`, data: `take:${r.id}` }]), [{ text: '🔄 Обновить', data: 'menu:queue' }, ...back]] };
    }
    return {
      text: `[b]Confluence Координатор[/b] — бот проверки правок в статьях.\n\nЗаявки отправляются кнопкой «На проверку правок» в Confluence. Здесь — ${isSenior(actor) ? 'ваши заявки и очередь' : 'ваши заявки и их статус'}. Писать боту ничего не нужно, кроме случаев, когда он сам попросит вопрос или ответ.\n\n${actor.confluence ? cabinetLink() : installLines()}${isSenior(actor) ? `\n\nУведомления старшего: ${actor.notify !== false ? '🔔 включены' : '🔕 выключены'}` : ''}`,
      buttons: menuButtons(actor),
    };
  }
  /** Показать раздел меню: в личке — заменить текст того же сообщения, иначе прислать новое */
  async function showMenu(sub, actor, u) {
    const v = await menuView(sub, actor);
    const inDm = u && u.chat && u.chat.p2p !== false && u.button && u.button.messageId;
    if (inDm && (await edit(u.chat.id, u.button.messageId, v.text, v.buttons))) return;
    await dm(actor.id, v.text, v.buttons);
  }

  async function command(cmd, arg, actor, isPrivate) {
    const me = actor.id;
    if (!isPrivate && !['/queue', '/stats', '/my', '/week'].includes(cmd)) return;
    if (cmd === '/start' || cmd === '/help' || cmd === '/menu') return showMenu('home', actor);
    if (cmd === '/id') return showMenu('me', actor);
    if (cmd === '/mute' || cmd === '/unmute') { await saveUser(me, (u) => { u.notify = cmd === '/unmute'; return u; }); return dm(me, cmd === '/mute' ? '🔕 Уведомления старшего выключены' : '🔔 Уведомления старшего включены'); }
    if (cmd === '/cancel') { const m = await getMeta(); await clearPending(me); return dm(me, m.pending[me] ? 'Ок, отменено. Заявка осталась в прежнем статусе.' : 'Нечего отменять.'); }
    // Команды из меню Reddy открывают те же разделы, что и кнопки
    if (cmd === '/my') return showMenu('my', actor);
    if (cmd === '/queue') return showMenu('queue', actor);
    if (cmd === '/stats') return showMenu(/week|нед/i.test(arg) ? 'week' : 'today', actor);
    if (cmd === '/week') return showMenu('week', actor);
    if (cmd === '/cabinet') return showMenu('me', actor);
    return showMenu('home', actor);
  }

  // ─────────────────────────── Контроль сроков и сводка (делает дежурная вкладка) ───────────────────────────
  async function tick() {
    const cfg = await getConfig();
    let due = false;
    await updateMeta((m) => { if (now() - (m.lastTick || 0) < 55000) return undefined; m.lastTick = now(); due = true; return m; });
    if (!due) return;
    const m = await getMeta();
    const takeMs = (cfg.slaTake ?? 30) * 60000, checkMs = (cfg.slaCheck ?? 240) * 60000;
    for (const [id, o] of Object.entries(m.open)) {
      if (o.s === 'new' && now() - o.c > takeMs) {
        const { req: r, error } = await changeReq(id, (r) => { if (r.status !== 'new' || r.slaTake) return 'skip'; r.slaTake = true; ev(r, null, 'sla_take', `не взята ${dur(now() - r.created)}`); });
        if (!error) await toThread(r, `❗ Заявку никто не взял уже ${dur(now() - r.created)}`);
        if (!error) await seniorsView(r, `❗ Заявку #${r.id}${r.urgent ? ' 🔥' : ''} ${title(r)} никто не взял уже ${dur(now() - r.created)}`);
      }
      if (o.s === 'in_progress' && o.st && now() - o.st > checkMs) {
        const { req: r, error } = await changeReq(id, (r) => { if (r.status !== 'in_progress' || r.slaCheck) return 'skip'; r.slaCheck = true; ev(r, null, 'sla_check', `на проверке ${dur(now() - r.stage)}`); });
        if (!error) await toThread(r, `❗ Проверка идёт уже ${dur(now() - r.stage)}`);
        if (!error) await seniorsView(r, `❗ ${safe(r.reviewerName)} проверяет #${r.id} ${title(r)} уже ${dur(now() - r.stage)}`);
      }
    }
    const hour = +tzParts(now()).hour;
    if ((cfg.digestHour ?? 18) >= 0 && hour === (cfg.digestHour ?? 18) && m.lastDigest !== dayKey()) {
      let mine = false;
      await updateMeta((mm) => { if (mm.lastDigest === dayKey()) return undefined; mm.lastDigest = dayKey(); mine = true; return mm; });
      if (mine) await notifySeniors(statsText(await listProps('crb-r-'), await listProps('crb-u-'), startOfDay(), 'за сегодня'));
    }
  }

  // ─────────────────────────── Дежурная вкладка: получение обновлений ───────────────────────────
  const dutyState = { role: 'idle', since: 0, lastError: '' };
  async function dutyLoop() {
    // В одном браузере опрашивает только одна вкладка
    setInterval(() => { if (dutyState.role === 'duty') GM_setValue('crb_duty', { tab: TAB, ts: now() }); }, 5000);
    window.addEventListener('beforeunload', () => { const l = GM_getValue('crb_duty'); if (l && l.tab === TAB) GM_setValue('crb_duty', null); });
    for (;;) {
      let cfg;
      try { cfg = await getConfig(); } catch (e) { cfg = null; }
      if (!cfg || !cfg.token) { await sleep(60000); continue; }
      // Вкладка, открытая до обновления плагина, не должна забирать сообщения у бота
      try {
        const mv = (await getMeta()).version;
        if (verNum(mv) > verNum(VERSION)) { dutyState.role = 'outdated'; dutyState.lastError = `в этой вкладке старая версия плагина (${VERSION}), установлена ${mv} — обновите страницу (F5)`; await sleep(60000); continue; }
        if (verNum(mv) < verNum(VERSION)) await updateMeta((m) => { if (verNum(m.version) >= verNum(VERSION)) return undefined; m.version = VERSION; return m; });
      } catch (e) { /* хранилище недоступно — попробуем позже */ }
      const lock = GM_getValue('crb_duty');
      if (lock && lock.tab !== TAB && now() - lock.ts < 15000) { dutyState.role = 'standby'; await sleep(5000 + Math.random() * 3000); continue; } // в этом браузере дежурит другая вкладка
      GM_setValue('crb_duty', { tab: TAB, ts: now() });
      const r = await gm('GET', `${REDDY_API}/v2/getupdate`, { token: cfg.token, timeout: 100000 });
      if (r.status === 409) { // дежурит другой пользователь
        dutyState.role = 'standby'; GM_setValue('crb_duty', { tab: TAB, ts: now() }); // не даём соседним вкладкам зря стучаться
        await sleep(8000 + Math.random() * 7000); continue; // после перехода дежурного на другую страницу подхватываем быстро
      }
      if (r.status === -1) continue; // длинный запрос истёк — просто повторяем
      if (r.status !== 200) { dutyState.role = 'error'; dutyState.lastError = `Reddy ${r.status}`; await sleep(15000); continue; }
      if (r.json && !Array.isArray(r.json) && (r.json.Code === 27 || r.json.code === 27)) { dutyState.role = 'error'; dutyState.lastError = 'У бота настроен вебхук в BotMother — его нужно убрать'; await sleep(60000); continue; }
      if (dutyState.role !== 'duty') { dutyState.role = 'duty'; dutyState.since = now(); log('эта вкладка дежурит'); }
      const list = Array.isArray(r.json) ? r.json : (r.json && (r.json.updates || r.json.result)) || [];
      for (const upd of list) {
        let note;
        try { note = await handleUpdate(upd); } catch (e) { note = 'ОШИБКА: ' + e.message; console.error('[Координатор] ошибка обработки:', e); }
        await journal(upd, note);
      }
      try { await tick(); } catch (e) { console.warn('[Координатор] tick:', e.message); }
    }
  }

  // ─────────────────────────── Текущий пользователь ───────────────────────────
  let confluenceMe = null;
  const meLocal = () => GM_getValue('me', null); // { id, name, key } — привязанный аккаунт Reddy
  async function meUser() { const m = meLocal(); return m ? (await getUser(m.id)) || null : null; }

  // ─────────────────────────── Интерфейс (Shadow DOM, чтобы не конфликтовать со стилями Confluence) ───────────────────────────
  const host = document.createElement('div');
  host.id = 'crb-host';
  host.style.cssText = 'all:initial;position:fixed;z-index:2147483000;top:0;left:0;width:0;height:0';
  document.body.appendChild(host);
  const root = host.attachShadow({ mode: 'open' });
  root.innerHTML = `<style>
  :host{--bg:#f5f6f8;--panel:#fff;--ink:#1c2330;--muted:#6b7480;--line:#e3e6eb;--soft:#eef1f5;--accent:#2f5bea;--ok:#1f8a4c;--warn:#b7791f;--bad:#c2410c}
  *{box-sizing:border-box;font-family:-apple-system,"Segoe UI",Roboto,Arial,sans-serif}
  .fab{position:fixed;right:22px;bottom:84px;display:flex;gap:8px;align-items:center}
  .btn{border:1px solid var(--line);background:var(--panel);color:var(--ink);border-radius:8px;padding:8px 13px;font-size:14px;cursor:pointer;line-height:1.2}
  .btn:hover{border-color:var(--muted)} .btn:disabled{opacity:.5;cursor:default}
  .primary{background:var(--accent);border-color:var(--accent);color:#fff;font-weight:600}
  .ghost{border-color:transparent;background:transparent}
  .pill{border-radius:24px;padding:10px 16px;box-shadow:0 4px 14px rgba(0,0,0,.18);font-weight:600}
  .pill.state{background:#fff;color:var(--ink);border:1px solid #d5d9e0}
  .dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:7px;vertical-align:1px}
  .hint{position:fixed;right:22px;bottom:138px;max-width:300px;background:#1c2330;color:#fff;padding:10px 12px;border-radius:10px;font-size:13px;line-height:1.4;box-shadow:0 4px 14px rgba(0,0,0,.2)}
  .overlay{position:fixed;inset:0;background:rgba(9,15,30,.45);display:flex;align-items:center;justify-content:center}
  .dialog{background:var(--panel);color:var(--ink);width:min(520px,94vw);max-height:92vh;overflow:auto;border-radius:12px;padding:22px;box-shadow:0 10px 40px rgba(0,0,0,.3);font-size:14px;line-height:1.45}
  .dialog h3{margin:0 0 4px;font-size:18px} .sub{color:var(--muted);margin-bottom:10px}
  label{display:block;font-size:12px;color:var(--muted);margin:12px 0 4px}
  input,textarea,select{width:100%;border:1px solid var(--line);border-radius:8px;padding:8px 10px;font-size:14px;color:var(--ink);background:#fff}
  textarea{min-height:84px;resize:vertical}
  .row{display:flex;gap:8px;align-items:center;flex-wrap:wrap} .end{justify-content:flex-end;margin-top:16px} .spacer{flex:1}
  .err{color:var(--bad);font-size:13px;margin-top:10px} .note{color:var(--muted);font-size:12px}
  .box{background:#fff4ed;border-radius:8px;padding:8px 10px;margin:8px 0}
  a{color:var(--accent)}
  /* кабинет */
  .cab{position:fixed;inset:0;background:var(--bg);color:var(--ink);display:flex;flex-direction:column;font-size:14px;line-height:1.45}
  .cab header{display:flex;align-items:center;gap:14px;padding:12px 20px;background:var(--panel);border-bottom:1px solid var(--line)}
  .cab header b{font-size:15px} .who{margin-left:auto;display:flex;gap:10px;align-items:center;color:var(--muted)}
  .badge{display:inline-block;font-size:12px;padding:2px 8px;border-radius:999px;background:var(--soft);white-space:nowrap}
  nav{display:flex;gap:2px;padding:0 20px;background:var(--panel);border-bottom:1px solid var(--line);overflow-x:auto}
  nav button{border:0;border-bottom:2px solid transparent;background:none;padding:11px 12px;color:var(--muted);cursor:pointer;font-size:14px;white-space:nowrap}
  nav button.on{color:var(--ink);border-bottom-color:var(--accent);font-weight:600}
  main{flex:1;overflow:auto;padding:20px} .wrap{max-width:1200px;margin:0 auto}
  h2{font-size:16px;margin:22px 0 10px} h2:first-child{margin-top:0}
  .tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px}
  .tile{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:14px 16px}
  .tile .v{font-size:26px;font-weight:700} .tile .k{color:var(--muted);font-size:12px}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:10px;overflow:auto}
  table{width:100%;border-collapse:collapse} th,td{text-align:left;padding:9px 12px;border-bottom:1px solid var(--line);vertical-align:top}
  th{font-size:12px;color:var(--muted);font-weight:600;background:var(--soft)} tr:last-child td{border-bottom:0}
  tr.click{cursor:pointer} tr.click:hover td{background:var(--soft)} .num{text-align:right;font-variant-numeric:tabular-nums}
  .empty{padding:22px;color:var(--muted);text-align:center}
  .s-new{color:var(--accent)} .s-in_progress{color:var(--warn)} .s-returned{color:var(--bad)} .s-approved{color:var(--ok)} .s-cancelled{color:var(--muted)}
  .bars{display:flex;align-items:flex-end;gap:6px;height:130px;padding:14px 14px 0}
  .bar{flex:1;display:flex;flex-direction:column;align-items:center;gap:4px;height:100%;justify-content:flex-end}
  .stack{width:100%;max-width:26px;display:flex;flex-direction:column-reverse;border-radius:4px 4px 0 0;overflow:hidden}
  .s1{background:var(--ok)} .s2{background:var(--bad)} .bar span{font-size:10px;color:var(--muted)}
  .legend{display:flex;gap:14px;padding:8px 14px 12px;font-size:12px;color:var(--muted)} .legend i{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:5px}
  .feed{list-style:none;margin:0;padding:0} .feed li{padding:9px 14px;border-bottom:1px solid var(--line)}
  .drawer{position:fixed;top:0;right:0;bottom:0;width:min(540px,100%);background:var(--panel);overflow:auto;padding:22px;border-left:1px solid var(--line);box-shadow:-8px 0 30px rgba(0,0,0,.12)}
  .kv{display:grid;grid-template-columns:130px 1fr;gap:6px 12px;margin:14px 0} .kv div:nth-child(odd){color:var(--muted)}
  .tl{list-style:none;padding:0;margin:0;border-left:2px solid var(--line)} .tl li{padding:0 0 12px 14px;position:relative}
  .tl li:before{content:"";position:absolute;left:-6px;top:5px;width:10px;height:10px;border-radius:50%;background:var(--accent)}
  .panel{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:16px 18px;max-width:640px;margin-bottom:16px}
  .toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);background:#1c2330;color:#fff;padding:10px 16px;border-radius:8px}
  .status-line{font-size:12px;color:var(--muted)}
  </style><div id="layer"></div>`;
  const layer = root.getElementById('layer');
  // Клик по подписи поля ставит курсор в поле под ней. Помогает, если из-за масштаба страницы
  // (часто при удалённом рабочем столе) клик попадает на несколько пикселей выше поля.
  root.addEventListener('click', (e) => {
    const t = e.composedPath()[0];
    const l = t && t.closest && t.closest('label');
    if (!l || l.control) return;
    let n = l.nextElementSibling;
    if (n && !/^(INPUT|TEXTAREA|SELECT)$/.test(n.tagName)) n = n.querySelector && n.querySelector('input,textarea,select');
    if (n && /^(INPUT|TEXTAREA|SELECT)$/.test(n.tagName) && n.type !== 'checkbox') { n.focus(); if (n.tagName === 'SELECT' && n.showPicker) try { n.showPicker(); } catch (x) { /* */ } }
  });
  // Горячие клавиши Confluence слушают нажатия на всей странице. Из-за изолированного блока (Shadow DOM)
  // Confluence не видит, что человек печатает в нашем поле, и перехватывает буквы и цифры.
  // Не выпускаем события клавиатуры из интерфейса плагина наружу.
  for (const type of ['keydown', 'keypress', 'keyup', 'input']) {
    root.addEventListener(type, (e) => {
      const t = e.composedPath()[0];
      if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable)) e.stopPropagation();
    });
  }
  const $ = (sel, el = root) => el.querySelector(sel);
  const $$ = (sel, el = root) => [...el.querySelectorAll(sel)];
  function toast(t) { const d = document.createElement('div'); d.className = 'toast'; d.textContent = t; root.appendChild(d); setTimeout(() => d.remove(), 3000); }
  function el(html) { const d = document.createElement('div'); d.innerHTML = html; return d.firstElementChild; }
  function modal(html) {
    const ov = el(`<div class="overlay"><div class="dialog" role="dialog" aria-modal="true">${html}</div></div>`);
    ov.addEventListener('mousedown', (e) => { if (e.target === ov) ov.remove(); });
    root.appendChild(ov);
    return ov;
  }

  // ── Привязка аккаунта Reddy (по коду от бота) ──
  function linkDialog(onDone) {
    const ov = modal(`<h3>Привязка Reddy</h3><div class="sub">Чтобы бот присылал вам уведомления и знал, кто отправил статью.</div>
      <div id="s1"><label>Ваш номер в Reddy (из профиля)</label><input id="key" placeholder="например, 73210000111" inputmode="numeric">
      <div class="note" style="margin-top:6px">Бот пришлёт вам в Reddy код подтверждения.</div></div>
      <div id="s2" hidden><label>Код из сообщения бота</label><input id="code" inputmode="numeric" maxlength="6"></div>
      <div class="err" id="err"></div><div class="row end"><button class="btn" id="x">Отмена</button><button class="btn primary" id="go">Получить код</button></div>`);
    let found = null, code = null;
    $('#x', ov).onclick = () => ov.remove();
    $('#key', ov).focus();
    $('#go', ov).onclick = async () => {
      const b = $('#go', ov); $('#err', ov).textContent = ''; b.disabled = true;
      try {
        if (!found) {
          const key = $('#key', ov).value.trim().replace(/^\+/, '');
          if (!/^\d{6,15}$/.test(key)) throw new Error('Введите номер цифрами');
          found = await findUser(key);
          if (!found) throw new Error('Пользователь с таким номером не найден в Reddy');
          found.key = key;
          code = String(Math.floor(100000 + Math.random() * 900000));
          const sent = await dm(found.id, `🔐 Код для привязки плагина Confluence Координатор: ${code}\nЕсли вы его не запрашивали — просто проигнорируйте.`);
          if (sent === null) throw new Error('Не удалось отправить код. Возможно, вы заблокировали бота — напишите ему /start');
          $('#s1', ov).hidden = true; $('#s2', ov).hidden = false; $('#code', ov).focus(); b.textContent = 'Подтвердить';
        } else {
          if ($('#code', ov).value.trim() !== code) throw new Error('Неверный код');
          const cfg = await getConfig();
          const u = await saveUser(found.id, (u) => {
            u.name = u.name || found.name; u.key = found.key; u.confluence = (confluenceMe && (confluenceMe.username || confluenceMe.accountId)) || u.confluence;
            if ((cfg.admins || []).includes(found.id)) u.role = 'admin';
            return u;
          });
          GM_setValue('me', { id: u.id, name: u.name, key: u.key });
          ov.remove(); toast(`Готово, ${u.name || 'аккаунт'} привязан`);
          if (onDone) onDone(u);
        }
      } catch (e) { $('#err', ov).textContent = e.message; }
      b.disabled = false;
    };
    $$('input', ov).forEach((i) => (i.onkeydown = (e) => { if (e.key === 'Enter') $('#go', ov).click(); }));
  }

  // ── Мастер первой настройки (для администратора) ──
  function setupDialog() {
    const ov = modal(`<h3>Первая настройка</h3><div class="sub">Это делается один раз. Вы станете администратором.</div>
      <label>Токен бота (от BotMother)</label><input id="tok" type="password" autocomplete="off">
      <label>ID чата проверяющих (команда /show_address в чате)</label><input id="chat" inputmode="numeric">
      <label>Ваш номер в Reddy (из профиля)</label><input id="key" inputmode="numeric">
      <div class="note" style="margin-top:10px">Токен сохранится в скрытых свойствах этой служебной страницы. Ограничьте доступ к странице командой, которая пользуется плагином.</div>
      <div class="err" id="err"></div><div class="row end"><button class="btn" id="x">Отмена</button><button class="btn primary" id="go">Проверить и сохранить</button></div>`);
    $('#x', ov).onclick = () => ov.remove();
    $('#go', ov).onclick = async () => {
      const b = $('#go', ov); b.disabled = true; $('#err', ov).textContent = '';
      try {
        const token = $('#tok', ov).value.trim(), chat = $('#chat', ov).value.trim(), key = $('#key', ov).value.trim().replace(/^\+/, '');
        if (!token || !/^-?\d+$/.test(chat) || !/^\d{6,15}$/.test(key)) throw new Error('Заполните все поля (ID чата и номер — цифрами)');
        const who = await reddy('whoami', null, null, token).catch(() => { throw new Error('Reddy не принял токен — проверьте его'); });
        const admin = await findUser(key, token);
        if (!admin) throw new Error('Пользователь с таким номером не найден в Reddy');
        const chatErr = await checkGroupChat(chat, token);
        if (chatErr) throw new Error(chatErr);
        const probe = await gm('POST', `${REDDY_API}/v2/send`, { token, body: { chat: Number(chat), msg: `👋 [b]Confluence Координатор[/b] подключён. Сюда будут приходить заявки на проверку правок.` } });
        if (probe.status !== 200) throw new Error(probe.status === 404 ? 'Бот не состоит в этом чате или ID чата неверный' : `Не удалось написать в чат (Reddy ${probe.status})`);
        await putProp(K.config, { token, chatId: chat, admins: [admin.id], slaTake: 30, slaCheck: 240, digestHour: 18, bot: who.username || '', created: now() }, (await getProp(K.config) || {}).version ?? null);
        cfgCache = null;
        await saveUser(admin.id, (u) => { u.name = u.name || admin.name; u.key = key; u.role = 'admin'; u.confluence = confluenceMe && (confluenceMe.username || confluenceMe.accountId); return u; });
        GM_setValue('me', { id: admin.id, name: admin.name, key });
        ov.remove(); toast('Настроено! В чат отправлено приветствие.');
        openCabinet();
      } catch (e) { $('#err', ov).textContent = e.message; }
      b.disabled = false;
    };
  }

  // ── Кнопка на странице статьи ──
  const STATE_TEXT = { new: (r) => `Ждёт проверяющего (#${r.id})`, in_progress: (r) => `На проверке у ${r.reviewerName} (#${r.id})`, returned: (r) => `Вопрос от ${r.reviewerName} (#${r.id})`, approved: (r) => `Проверено ✓ (#${r.id})`, cancelled: (r) => `Заявка #${r.id} отменена` };
  const DOT = { new: '#2f5bea', in_progress: '#b7791f', returned: '#c2410c', approved: '#1f8a4c', cancelled: '#6b7480' };
  function pageId() { const m = location.pathname.match(/\/pages\/(\d+)/) || location.search.match(/[?&]pageId=(\d+)/); return (m && m[1]) || meta('ajs-page-id') || null; }
  const sameUser = (a, b) => !!a && !!b && ((a.accountId && a.accountId === b.accountId) || (a.userKey && a.userKey === b.userKey) || (a.username && a.username === b.username));
  let pageCtx = null;

  async function renderFab() {
    $$('.fab,.hint', layer).forEach((x) => x.remove());
    const id = pageId();
    const fab = el('<div class="fab"></div>');
    layer.appendChild(fab);
    const cabBtn = el('<button class="btn pill state" title="Кабинет Confluence Координатора">🗂</button>');
    cabBtn.onclick = openCabinet;
    if (updateAvailable()) {
      const ub = el(`<button class="btn pill primary" title="Откроется окно Tampermonkey — нажмите «Обновить»">⬆️ Обновить плагин до ${esc(upd.latest)}</button>`);
      ub.onclick = () => window.open(upd.url, '_blank');
      fab.appendChild(ub);
    }
    if (id === String(STORAGE_PAGE_ID)) { // служебная страница с инструкцией — заметная кнопка кабинета и чек-лист подключения
      const big = el('<button class="btn pill primary">🗂 Открыть кабинет</button>'); big.onclick = () => openCabinet(); fab.appendChild(big);
      renderChecklist();
      return;
    }
    if (!id) { fab.appendChild(cabBtn); return; }

    let page;
    try { page = await cf('GET', `/rest/api/content/${id}?expand=version,space`); } catch (e) { fab.appendChild(cabBtn); return; }
    if (pageId() !== id) return;
    let req = null;
    try { const m = await getMeta(); req = m.byPage[id] ? await getReq(m.byPage[id]) : null; if (!req) { const all = await listProps('crb-r-'); req = all.filter((r) => r.pageId === id).sort((a, b) => b.id - a.id)[0] || null; } } catch (e) { /* хранилище недоступно */ }
    const byMe = sameUser(page.version && page.version.by, confluenceMe);
    const open = req && OPEN.includes(req.status);
    pageCtx = { page, req, byMe };
    const showState = req && req.status !== 'cancelled' && (req.version >= page.version.number || (open && !byMe));
    if (!byMe && !open) { fab.appendChild(cabBtn); return; }

    const b = el('<button class="btn pill"></button>');
    if (showState) {
      b.className = 'btn pill state';
      b.innerHTML = `<span class="dot" style="background:${DOT[req.status]}"></span>${esc(STATE_TEXT[req.status](req))}`;
      b.onclick = () => (req.status === 'returned' && byMe ? submitDialog() : openCabinet(req.id));
    } else {
      b.className = 'btn pill primary';
      b.textContent = open ? '🔄 Отправить доработку на проверку' : '📝 На проверку правок';
      b.onclick = submitDialog;
      const hk = `crb-hint-${id}-${page.version.number}`;
      if (byMe && !open && !sessionStorage.getItem(hk)) {
        sessionStorage.setItem(hk, '1');
        const h = el('<div class="hint">Вы изменили статью. Отправьте правки на проверку — заявка уйдёт в чат проверяющих в Reddy.</div>');
        layer.appendChild(h); setTimeout(() => h.remove(), 9000);
      }
    }
    fab.append(cabBtn, b);
  }

  /** Чек-лист «Подключение» вверху служебной страницы: что уже готово и чего не хватает */
  let checklistHost = null, canEditStorage = null;
  async function renderChecklist() {
    if (!checklistHost) {
      checklistHost = document.createElement('div');
      checklistHost.id = 'crb-setup';
      const anchor = document.getElementById('main-content') || document.querySelector('.wiki-content') || document.body.firstElementChild;
      anchor.parentNode.insertBefore(checklistHost, anchor);
      checklistHost.attachShadow({ mode: 'open' });
      try {
        const ops = (await cf('GET', `/rest/api/content/${encodeURIComponent(STORAGE_PAGE_ID)}?expand=operations`)).operations;
        canEditStorage = Array.isArray(ops) ? ops.some((o) => o.operation === 'update') : null;
      } catch (e) { canEditStorage = null; }
      setInterval(() => { if (checklistHost.isConnected && !document.hidden) drawChecklist(); }, 5000);
    }
    drawChecklist();
  }
  function drawChecklist() {
    const me = meLocal();
    const bot = { duty: ['ok', 'Бот на связи — эта вкладка обрабатывает кнопки'], standby: ['ok', 'Бот на связи'], idle: ['wait', 'Подключение к боту…'], error: ['bad', 'Бот: ' + dutyState.lastError], outdated: ['bad', dutyState.lastError] }[dutyState.role] || ['wait', 'Подключение к боту…'];
    const rows = [
      ['ok', `Плагин установлен (версия ${VERSION})`, updateAvailable() ? `<button data-x="upd">⬆️ Обновить до ${esc(upd.latest)}</button>` : ''],
      me ? ['ok', `Reddy привязан: ${esc(me.name || me.key || me.id)}`, ''] : ['bad', 'Reddy не привязан — бот не сможет присылать вам уведомления', '<button data-x="link" class="p">Привязать Reddy</button>'],
      canEditStorage === false ? ['bad', 'Нет права редактировать эту страницу — без него заявки не сохранятся. Попросите администратора.', ''] : canEditStorage ? ['ok', 'Право редактировать эту страницу есть', ''] : null,
      [bot[0], esc(bot[1]), ''],
    ].filter(Boolean);
    const ready = rows.every((r) => r[0] === 'ok');
    checklistHost.shadowRoot.innerHTML = `<style>
      .box{font:14px/1.45 -apple-system,"Segoe UI",Roboto,Arial,sans-serif;border:1px solid ${ready ? '#abf5d1' : '#ffe2bd'};background:${ready ? '#f3fcf7' : '#fffaf2'};border-radius:8px;padding:12px 16px;margin:0 0 18px;color:#172b4d}
      .h{font-weight:700;margin-bottom:6px} .r{display:flex;gap:10px;align-items:center;padding:3px 0} .i{width:20px;text-align:center}
      button{font:inherit;font-size:13px;padding:4px 10px;border-radius:6px;border:1px solid #c1c7d0;background:#fff;cursor:pointer} button.p{background:#0052cc;border-color:#0052cc;color:#fff}
    </style><div class="box"><div class="h">${ready ? '✅ Вы подключены к Confluence Координатору' : '🔧 Подключение к Confluence Координатору'}</div>
      ${rows.map(([st, text, act]) => `<div class="r"><span class="i">${{ ok: '✅', bad: '⚠️', wait: '⏳' }[st]}</span><span>${text}</span>${act}</div>`).join('')}</div>`;
    const sr = checklistHost.shadowRoot;
    const lk = sr.querySelector('[data-x="link"]'); if (lk) lk.onclick = () => linkDialog(() => { drawChecklist(); openCabinet(); });
    const ub = sr.querySelector('[data-x="upd"]'); if (ub) ub.onclick = () => window.open(upd.url, '_blank');
  }

  async function findBaseVersion(id, cur) {
    let base = cur - 1;
    for (let v = cur - 1, i = 0; v >= 1 && i < 15; v--, i++) {
      try { const old = await cf('GET', `/rest/api/content/${id}?status=historical&version=${v}&expand=version`); if (sameUser(old.version && old.version.by, confluenceMe)) base = v - 1; else break; } catch (e) { break; }
    }
    return Math.max(base, 1);
  }

  async function submitDialog() {
    if (!(await getConfig())) return toast('Плагин ещё не настроен — откройте кабинет 🗂');
    if (!meLocal()) return linkDialog(() => submitDialog());
    const { page, req } = pageCtx;
    const cur = page.version.number;
    const resub = req && OPEN.includes(req.status);
    const base = resub && req.baseVersion ? req.baseVersion : await findBaseVersion(page.id, cur);
    const opts = [];
    for (let v = cur - 1; v >= Math.max(1, cur - 30); v--) opts.push(`<option value="${v}" ${v === base ? 'selected' : ''}>версия ${v}${v === base ? ' (до ваших правок)' : ''}</option>`);
    const ov = modal(`<h3>${resub ? 'Отправить доработку' : 'Проверка правок'}</h3><div class="sub">«${esc(page.title)}» · версия ${cur}</div>
      ${req && req.status === 'returned' ? `<div class="box"><b>Вопрос проверяющего:</b> ${esc(req.question)}</div>` : ''}
      <label>Что изменилось и на что обратить внимание</label><textarea id="c" placeholder="Например: обновил сроки в разделе 3"></textarea>
      <label>Подтверждение согласования — ссылка на сообщение в Reddy <span style="color:var(--bad)">*</span></label>
      <input id="appr" placeholder="https://lnk.reddy.team/message/…" value="${esc((req && OPEN.includes(req.status) && req.approvalUrl) || '')}" autocomplete="off">
      <div class="note" style="margin-top:4px">В Reddy: правой кнопкой по сообщению с согласованием → «Копировать ссылку».</div>
      ${cur > 1 && !resub ? `<label>Сравнивать с</label><select id="base">${opts.join('')}</select>` : ''}
      <label style="display:flex;gap:8px;align-items:center;color:var(--ink);font-size:14px"><input type="checkbox" id="u" style="width:auto"> Срочно</label>
      <div class="err" id="err"></div><div class="row end"><button class="btn" id="x">Отмена</button><button class="btn primary" id="go">Отправить</button></div>`);
    $('#x', ov).onclick = () => ov.remove();
    $('#c', ov).focus();
    $('#appr', ov).oninput = () => { $('#err', ov).textContent = ''; };
    $('#go', ov).onclick = async () => {
      const b = $('#go', ov); $('#err', ov).textContent = '';
      const approvalUrl = $('#appr', ov).value.trim();
      if (!approvalUrl) { $('#err', ov).textContent = 'Прикрепите ссылку на сообщение в Reddy с согласованием правок'; $('#appr', ov).focus(); return; }
      if (!APPROVAL_RE.test(approvalUrl)) { $('#err', ov).textContent = 'Это не похоже на ссылку на сообщение Reddy. Нужна ссылка вида https://lnk.reddy.team/message/…'; $('#appr', ov).focus(); return; }
      b.disabled = true; b.textContent = 'Отправляю…';
      try {
        const me = await touchUser(meLocal().id, meLocal().name, meLocal().key);
        const res = await submit(me, {
          pageId: String(page.id), title: page.title, space: page.space && page.space.key,
          url: page._links && page._links.webui ? CB + page._links.webui : location.href.split('#')[0],
          version: cur, prevVersion: $('#base', ov) ? +$('#base', ov).value : base,
          authorLogin: confluenceMe && (confluenceMe.username || confluenceMe.accountId), authorName: confluenceMe && confluenceMe.displayName,
          comment: $('#c', ov).value.trim(), urgent: $('#u', ov).checked, approvalUrl,
        });
        if (!res.ok) throw new Error(res.error);
        const msg = { created: `Заявка #${res.req.id} отправлена в чат проверяющих`, updated: `Заявка #${res.req.id} обновлена до версии ${cur}`, resubmitted: `Доработка по заявке #${res.req.id} отправлена ${res.req.reviewerName || ''}` }[res.action];
        const warn = res.cardFailed ? `<p class="err">Заявка сохранена, но карточку не удалось отправить в чат — она появится там при следующем действии. Проверьте, что бот состоит в чате.</p>` : '';
        $('.dialog', ov).innerHTML = `<h3>Готово ✓</h3><p>${esc(msg)}</p>${warn}<p class="sub">Статус придёт вам в Reddy. Все заявки — в кабинете 🗂.</p><div class="row end"><button class="btn primary" id="ok">Ок</button></div>`;
        $('#ok', ov).onclick = () => { ov.remove(); renderFab(); };
      } catch (e) { b.disabled = false; b.textContent = 'Отправить'; $('#err', ov).textContent = e.message; }
    };
  }

  // ── Кабинет ──
  const RU_STATUS = { new: 'Ожидает', in_progress: 'На проверке', returned: 'Вопрос автору', approved: 'Подтверждено', cancelled: 'Отменено' };
  const ROLE = { employee: 'Сотрудник', senior: 'Старший', admin: 'Администратор' };
  const ACTION = { created: 'создал(а) заявку', updated: 'обновил(а) версию', resubmitted: 'отправил(а) доработку', taken: 'взял(а) в работу', released: 'вернул(а) в очередь', returned: 'вернул(а) с вопросом', answered: 'ответил(а) на вопрос', approved: 'подтвердил(а)', cancelled: 'отменил(а)', sla_take: 'никто не взял вовремя', sla_check: 'проверка затянулась' };
  const badge = (s) => `<span class="badge s-${s}">${RU_STATUS[s]}</span>`;
  const sdur = (ms) => dur(ms).replace('меньше минуты', '<1 мин');
  let CAB = null;

  async function openCabinet(openId) {
    let cfg;
    try { cfg = await getConfig(true); } catch (e) {
      return toast(e.status === 404 ? 'Служебная страница не найдена — проверьте ID в настройках плагина' : e.status === 403 ? 'Нет доступа к служебной странице' : 'Не удалось открыть хранилище: ' + e.message);
    }
    if (!cfg) return setupDialog();
    if (!meLocal()) return linkDialog(() => openCabinet(openId));
    $$('.cab', layer).forEach((x) => x.remove());
    CAB = { tab: sessionStorage.getItem('crb-tab') || 'mine', openId: null };
    const c = el(`<div class="cab"><header><b>🗂 Confluence Координатор</b><span class="status-line" id="duty"></span><div class="who"><span id="nm"></span><span class="badge" id="rl"></span><button class="btn ghost" id="cls">✕ Закрыть</button></div></header><nav id="nav"></nav><main><div class="wrap" id="view"><div class="empty">Загрузка…</div></div></main></div>`);
    layer.appendChild(c);
    $('#cls', c).onclick = () => { c.remove(); CAB = null; };
    await cabRender();
    if (openId) openReq(openId);
  }

  async function cabLoad() {
    const [reqs, users, me, meta, logp] = await Promise.all([listProps('crb-r-'), listProps('crb-u-'), meUser(), getMeta(), getProp('crb-log')]);
    return { reqs: reqs.sort((a, b) => b.id - a.id), users, me: me || { ...meLocal(), role: 'employee' }, meta, botLog: (logp && logp.value.items) || [] };
  }

  async function cabRender() {
    if (!CAB) return;
    const c = $('.cab', layer);
    const data = await cabLoad();
    CAB.data = data;
    const me = data.me, senior = isSenior(me);
    $('#nm', c).textContent = me.name || me.id; $('#rl', c).textContent = ROLE[me.role] || '';
    $('#duty', c).textContent = { duty: '● эта вкладка сейчас дежурит — обрабатывает кнопки бота', standby: '● бот на связи (дежурит другая вкладка)', error: '⚠ ' + dutyState.lastError, outdated: '⚠ ' + dutyState.lastError, idle: '● подключение к Reddy…' }[dutyState.role];
    const queueN = data.reqs.filter((r) => r.status === 'new').length;
    const tabs = [['mine', 'Мои заявки']];
    if (senior) tabs.push(['queue', `Очередь${queueN ? ` (${queueN})` : ''}`], ['team', 'Команда'], ['all', 'Все заявки'], ['users', 'Сотрудники']);
    tabs.push(['settings', 'Настройки']);
    if (!tabs.some(([k]) => k === CAB.tab)) CAB.tab = 'mine';
    $('#nav', c).innerHTML = tabs.map(([k, t]) => `<button data-t="${k}" class="${k === CAB.tab ? 'on' : ''}">${t}</button>`).join('');
    $$('#nav button', c).forEach((b) => (b.onclick = () => { CAB.tab = b.dataset.t; sessionStorage.setItem('crb-tab', CAB.tab); CAB.openId = null; $$('.drawer', layer).forEach((x) => x.remove()); cabRender(); }));
    const v = $('#view', c);
    try { VIEWS[CAB.tab](v, data); } catch (e) { v.innerHTML = `<div class="empty">${esc(e.message)}</div>`; }
    $$('[data-open]', v).forEach((x) => (x.onclick = (e) => { e.preventDefault(); openReq(+x.dataset.open); }));
    $$('[data-take]', v).forEach((x) => (x.onclick = (e) => { e.stopPropagation(); act(+x.dataset.take, 'take'); }));
  }

  function reqTable(rows, { reviewer = true, author = true, empty = 'Пусто' } = {}) {
    if (!rows.length) return `<div class="card"><div class="empty">${empty}</div></div>`;
    return `<div class="card"><table><thead><tr><th>#</th><th>Статья</th><th>Статус</th>${author ? '<th>Автор</th>' : ''}${reviewer ? '<th>Проверяющий</th>' : ''}<th>Создана</th><th class="num">Ожидание / проверка</th></tr></thead><tbody>
    ${rows.map((r) => `<tr class="click" data-open="${r.id}"><td>${r.id}</td><td>${r.urgent ? '🔥 ' : ''}${esc(r.title || 'Без названия')}<div class="note">${esc(r.space || '')} · версия ${esc(r.version ?? '?')}${r.round > 1 ? ' · круг ' + r.round : ''}</div></td>
      <td>${badge(r.status)}</td>${author ? `<td>${esc(r.authorName || '—')}</td>` : ''}${reviewer ? `<td>${esc(r.reviewerName || '—')}</td>` : ''}<td class="note">${fmtT(r.created)}</td>
      <td class="num note">${r.status === 'new' ? 'ждёт ' + sdur(now() - r.created) : r.taken ? sdur((r.closed || now()) - r.taken) : '—'}</td></tr>`).join('')}</tbody></table></div>`;
  }

  const VIEWS = {
    mine(v, { reqs, me }) {
      if (!isSenior(me)) return VIEWS.mineEmployee(v, { reqs, me });
      const rev = reqs.filter((r) => r.reviewerId === me.id && OPEN.includes(r.status));
      const auth = reqs.filter((r) => r.authorId === me.id);
      const since = now() - 30 * 864e5;
      const st = teamStats(reqs, [], since).find((x) => x.id === me.id) || {};
      v.innerHTML = `<div class="tiles">
        <div class="tile"><div class="v">${rev.filter((r) => r.status === 'in_progress').length}</div><div class="k">у меня на проверке</div></div>
        <div class="tile"><div class="v">${st.approved || 0}</div><div class="k">подтверждено за 30 дней</div></div>
        <div class="tile"><div class="v">${st.returned || 0}</div><div class="k">возвращено с вопросом за 30 дней</div></div>
        <div class="tile"><div class="v">${sdur(st.avgCheck)}</div><div class="k">среднее время проверки</div></div>
        <div class="tile"><div class="v">${reqs.filter((r) => r.status === 'new').length}</div><div class="k">в общей очереди</div></div></div>
        <h2>Я проверяю</h2>${reqTable(rev, { reviewer: false, empty: 'Нет заявок. Возьмите из очереди.' })}
        <h2>Мои заявки как автора</h2>${reqTable(auth.slice(0, 50), { author: false, empty: 'Вы ещё не отправляли статьи на проверку.' })}`;
    },
    /** Сотрудник видит только свои заявки: что отправил, что ждёт его ответа, что подтверждено */
    mineEmployee(v, { reqs, me }) {
      const auth = reqs.filter((r) => r.authorId === me.id);
      const since = now() - 30 * 864e5;
      const done = auth.filter((r) => r.status === 'approved' && r.closed > since);
      const avg = done.length ? done.reduce((a, r) => a + (r.closed - r.created), 0) / done.length : null;
      const ask = auth.filter((r) => r.status === 'returned');
      v.innerHTML = `<div class="tiles">
        <div class="tile"><div class="v">${auth.filter((r) => r.status === 'new' || r.status === 'in_progress').length}</div><div class="k">на проверке и в очереди</div></div>
        <div class="tile"><div class="v" style="${ask.length ? 'color:var(--bad)' : ''}">${ask.length}</div><div class="k">ждут моего ответа</div></div>
        <div class="tile"><div class="v">${done.length}</div><div class="k">подтверждено за 30 дней</div></div>
        <div class="tile"><div class="v">${sdur(avg)}</div><div class="k">в среднем от отправки до подтверждения</div></div></div>
        ${ask.length ? `<h2>❓ Ждут моего ответа</h2>${reqTable(ask, { author: false })}` : ''}
        <h2>Мои заявки</h2>${reqTable(auth.slice(0, 100), { author: false, empty: 'Вы ещё не отправляли статьи на проверку. Внесите правки в статью и нажмите «📝 На проверку правок».' })}`;
    },
    queue(v, { reqs }) {
      const q = reqs.filter((r) => r.status === 'new').sort((a, b) => b.urgent - a.urgent || a.id - b.id);
      v.innerHTML = q.length ? `<div class="card"><table><thead><tr><th>#</th><th>Статья</th><th>Автор</th><th>Комментарий</th><th class="num">Ждёт</th><th></th></tr></thead><tbody>
        ${q.map((r) => `<tr class="click" data-open="${r.id}"><td>${r.id}</td><td>${r.urgent ? '🔥 ' : ''}${esc(r.title)}<div class="note">${esc(r.space || '')}</div></td><td>${esc(r.authorName)}</td>
          <td class="note">${esc((r.comment || '').slice(0, 140))}</td><td class="num">${sdur(now() - r.created)}</td><td><button class="btn primary" data-take="${r.id}">Взять</button></td></tr>`).join('')}</tbody></table></div>`
        : '<div class="card"><div class="empty">Очередь пуста 🎉</div></div>';
    },
    team(v, { reqs, users }) {
      const period = sessionStorage.getItem('crb-period') || 'today';
      const from = { today: startOfDay(), week: now() - 7 * 864e5, month: now() - 30 * 864e5 }[period];
      const rows = teamStats(reqs, users, from).filter((x) => x.taken || x.approved || x.returned || x.inWork || x.authored);
      const open = reqs.filter((r) => r.status === 'new');
      const series = [];
      for (let i = 13; i >= 0; i--) {
        const f = startOfDay() - i * 864e5, t = f + 864e5;
        const evs = reqs.flatMap((r) => r.events || []).filter((e) => e.t >= f && e.t < t);
        series.push({ d: f, a: evs.filter((e) => e.a === 'approved').length, r: evs.filter((e) => e.a === 'returned').length });
      }
      const max = Math.max(1, ...series.map((x) => x.a + x.r));
      const feed = reqs.flatMap((r) => (r.events || []).map((e) => ({ ...e, id: r.id, title: r.title }))).sort((a, b) => b.t - a.t).slice(0, 50);
      v.innerHTML = `<div class="tiles">
        <div class="tile"><div class="v">${open.length}</div><div class="k">в очереди${open.length ? ' · самая старая ' + sdur(now() - Math.min(...open.map((r) => r.created))) : ''}</div></div>
        <div class="tile"><div class="v">${reqs.filter((r) => r.status === 'in_progress').length}</div><div class="k">на проверке сейчас</div></div>
        <div class="tile"><div class="v">${reqs.filter((r) => r.status === 'returned').length}</div><div class="k">ждут ответа автора</div></div></div>
        <div class="row" style="margin-top:22px"><h2 style="margin:0">Сотрудники</h2><span class="spacer"></span><select id="per" style="width:auto"><option value="today">Сегодня</option><option value="week">7 дней</option><option value="month">30 дней</option></select></div><div style="height:10px"></div>
        ${rows.length ? `<div class="card"><table><thead><tr><th>Сотрудник</th><th class="num">Взял</th><th class="num">Подтвердил</th><th class="num">Вернул</th><th class="num">Отказался</th><th class="num">В работе</th><th class="num">Ср. проверка</th><th class="num">Ср. ожидание</th><th class="num">Отправил сам</th></tr></thead><tbody>
          ${rows.map((x) => `<tr class="click" data-rev="${esc(x.id)}"><td>${esc(x.name)}</td><td class="num">${x.taken}</td><td class="num">${x.approved}</td><td class="num">${x.returned}</td><td class="num">${x.released || ''}</td><td class="num">${x.inWork || ''}</td><td class="num">${sdur(x.avgCheck)}</td><td class="num">${sdur(x.avgWait)}</td><td class="num">${x.authored || ''}</td></tr>`).join('')}</tbody></table></div>`
          : '<div class="card"><div class="empty">За период активности не было</div></div>'}
        <h2>Последние 14 дней</h2><div class="card"><div class="bars">${series.map((x) => `<div class="bar" title="подтверждено ${x.a}, возвращено ${x.r}"><div class="stack" style="height:${((x.a + x.r) / max) * 100}%">${x.a ? `<div class="s1" style="flex:${x.a}"></div>` : ''}${x.r ? `<div class="s2" style="flex:${x.r}"></div>` : ''}</div><span>${tzParts(x.d + 3600000).day}</span></div>`).join('')}</div>
        <div class="legend"><span><i style="background:var(--ok)"></i>подтверждено</span><span><i style="background:var(--bad)"></i>возвращено с вопросом</span></div></div>
        <h2>Лента событий</h2><div class="card"><ul class="feed">${feed.map((e) => `<li><span class="note">${fmtT(e.t)}</span> · <b>${esc(e.n)}</b> ${ACTION[e.a] || e.a} <a href="#" data-open="${e.id}">#${e.id}</a> <span class="note">${esc(e.title || '')}</span>${e.d && ['returned', 'answered'].includes(e.a) ? `<div class="note">${esc(e.d)}</div>` : ''}</li>`).join('') || '<li class="note">Пока пусто</li>'}</ul></div>`;
      const sel = $('#per', v); sel.value = period; sel.onchange = () => { sessionStorage.setItem('crb-period', sel.value); cabRender(); };
      $$('[data-rev]', v).forEach((tr) => (tr.onclick = () => { sessionStorage.setItem('crb-frev', tr.dataset.rev); CAB.tab = 'all'; cabRender(); }));
    },
    all(v, { reqs, users }) {
      const f = { s: sessionStorage.getItem('crb-fs') || '', q: sessionStorage.getItem('crb-fq') || '', rv: sessionStorage.getItem('crb-frev') || '' };
      const rows = reqs.filter((r) => (!f.s || r.status === f.s) && (!f.rv || r.reviewerId === f.rv) && (!f.q || `${r.title} ${r.space} #${r.id}`.toLowerCase().includes(f.q.toLowerCase())));
      v.innerHTML = `<div class="row"><input id="fq" placeholder="Поиск по названию, пространству или #номеру" style="max-width:320px" value="${esc(f.q)}">
        <select id="fs" style="width:auto"><option value="">Все статусы</option>${Object.entries(RU_STATUS).map(([k, t]) => `<option value="${k}">${t}</option>`).join('')}</select>
        <select id="fr" style="width:auto"><option value="">Все проверяющие</option>${users.map((u) => `<option value="${esc(u.id)}">${esc(u.name || u.id)}</option>`).join('')}</select>
        <span class="spacer"></span><span class="note">${rows.length} шт.</span></div><div style="height:12px"></div>${reqTable(rows.slice(0, 300))}`;
      $('#fs', v).value = f.s; $('#fr', v).value = f.rv;
      const apply = () => { sessionStorage.setItem('crb-fs', $('#fs', v).value); sessionStorage.setItem('crb-fq', $('#fq', v).value.trim()); sessionStorage.setItem('crb-frev', $('#fr', v).value); cabRender(); };
      $('#fs', v).onchange = apply; $('#fr', v).onchange = apply; $('#fq', v).onkeydown = (e) => { if (e.key === 'Enter') apply(); };
    },
    users(v, { users, me }) {
      const admin = me.role === 'admin';
      v.innerHTML = `<div class="panel" style="max-width:none"><h2>Добавить сотрудника</h2>
        <div class="row" style="align-items:flex-end">
          <div style="flex:1;min-width:200px"><label>Номер из профиля Reddy</label><input id="addKey" inputmode="numeric" placeholder="например, 73210000111"></div>
          <div style="width:200px"><label>Роль</label><select id="addRole">${Object.entries(ROLE).filter(([k]) => admin || k === 'employee').map(([k, t]) => `<option value="${k}">${t}</option>`).join('')}</select></div>
          <button class="btn primary" id="addGo">Добавить</button>
        </div>
        <div class="err" id="addErr"></div>
        <p class="note" style="margin:8px 0 0">Бот найдёт человека в Reddy и пришлёт ему приветствие с меню.${admin ? '' : ' Назначать старших и администраторов может только администратор.'}</p></div>
        <p class="note">Сотрудник также появляется здесь сам — после первого сообщения боту, нажатия кнопки или привязки плагина. </p>
        <p class="note"><b>Сотрудник</b> — отправляет статьи на проверку и видит только свои заявки (так получают все, кто подключился сам). <b>Старший</b> — берёт и проверяет заявки, видит очередь и статистику. <b>Администратор</b> — то же плюс роли и настройки. Повысить роль может только администратор.</p>
        <div class="card"><table><thead><tr><th>Имя</th><th>Номер Reddy</th><th>Логин Confluence</th><th>Роль</th><th>Активен</th></tr></thead><tbody>
        ${users.sort((a, b) => (a.name || '').localeCompare(b.name || '')).map((u) => `<tr data-u="${esc(u.id)}"><td>${admin ? `<input data-f="name" value="${esc(u.name || '')}">` : esc(u.name || '—')}</td><td class="note">${esc(u.key || 'ID ' + u.id)}</td><td class="note">${esc(u.confluence || '—')}</td>
          <td>${admin ? `<select data-f="role">${Object.entries(ROLE).map(([k, t]) => `<option value="${k}" ${k === u.role ? 'selected' : ''}>${t}</option>`).join('')}</select>` : ROLE[u.role]}</td>
          <td>${admin ? `<input type="checkbox" data-f="active" ${u.active !== false ? 'checked' : ''} style="width:auto">` : u.active !== false ? 'да' : 'нет'}</td></tr>`).join('')}</tbody></table></div>`;
      $('#addKey', v).oninput = () => { $('#addErr', v).textContent = ''; };
      // Диагностика: если после клика поле не получило фокус — показываем, что перехватило клик или фокус
      const addKey = $('#addKey', v);
      const desc = (e) => (e ? `${e.tagName.toLowerCase()}${e.id ? '#' + e.id : ''}${typeof e.className === 'string' && e.className ? '.' + e.className.trim().split(/\s+/).slice(0, 3).join('.') : ''}` : '—');
      addKey.parentElement.addEventListener('pointerdown', (ev) => {
        const x = ev.clientX, y = ev.clientY;
        setTimeout(() => {
          if (root.activeElement === addKey || (root.activeElement && /^(INPUT|SELECT|TEXTAREA)$/.test(root.activeElement.tagName))) return; // всё в порядке
          const diag = `Диагностика: сверху «${desc(document.elementFromPoint(x, y))}», внутри «${desc(root.elementFromPoint(x, y))}», фокус у «${desc(document.activeElement)}» / «${desc(root.activeElement)}», disabled=${addKey.disabled}, readonly=${addKey.readOnly}. Пришлите этот текст разработчику. Пока можно оставить поле пустым и нажать «Добавить».`;
          $('#addErr', v).textContent = diag;
          console.warn('[Координатор]', diag);
        }, 300);
      }, true);
      $('#addKey', v).onkeydown = (e) => { if (e.key === 'Enter') $('#addGo', v).click(); };
      $('#addGo', v).onclick = async () => {
        const b = $('#addGo', v), err = $('#addErr', v);
        let key = $('#addKey', v).value.trim().replace(/^\+/, '');
        const role = $('#addRole', v).value;
        err.textContent = '';
        // Запасной путь: если в поле почему-то не получается ввести номер — спросим его отдельным окошком браузера
        if (!key) { const typed = prompt('Номер сотрудника из профиля Reddy'); if (typed === null) return; key = typed.trim().replace(/^\+/, ''); }
        if (!/^\d{6,15}$/.test(key)) { err.textContent = 'Введите номер цифрами — он есть в профиле Reddy'; return; }
        const res = await addUser(key, role, me);
        if (!res.ok) { err.textContent = res.error; return; }
        toast(res.message); cabRender();
      };
      if (admin) $$('[data-f]', v).forEach((x) => (x.onchange = async () => {
        const id = x.closest('tr').dataset.u, f = x.dataset.f, val = x.type === 'checkbox' ? x.checked : x.value;
        if (id === me.id && ((f === 'role' && val !== 'admin') || (f === 'active' && !val))) { toast('Нельзя снять права с себя'); return cabRender(); }
        await saveUser(id, (u) => { u[f] = val; return u; }); toast('Сохранено');
      }));
    },
    settings(v, { me, meta, botLog }) {
      const admin = me.role === 'admin', cfg = cfgCache || {};
      v.innerHTML = `<div class="panel"><h2>Мой профиль</h2>
        <div class="kv"><div>Имя</div><div>${esc(me.name || '—')}</div><div>Номер Reddy</div><div>${esc(me.key || '—')}</div><div>Роль</div><div>${ROLE[me.role] || '—'}</div></div>
        ${isSenior(me) ? `<label style="display:flex;gap:8px;align-items:center;color:var(--ink);font-size:14px"><input type="checkbox" id="ntf" style="width:auto" ${me.notify !== false ? 'checked' : ''}> Присылать мне в Reddy уведомления о действиях сотрудников</label>` : ''}
        <div class="row" style="margin-top:12px"><button class="btn" id="relink">Привязать другой аккаунт Reddy</button></div></div>
        ${admin ? `<div class="panel"><h2>Настройки бота (администратор)</h2>
          <div class="box" style="background:${meta.threadError || meta.threadSendError ? '#fff4ed' : '#eef8f1'}">${meta.threadError
            ? `🧵 Обсуждения под заявками пока не создаются: ${esc(meta.threadError)}.<br>Попросите поддержку Reddy включить боту возможности <b>создавать, читать и писать в треды</b>. Пока прав нет, вопросы и ответы идут через личку с ботом.`
            : meta.threadSendError ? `🧵 Обсуждения создаются, но бот не может в них написать: <b>${esc(meta.threadSendError)}</b>.<br>Скорее всего, у бота нет права <b>писать в треды</b> — уточните у поддержки Reddy.`
            : meta.threadCheckedAt ? '🧵 Обсуждения под заявками работают.' : '🧵 Обсуждения проверятся при следующей заявке.'}</div>
          <label style="display:flex;gap:8px;align-items:center;color:var(--ink);font-size:14px"><input type="checkbox" id="sdm" style="width:auto" ${cfg.seniorDm !== false ? 'checked' : ''}> Сообщать старшим в личку о заявках (по каждой заявке одно сообщение, которое обновляется; отдельно — только напоминания ❗)</label>
          <label>Ссылка на файл плагина на GitHub (пусто — репозиторий по умолчанию: kadirovm-eng/Problemati_Complaints)</label>
          <input id="inst" value="${esc(cfg.installUrl || '')}" placeholder="${esc(DEFAULT_INSTALL_URL)}">
          <label>ID общего чата, куда приходят заявки (бот должен быть участником; /show_address — в самом чате, не в личке с ботом)</label><input id="chat" value="${esc(cfg.chatId || '')}">
          <label>Напоминать старшим, если заявку не взяли за, минут</label><input id="st" type="number" value="${esc(cfg.slaTake ?? 30)}">
          <label>Напоминать старшим, если проверка идёт дольше, минут</label><input id="sc" type="number" value="${esc(cfg.slaCheck ?? 240)}">
          <label>Час ежедневной сводки старшим (по Москве, -1 — выключить)</label><input id="dh" type="number" value="${esc(cfg.digestHour ?? 18)}">
          <label>Новый токен бота (оставьте пустым, чтобы не менять)</label><input id="tok" type="password" autocomplete="off">
          <div class="row" style="margin-top:12px"><button class="btn primary" id="save">Сохранить</button></div></div>
          <div class="panel" style="max-width:none"><h2>Журнал бота</h2><p class="note">Последние 50 событий из Reddy и что бот с ними сделал. Версия плагина у вас: ${VERSION}${meta.version ? `, последняя установленная: ${esc(meta.version)}` : ''}.</p>
          ${botLog.length ? `<div class="card"><table><thead><tr><th>Время</th><th>Кто</th><th>Что пришло</th><th>Что сделал бот</th><th>Версия</th></tr></thead><tbody>${botLog.map((x) => `<tr><td class="note">${fmtT(x.t)}</td><td>${esc(x.who)}</td><td><span class="note">${esc(x.type)}</span><br>${esc(x.text)}</td><td>${esc(x.note)}</td><td class="note">${esc(x.v || '—')}</td></tr>`).join('')}</tbody></table></div>` : '<div class="empty">Пока пусто — события появятся после следующих нажатий и сообщений.</div>'}</div>` : ''}`;
      const n = $('#ntf', v); if (n) n.onchange = async () => { await saveUser(me.id, (u) => { u.notify = n.checked; return u; }); toast('Сохранено'); };
      $('#relink', v).onclick = () => linkDialog(() => cabRender());
      if (admin) $('#save', v).onclick = async () => {
        const newChat = $('#chat', v).value.trim(), newTok = $('#tok', v).value.trim();
        const inst = normalizeInstallUrl($('#inst', v).value);
        if (inst && inst !== cfg.installUrl) {
          if (!INSTALL_RE.test(inst)) return toast('Нужна ссылка «Raw» на файл .user.js на GitHub (raw.githubusercontent.com или gist.githubusercontent.com)');
          const ver = await fetchRemoteVersion(inst);
          if (!ver) return toast('По этой ссылке не удалось прочитать плагин — проверьте, что файл открывается без входа в GitHub');
          toast(`Файл плагина найден, версия ${ver}`);
        }
        if (newChat && (newChat !== String(cfg.chatId) || newTok)) {
          const err = await checkGroupChat(newChat, newTok || cfg.token);
          if (err) return toast(err);
          if (newChat !== String(cfg.chatId)) {
            const sent = await gm('POST', `${REDDY_API}/v2/send`, { token: newTok || cfg.token, body: { chat: Number(newChat), msg: '👋 [b]Confluence Координатор[/b] подключён к этому чату. Сюда будут приходить заявки на проверку правок.' } });
            if (sent.status !== 200) return toast(`Не удалось написать в этот чат (Reddy ${sent.status})`);
          }
        }
        await mutate(K.config, (c) => {
          c.installUrl = inst || null; c.seniorDm = $('#sdm', v).checked; c.chatId = newChat || c.chatId; c.slaTake = +$('#st', v).value; c.slaCheck = +$('#sc', v).value; c.digestHour = +$('#dh', v).value;
          if (newTok) c.token = newTok;
          return c;
        }, null);
        cfgCache = null; await getConfig(true); toast('Настройки сохранены');
        await checkUpdate(true); renderFab();
      };
    },
  };

  /** Добавить сотрудника по номеру Reddy. Старший может добавить только сотрудника, роли выше — только администратор. */
  async function addUser(key, role, me) {
    if (!isSenior(me)) return fail('Добавлять сотрудников могут старшие и администратор');
    if (role !== 'employee' && me.role !== 'admin') return fail('Назначать старших и администраторов может только администратор');
    const exist = (await listProps('crb-u-')).find((u) => u.key === key);
    if (exist) return fail(`${exist.name || 'Этот сотрудник'} уже в списке${exist.role !== role && me.role === 'admin' ? ' — роль можно поменять в таблице ниже' : ''}`);
    const found = await findUser(key);
    if (!found) return fail('Пользователь с таким номером не найден в Reddy (или заблокировал бота)');
    const before = await getUser(found.id);
    if (before) return fail(`${before.name || 'Этот сотрудник'} уже в списке`);
    const u = await saveUser(found.id, (x) => { x.name = x.name || found.name; x.key = key; x.role = role; x.addedBy = me.id; return x; });
    const v = await menuView('home', u);
    await dm(u.id, `👋 ${safe(uname(me))} добавил(а) вас в Confluence Координатор${role === 'employee' ? '' : ` с ролью «${ROLE[role]}»`}.\n\n${v.text}`, v.buttons);
    return { ok: true, message: `${u.name || key} добавлен(а)` };
  }

  async function openReq(id) {
    if (!CAB) return;
    CAB.openId = id;
    const r = await getReq(id);
    if (!r) return toast('Заявка не найдена');
    const me = CAB.data.me, senior = isSenior(me), isAuth = r.authorId === me.id;
    if (!senior && !isAuth) { CAB.openId = null; return toast('Сотрудникам доступны только их собственные заявки'); }
    let actions = '';
    if (r.status === 'new' && senior) actions += '<button class="btn primary" data-a="take">Взять в работу</button>';
    if (r.status === 'in_progress' && senior) actions += '<button class="btn primary" data-a="approve">Подтвердить корректность</button><button class="btn" data-a="return">Вернуть с вопросом</button><button class="btn" data-a="release">Вернуть в очередь</button>';
    if (r.status === 'returned' && isAuth) actions += '<button class="btn primary" data-a="answer">Ответить проверяющему</button>';
    if (OPEN.includes(r.status) && (isAuth || senior)) actions += '<button class="btn ghost" data-a="cancel">Отменить заявку</button>';
    $$('.drawer', layer).forEach((x) => x.remove());
    const d = el(`<aside class="drawer"><div class="row"><span class="note">Заявка #${r.id}</span><span class="spacer"></span><button class="btn ghost" id="dx">✕</button></div>
      <h3 style="margin:6px 0">${r.urgent ? '🔥 ' : ''}${esc(r.title)}</h3>${badge(r.status)}
      <div class="row" style="margin-top:12px">${r.url ? `<a href="${esc(r.url)}" target="_blank">Открыть статью ↗</a>` : ''}${r.diffUrl ? `<a href="${esc(r.diffUrl)}" target="_blank">Сравнить версии ↗</a>` : ''}${chatLink(r) ? `<a href="${esc(chatLink(r))}" target="_blank">Карточка в чате Reddy ↗</a>` : ''}${r.approvalUrl ? `<a href="${esc(r.approvalUrl)}" target="_blank">🤝 Согласование ↗</a>` : ''}</div>
      <div class="kv"><div>Пространство</div><div>${esc(r.space || '—')}</div><div>Версии</div><div>${esc(r.baseVersion ?? '?')} → ${esc(r.version ?? '?')}${r.round > 1 ? ' · круг ' + r.round : ''}</div>
      <div>Автор</div><div>${esc(r.authorName || '—')}</div><div>Проверяющий</div><div>${esc(r.reviewerName || '—')}</div><div>Создана</div><div>${fmtT(r.created)}</div>
      <div>Взята</div><div>${fmtT(r.taken)}${r.taken ? ' · ждала ' + sdur(r.taken - r.created) : ''}</div><div>Закрыта</div><div>${fmtT(r.closed)}${r.closed && r.taken ? ' · проверка ' + sdur(r.closed - r.taken) : ''}</div></div>
      ${r.comment ? `<label>Комментарий автора</label><div>${esc(r.comment)}</div>` : ''}${r.question ? `<label>Вопрос проверяющего</label><div>${esc(r.question)}</div>` : ''}
      ${actions ? `<div class="row" style="margin:14px 0">${actions}</div><div id="af"></div>` : ''}
      <h2>История</h2><ul class="tl">${(r.events || []).map((e) => `<li><div class="note">${fmtT(e.t)}</div><b>${esc(e.n)}</b> ${ACTION[e.a] || e.a}${e.d ? `<div class="note">${esc(e.d)}</div>` : ''}</li>`).join('')}</ul></aside>`);
    $('.cab', layer).appendChild(d);
    $('#dx', d).onclick = () => { d.remove(); CAB.openId = null; };
    $$('[data-a]', d).forEach((b) => (b.onclick = () => {
      const a = b.dataset.a;
      if (a === 'return' || a === 'answer') {
        $('#af', d).innerHTML = `<label>${a === 'return' ? 'Вопрос или замечание автору' : 'Ваш ответ проверяющему'}</label><textarea id="at"></textarea><div class="row" style="margin-top:8px"><button class="btn primary" id="as">Отправить</button></div>`;
        $('#at', d).focus(); $('#as', d).onclick = () => act(r.id, a, $('#at', d).value);
      } else if (a === 'cancel') { if (confirm('Отменить заявку?')) act(r.id, a); } else act(r.id, a);
    }));
  }

  async function act(id, action, text) {
    const me = await touchUser(meLocal().id, meLocal().name, meLocal().key);
    const fn = { take: () => take(id, me), approve: () => approve(id, me), release: () => release(id, me), return: () => returnWithQuestion(id, me, text), answer: () => authorAnswer(id, me, text), cancel: () => cancel(id, me) }[action];
    try { const r = await fn(); toast(r.ok ? r.message : r.error); } catch (e) { toast(e.message); }
    await cabRender();
    if (CAB && CAB.openId) openReq(CAB.openId);
  }

  // ─────────────────────────── Запуск ───────────────────────────
  window.addEventListener('beforeunload', (e) => { if (pending.size) { e.preventDefault(); e.returnValue = ''; } });
  GM_registerMenuCommand('Открыть кабинет', () => openCabinet());
  GM_registerMenuCommand('Привязать аккаунт Reddy', () => linkDialog());
  GM_registerMenuCommand('ID служебной страницы', () => { const v = prompt('ID служебной страницы Confluence', STORAGE_PAGE_ID); if (v) { GM_setValue('storagePageId', v.trim()); location.reload(); } });

  (async () => {
    if (!/^\d+$/.test(String(STORAGE_PAGE_ID))) { console.warn('[Координатор] не задан ID служебной страницы (меню Tampermonkey → «ID служебной страницы»)'); return; }
    try { confluenceMe = await cf('GET', '/rest/api/user/current'); } catch (e) { return; }
    if (!confluenceMe || confluenceMe.type === 'anonymous') return;
    await checkUpdate().catch(() => {});
    renderFab();
    setInterval(() => checkUpdate().then(() => { if (updateAvailable()) renderFab(); }), 3600000);
    // Служебная страница: по ссылке из бота (#crb-cabinet), при первой настройке и для непривязанных — сразу кабинет; иначе видна инструкция
    if (pageId() === String(STORAGE_PAGE_ID)) {
      let cfg = null; try { cfg = await getConfig(); } catch (e) { /* покажет кабинет с ошибкой */ }
      if (location.hash.includes('crb-cabinet') || !cfg) openCabinet(); // непривязанным — кнопка «Привязать Reddy» в чек-листе
    }
    window.addEventListener('hashchange', () => { if (location.hash.includes('crb-cabinet') && !CAB) openCabinet(); });
    let last = location.href;
    setInterval(() => { if (location.href !== last) { last = location.href; setTimeout(renderFab, 800); } }, 700);
    setInterval(() => { if (CAB && !document.hidden && !CAB.openId && !['settings', 'users'].includes(CAB.tab)) cabRender(); }, 30000);
    dutyLoop();
  })();

  // Для тестов
  if (GM_getValue('testHooks', false)) window.__crb = { handleUpdate, getMeta, getReq, listProps, dutyState, tick, openCabinet, mutate, take, getUser, submit, seniorsView, addUser, approve, release };
})();
