// ==UserScript==
// @name         Tickets: проверка ссылок в Support/User files
// @namespace    local.tickets.checker
// @version      2.2
// @description  Меню проверки жалоб: файлы (Support/User), Country, метод по стране, Agent wallet / User's wallet — каждая проверка включается отдельно; обход всех страниц, выгрузка в Excel
// @match        https://th-managment.com/en/admin/backoffice/ExtendedPaymentRequestList
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        GM_setClipboard
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @connect      docs.google.com
// @connect      googleusercontent.com
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  const CFG = {
    ticketHeader: 'ticket id',
    supportHeader: 'support team',   // часть заголовка колонки (в нижнем регистре)
    userHeader: 'user files',        // часть заголовка колонки (в нижнем регистре)
    // тикеты, которые игнорируются (не подсвечиваются и не попадают в список):
    // header — регулярка для заголовка колонки, match — регулярка для её содержимого
    ignore: [
      { name: 'Subagent', header: /(^|\s)sub\s*-?\s*agent/i, match: /hot\s*vouchers\s+new\s+deposit/i },
      { name: 'Inside comment', header: /inside\s+comment/i, match: /(?<![a-zа-яё])test(?![a-zа-яё])/i },
    ],
    // проверка колонки Country: тикет попадает в отдельный список, если колонка пуста или страны нет в списке
    countryHeader: /^country$/i,
    countries: [
      'Morocco', 'Egypt', 'United Arab Emirates', 'Somalia', 'Lebanon', 'Palestine', 'Libya', 'Oman', 'Syria',
      'Kyrgyzstan', 'Tunisia', 'Algeria', 'Sudan', 'South Sudan', 'Mauritania', 'Bahrain', 'Jordan', 'Iraq',
      'Qatar', 'Kuwait', 'Mongolia', 'Saudi Arabia', 'Venezuela', 'Bolivia', 'Nicaragua', 'Yemen', 'Iran',
      'Djibouti', 'Turkey', 'Azerbaijan', 'Malawi', 'Colombia', 'Panama', 'Paraguay', 'Honduras', 'Guatemala',
      'Costa Rica', 'Dominican Republic', 'Jamaica', 'Afghanistan', 'Taiwan', 'Canada',
    ],
    countryAliases: ['UAE', 'Türkiye', 'Turkiye', 'Kyrgyz Republic', 'Syrian Arab Republic', 'Palestinian'],
    // проверка метода: в названии субагента (колонка Subagent) должен быть метод страны из COUNTRY (таблица METHODS ниже)
    methodHeader: /(^|\s)sub\s*-?\s*agent/i,
    aliasTarget: { uae: 'United Arab Emirates', 'türkiye': 'Turkey', turkiye: 'Turkey', 'kyrgyz republic': 'Kyrgyzstan', 'syrian arab republic': 'Syria', palestinian: 'Palestine' },
    // ссылка на Google Таблицу со странами (1-я строка) и методами (под ними). Пусто — используется встроенная таблица METHODS.
    // Доступ к таблице: «Все, у кого есть ссылка» → Читатель. Можно вставить обычную ссылку из адресной строки (с #gid=… для нужного листа).
    methodsUrl: 'https://docs.google.com/spreadsheets/d/1duBI4LGn1qbIoc2yzDNjn7IwvL03qyslBmoT8oAGwjQ/edit?gid=0#gid=0',
    // проверка кошельков: колонки Agent's wallet / User's wallet; пусто = нет текста или заглушка («-», «null», «n/a»…)
    agentWalletHeader: /agent'?s?\s*wallet/i,
    userWalletHeader: /user'?s?\s*wallet/i,
    // если Subagent содержит это (Scratchcard, HotVoucher / Hot Vouchers) — НИ ОДИН кошелёк не проверяется (можно отключить в меню)
    skipWalletsRe: /scratch\s*-?\s*card|hot\s*-?\s*voucher/i,
    pagerSelector: null,             // например '.my-pager' — если автопоиск пагинации не сработает
    nextSelector: null,              // например '.pagination li:last-child a' — если автопоиск кнопки "далее" не сработает
    pageWaitMs: 20000,               // ожидание загрузки страницы (растёт с каждой повторной попыткой)
    retries: 3,                      // сколько раз пробовать перейти на следующую страницу
  };

  let METHODS = {
    "Morocco": ["CIH", "JIBI", "BMCE", "ATTIJARI", "Societe Generale", "Al Barid", "Cash Plus", "Orange Money", "Orange Elite", "L'bankalik", "JIBI base", "BARIDPAY", "CFGBANK", "Wepay Morocco", "Obsidian Morocco Umnia", "Bank-e", "Orange Scratch", "DabaPay", "Pyypl", "Pocket Bank", "CDM Bank", "MT Cash", "Damane Cash", "Chaabi Pay", "Banque Populaire", "CIH Prepayment", "Attijari Prepayment", "Lbankalik Prepayment", "Lbankalik", "CDM", "Wepay by CIH BANK", "Saham", "Attijari", "Cash", "Banque", "Jibi", "We pay"],
    "Egypt": ["Vodafone", "Etisalat", "Orange Cash", "We pay", "InstaPay", "BM Wallet", "Smartwallet", "Telda", "Klivvr", "Qahera cash", "Axis Pay", "VIP Pay", "Voxis Pay", "NBE", "We Pay", "Fawry", "Arig Bank", "Khaan Bank", "State Bank", "QPAY", "Capitron Bank"],
    "United Arab Emirates": ["AlMaryah", "Wise", "NBD", "E&Money", "Emirateslslamic", "ADCB", "DIB", "CBD", "ADIB", "RAKBANK", "AHBank", "AlMashreq", "FAB"],
    "Somalia": ["EVC", "ZAAD", "Sahal", "E-dahab", "Waafi"],
    "Lebanon": ["Whish", "Whish LBP", "Weepay", "Suyool", "W - Transfer QR", "WinkPay", "Neo by Bank Audi", "OMT", "Mahfaza"],
    "Palestine": ["JawwalPay QR", "JawwalPay", "Reflect", "PalPay", "Bank of Palestine", "Bank of Palestine ILS", "PalPay USD", "Maalchat"],
    "Libya": ["NCB", "BT Egypt Libya", "Jumhouria Bank", "Libyana", "EVO Wallet", "NAB", "BCD"],
    "Oman": ["Sohar", "Bank Muscat", "NBO", "Dhofar Bank", "Meethaq Islamic Bank", "Bank Nizwa", "Walletii"],
    "Syria": ["Syriatel", "Lirat", "MTN", "Sham Cash SYP", "Sham Cash USD"],
    "Kyrgyzstan": ["Kyrgyzstan DemirBank", "Kyrgyzstan Bakai Bank", "Kyrgyzstan O!Money", "Kyrgyzstan MBANK", "Kyrgyzstan Optima", "Kyrgyzstan BakaiBank", "Kyrgyzstan SIMBANK", "Kyrgyzstan ABank", "Kyrgyzstan Компаньон", "Kyrgyzstan Элкарт"],
    "Tunisia": ["Sobflous Tunisia", "D17 QR-code Tunisia", "D17 Tunisia", "MyBIAT Tunisia", "Runpay Tunisisa", "Flouci", "IZI", "Swared", "Webank Tunisia", "Payoneer", "ATB", "Individual Payment Service", "RedotPay", "Attijari", "Ooredoo"],
    "Algeria": ["Payoneer", "Individual Payment Service", "BaridiMob Algeria", "Paysera Algeria", "Banxy Algeria", "RedotPay", "Opay", "MyFin", "SofizPay", "Когда приходят", "BaridiMob"],
    "Sudan": ["Bankak", "Fawry", "MyCashi", "Bravo", "Ocash", "MTN MOMO", "mGurush", "Digicash"],
    "Mauritania": ["Bankily", "Mauritania Masrvi", "Mauritania Sedad", "Mauritania BIMBank", "Mauritania Amanty", "Mauritania Click by BNM", "Mauritania Bamis Digital", "Mauritania BCIPay", "Mauritania Gimtel", "MOOV Money", "Mauritania GazaPay", "Mauritania Rassidy", "RedotPay", "Prepayment MRU"],
    "Bahrain": ["BenefitPay", "Bahrain AlSalam", "STCPAY"],
    "Jordan": ["Jordan Orange", "uWallet", "ZainCash", "Cliq", "Dinarak", "Bank al Etihad"],
    "Iraq": ["ZainCash", "Scratchcard", "Asiacell", "Asiacell High Limit", "FIB", "NassWallet", "AsiaPay", "FastPay", "AlTaif Bank", "Zain ScratchCard", "Qi cards", "Korek Scratchcard", "Korek transfer"],
    "Qatar": ["Ooredoo QNB", "Ooredoo MW", "QIIB", "IPay", "Doha Bank Qatar", "Dukhan bank"],
    "Kuwait": ["Al-Tijari", "AUB", "NBK", "GBK", "ABK", "BurganBank", "KFH", "KIB"],
    "Mongolia": ["Arig", "Golomt", "Khaan", "TDB", "XacBank", "State", "M Bank", "Capitron", "Most Money"],
    "Saudi Arabia": ["AlRajhi", "STCPAY", "SNB", "Riyad Bank", "Alfransi Borneo", "UrPay", "MobilyPay", "Barq", "Tiqmo"],
    "Venezuela": ["Banesco", "Banco De Venezuela", "Banco Provincial", "Bancaribe"],
    "Bolivia": ["BNB", "Banco Mercantil", "Yape", "Banco Union", "Tigo"],
    "Nicaragua": ["Banpro", "Lafise"],
    "Yemen": ["Alkuraimi USD", "Jawali", "Jaib", "Jaib minimum"],
    "Iran": ["Bank Transfer", "Hot Vouchers"],
    "Djibouti": ["Waafi", "D-Money", "Saba Bank", "CAC Bank"],
    "Turkey": ["Akbank", "Papara BT", "Ziraat Bank", "YapıKredi Bank", "Vakif bank", "Garanti Bank", "Banka Havalesi", "QNB", "Para QR", "ING Bank", "TEB", "Halk Bank", "VIP HAVALE", "İş Bankası", "Kuveyt Turk", "Al Baraka", "Turkiye Finans", "Enpara", "Ballsh Super Banka Turkey", "Trink Para (Havale/EFT 7/24", "PTT Bank", "HotVoucher", "Paycell", "PayPay", "Jet Transfer", "HIZLI HAVALE", "Premium Havale", "VAKIFBANK", "FAST 7/24", "Ziraat bank", "AkBank", "Yapı Kredi", "Parolapara", "Fast Bank", "Kolay Onay 7/24", "Maxi Havale", "Papara", "Maxi Para", "Deniz Bank", "Papara Multipay new", "Paratim Multipay new", "Paratim paytiqo", "Payco", "FAST + BillBlend", "FAST + paytiqo", "Pep", "Havale Yeni", "Banka Transferi", "Papel", "Hizli Banka Havale-EFT", "Ziraat Bank Seker Bank", "Garanti (BizBet)", "Bankası (BizBet)", "FastPay (BizBet)", "Akbank (BizBet)", "TrinkPara Yeni (BizBet)", "QNB (BizBet)", "Ziraat Bank (BizBet)", "Banka Transferi (BizBet)", "FAST 7/24 BANKA TRANSFERİ", "Yapı Kredi Bank", "Vakif Bank"],
    "Azerbaijan": ["Card To Card 1xBT AZN", "Kapitalbank Astara", "Redir AZN 1xPay Cards BT", "AZN M10 1xRUDNBT", "Baku M10 Azerbaijan", "Astara M10 Azerbaijan", "Astara M10 Azerbaijan FTD", "Astara Azerbaijan BirBank", "LeoBank", "BirBank", "UniBank", "ABB", "expressbank", "HotVoucher"],
    "Malawi": ["Airtel"],
    "Colombia": ["Nequi banco digital", "Daviplata", "Bancolombia"],
    "Panama": ["Yappy", "ACH Xpress", "BAC", "QR Code Yappy", "Punto Pago", "Banistmo", "Banco General"],
    "Paraguay": ["Banco Itau", "UENO Bank", "SIPAP", "BNF"],
    "Honduras": ["BAC", "Banpais", "Ficohsa", "TigoMoney", "Dilo", "Tengo", "Banco Occidente", "Banco Atlantida"],
    "Guatemala": ["Industrial Bank", "GT Continental", "Industrial Transferencia", "Banrural", "BAC", "Zigi"],
    "Jamaica": ["Lynk"],
    "Afghanistan": ["FirstMicroPay", "HesabPay", "ATOMA Pay", "Azipay", "Azizi Bank", "Poolam", "HesabMailPay"],
    "Taiwan": ["Far Eastern International Bank", "TCB", "Chunghwa Post", "Changhua Bank (Не", "Line Bank", "Taishin International Bank", "Next Bank"],
    "Canada": ["TD BANK", "Wise", "Tangerine Bank", "Scotiabank", "National Bank of Canada", "Servus Credit Union", "Desjardins", "Simplii financial", "Royal Bank of Canada", "Interac Bank transfer", "CIBC Bank"],
  };

  const COLORS = {
    both: '#ff4d4f',     // файлы: пусты обе колонки
    user: '#ff9f1a',     // файлы: пуста только User files
    support: '#ffe066',  // файлы: пуста только Support team's files
    country: '#d3adf7',  // страна не из списка
    method: '#91d5ff',   // метод не соответствует стране
    agent: '#ffadd2',    // пуст Agent's wallet
    uwallet: '#87e8de',  // пуст User's wallet
    wallets: '#ff85c0',  // пусты оба кошелька
  };
  // порядок проверок; цвет-«чип» в меню; короткая метка в ячейке
  const KEYS = ['files', 'country', 'method', 'agent', 'user'];
  const CHIP = { files: '#ff4d4f', country: '#d3adf7', method: '#91d5ff', agent: '#ffadd2', user: '#87e8de' };
  const BADGE = { files: 'файлы', country: 'Country', method: 'метод', agent: 'Agent wallet', user: "User's wallet" };
  const COLNAME = { files: "Support team's files / User files", country: 'Country', method: 'Subagent', agent: "Agent's wallet", user: "User's wallet" };

  const norm = (s) => (s || '').replace(/[’`]/g, "'").replace(/\s+/g, ' ').trim().toLowerCase();
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // ---------- настройки меню (сохраняются между сессиями) ----------
  const SKEY = 'tc_settings_v1';
  const DEFAULTS = {
    files: true, filesMode: 'any',   // файлы: any — любая колонка пуста | both — обе | user | support
    country: true, method: true,
    agent: true, user: true, skipScratch: true,   // skipScratch: Scratchcard / HotVoucher — кошельки не проверять
    logic: 'any',                    // any — любая из выбранных проблем; all — только если проблемны все выбранные сразу
    useIgnore: true,                 // не учитывать «игнор»-тикеты (Hot Vouchers New Deposit, test)
    only: false, badges: true, first: true, open: false,
  };
  function loadSettings() {
    let raw = null;
    try { raw = typeof GM_getValue === 'function' ? GM_getValue(SKEY, null) : localStorage.getItem(SKEY); } catch (e) { /* ignore */ }
    try { return Object.assign({}, DEFAULTS, raw ? JSON.parse(raw) : {}); } catch (e) { return Object.assign({}, DEFAULTS); }
  }
  const S = loadSettings();
  // фиксированные правила (в меню не настраиваются; перекрывают значения, сохранённые прежними версиями)
  Object.assign(S, {
    filesMode: 'both',    // файлы: тикет проблемный, только если ПУСТЫ ОБЕ колонки (Support team's files и User files)
    logic: 'any',         // тикет попадает в список, если провалена любая из включённых проверок
    useIgnore: true,      // всегда игнорируем тикеты Hot Vouchers New Deposit и test
    skipScratch: true,    // у Scratchcard и HotVoucher кошельки не проверяются
  });
  function saveSettings() {
    try {
      const raw = JSON.stringify(S);
      if (typeof GM_setValue === 'function') GM_setValue(SKEY, raw); else localStorage.setItem(SKEY, raw);
    } catch (e) { /* ignore */ }
  }

  // ---------- поиск таблицы и колонок ----------
  function locate() {
    const ths = [...document.querySelectorAll('th')];
    const tTh = ths.find((th) => norm(th.innerText).includes(CFG.ticketHeader));
    if (!tTh) return null;
    const headRow = tTh.parentElement;
    const heads = [...headRow.children].map((c) => norm(c.innerText));
    const idx = {
      ticket: heads.findIndex((h) => h.includes(CFG.ticketHeader)),
      support: heads.findIndex((h) => h.includes(CFG.supportHeader)),
      user: heads.findIndex((h) => h.includes(CFG.userHeader)),
    };
    if (idx.ticket < 0) return null;
    idx.country = heads.findIndex((h) => CFG.countryHeader.test(h));
    idx.method = heads.findIndex((h) => CFG.methodHeader.test(h));
    idx.awallet = heads.findIndex((h) => CFG.agentWalletHeader.test(h));
    idx.uwallet = heads.findIndex((h) => CFG.userWalletHeader.test(h));
    idx.ign = CFG.ignore.map((r) => heads.findIndex((h) => r.header.test(h)));

    // строки: сначала из той же таблицы, иначе берём таблицу с наибольшим числом строк
    let rows = [...(tTh.closest('table')?.querySelectorAll('tbody tr') || [])];
    if (!rows.length) {
      const tables = [...document.querySelectorAll('table')];
      tables.sort((a, b) => b.querySelectorAll('tbody tr').length - a.querySelectorAll('tbody tr').length);
      rows = tables[0] ? [...tables[0].querySelectorAll('tbody tr')] : [];
    }
    return { idx, rows };
  }

  function hasLink(cell) {
    if (!cell) return false;
    if (cell.querySelector('a[href]')) return true;
    return /https?:\/\/\S+|\.(jpe?g|png|gif|pdf|docx?|xlsx?|zip|mp4|webp)\b/i.test(cell.innerText || '');
  }

  // ---------- страны ----------
  const escRe = (x) => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
  let COUNTRY_RE = [];   // собирается в buildLookups()
  const countryOk = (t) => COUNTRY_RE.some((re) => re.test(t));
  const countryEmpty = (t) => !t || /^([-–—.?]*|n\/?a|null|none|unknown)$/i.test(t);

  // ---------- методы по странам ----------
  let CANON = [];        // собирается в buildLookups()
  // страна из ячейки -> каноническое название (при совпадении нескольких берём самое длинное: «South Sudan» важнее «Sudan»)
  function canonCountry(t) {
    let best = null;
    CANON.forEach((x) => { if (x.re.test(t) && (!best || x.len > best.len)) best = x; });
    return best ? best.canon : null;
  }
  // нормализация: регистр, пробелы, диакритика (İ/ı/ü…), апострофы
  const mFold = (s) => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/ı/g, 'i')
    .replace(/[’`]/g, "'").replace(/\s+/g, ' ').trim().toLowerCase();
  // ключи метода: полное название и название без хвоста в скобках («Garanti (BizBet)» -> «garanti»)
  const mKeys = (s) => { const a = mFold(s); const b = a.replace(/\s*\(.*$/, '').trim(); return b && b !== a ? [a, b] : [a]; };
  // метод должен входить в название субагента как отдельное слово/фраза («Vodafone» в «Vodafone Alpha»)
  const keyRe = (k) => new RegExp('(^|[^a-z0-9а-яё])' + k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+') + '($|[^a-z0-9а-яё])');
  let METHOD_RES = [];                   // все методы всех стран: {country, len, re}
  let RES_BY_COUNTRY = new Map();        // страна -> её методы
  // пересборка всех справочников (при старте и после загрузки таблицы из Google).
  // Страны из заголовков таблицы добавляются к списку CFG.countries автоматически.
  function buildLookups() {
    const names = CFG.countries.concat(Object.keys(METHODS).filter((k) => !CFG.countries.includes(k)));
    const mk = (name) => new RegExp('(^|[^a-zа-яё])' + escRe(name) + '($|[^a-zа-яё])', 'i');
    COUNTRY_RE = names.concat(CFG.countryAliases).map(mk);
    CANON = names.map((c) => [c, c]).concat(CFG.countryAliases.map((a) => [a, CFG.aliasTarget[a.toLowerCase()] || a]))
      .map(([name, canon]) => ({ re: mk(name), canon, len: name.length }));
    METHOD_RES = [];
    RES_BY_COUNTRY = new Map();
    Object.entries(METHODS).forEach(([country, list]) => {
      const arr = [];
      const seen = new Set();
      list.forEach((m) => mKeys(m).forEach((k) => {
        if (!k || seen.has(k)) return;
        seen.add(k);
        arr.push({ country, len: k.length, re: keyRe(k) });
      }));
      RES_BY_COUNTRY.set(country, arr);
      METHOD_RES.push(...arr);
    });
  }
  buildLookups();

  // ---------- проверки ----------
  const isEmptyVal = (t) => !t || /^([-–—.?]*|n\/?a|null|none|undefined|unknown)$/i.test(t);

  function filesFail(f, mode) {
    switch (mode) {
      case 'both': return !f.sup && !f.usr;
      case 'user': return !f.usr;
      case 'support': return !f.sup;
      default: return !f.sup || !f.usr; // any
    }
  }

  // оценка строки при текущих настройках: ap — какие проверки применимы, fail — какие не пройдены,
  // problem — попадает ли тикет в выборку (логика «любая» / «все выбранные сразу»)
  function evaluate(f, s, cols) {
    const ign = s.useIgnore && f.ignored;
    const ap = {
      files: s.files && cols.files && !ign,
      country: s.country && cols.country && !ign,
      method: s.method && cols.method && !ign && f.mApplicable,
      agent: s.agent && cols.agent && !ign && !(s.skipScratch && f.noWallets),
      user: s.user && cols.user && !ign && !(s.skipScratch && f.noWallets),
    };
    const fail = {
      files: ap.files && filesFail(f, s.filesMode),
      country: ap.country && f.cBad,
      method: ap.method && f.mProblem,
      agent: ap.agent && f.agentEmpty,
      user: ap.user && f.userEmpty,
    };
    const failed = KEYS.filter((k) => fail[k]);
    const nAp = KEYS.filter((k) => ap[k]).length;
    const problem = s.logic === 'all' ? failed.length > 0 && failed.length === nAp : failed.length > 0;
    return { ap, fail, failed, problem };
  }

  // ---------- анализ ----------
  // собирает «сырые» факты по каждой строке (без учёта настроек) — решение «проблема / нет» принимает evaluate()
  function analyze() {
    const loc = locate();
    if (!loc) return null;
    const { idx, rows } = loc;
    const maxIdx = Math.max(idx.ticket, idx.support, idx.user, idx.country, idx.method, idx.awallet, idx.uwallet);
    const cols = {
      files: idx.support >= 0 && idx.user >= 0,
      country: idx.country >= 0,
      method: idx.method >= 0,
      agent: idx.awallet >= 0,
      user: idx.uwallet >= 0,
    };
    const res = [];
    rows.forEach((tr) => {
      const cells = tr.children;
      if (cells.length <= maxIdx) return;
      const id = (cells[idx.ticket].innerText || '').trim();
      if (!id) return;
      const text = (i) => (i >= 0 ? (cells[i].textContent || '').replace(/\s+/g, ' ').trim() : '');
      const f = { tr, id, tCell: cells[idx.ticket], cols };
      f.sup = idx.support >= 0 ? hasLink(cells[idx.support]) : false;
      f.usr = idx.user >= 0 ? hasLink(cells[idx.user]) : false;
      f.ignored = CFG.ignore.some((rule, i) => {
        const c = idx.ign[i] >= 0 ? cells[idx.ign[i]] : null;
        return c && rule.match.test(c.textContent || '');
      });
      f.country = text(idx.country);
      f.cBad = idx.country >= 0 && (countryEmpty(f.country) || !countryOk(f.country));
      // в названии субагента (Subagent) должен присутствовать метод этой страны; без валидной страны / без таблицы для страны — не проверяем
      f.method = text(idx.method);
      f.mCountry = canonCountry(f.country);
      f.mApplicable = idx.method >= 0 && !!f.mCountry && RES_BY_COUNTRY.has(f.mCountry) && !countryEmpty(f.method);
      f.mProblem = false;
      f.mHint = '';
      if (f.mApplicable) {
        const txt = mFold(f.method);
        if (!RES_BY_COUNTRY.get(f.mCountry).some((x) => x.re.test(txt))) {
          f.mProblem = true;
          // подсказка: у каких стран найден метод (берём самые длинные совпадения, чтобы не шуметь короткими словами)
          const hits = METHOD_RES.filter((x) => x.re.test(txt));
          const max = Math.max(0, ...hits.map((x) => x.len));
          f.mHint = hits.length ? [...new Set(hits.filter((x) => x.len === max).map((x) => x.country))].join(', ') : 'метод не найден в таблице';
        }
      }
      f.agent = text(idx.awallet);
      f.agentEmpty = idx.awallet >= 0 && isEmptyVal(f.agent);
      f.uwallet = text(idx.uwallet);
      f.userEmpty = idx.uwallet >= 0 && isEmptyVal(f.uwallet);
      f.noWallets = CFG.skipWalletsRe.test(f.method);
      res.push(f);
    });
    res.cols = cols;
    return res;
  }

  function explain(f, ev) {
    const out = [];
    if (ev.fail.files) out.push('Нет файлов: ' + [!f.sup ? "Support team's files" : '', !f.usr ? 'User files' : ''].filter(Boolean).join(' и '));
    if (ev.fail.country) out.push(f.country ? `Country «${f.country}» не из списка` : 'Country пуст');
    if (ev.fail.method) out.push(`Subagent «${f.method}» не содержит метод страны ${f.mCountry}. Найден метод страны: ${f.mHint}`);
    if (ev.fail.agent) out.push("Agent's wallet пуст");
    if (ev.fail.user) out.push("User's wallet пуст");
    if ((S.agent || S.user) && S.skipScratch && f.noWallets) out.push('Scratchcard / HotVoucher: кошельки не проверяются');
    return out.join('\n');
  }

  function paint(res) {
    res.forEach((f) => {
      const ev = evaluate(f, S, res.cols);
      const sh = ev.problem ? ev.failed : [];
      // приоритет заливки: файлы → кошельки → страна → метод; остальные проблемы — в метках и подсказке
      let color = '';
      if (sh.includes('files')) color = !f.sup && !f.usr ? COLORS.both : !f.usr ? COLORS.user : COLORS.support;
      else if (sh.includes('agent') && sh.includes('user')) color = COLORS.wallets;
      else if (sh.includes('agent')) color = COLORS.agent;
      else if (sh.includes('user')) color = COLORS.uwallet;
      else if (sh.includes('country')) color = COLORS.country;
      else if (sh.includes('method')) color = COLORS.method;
      f.tCell.style.background = color;
      f.tCell.style.color = color === COLORS.both ? '#fff' : '';
      f.tCell.style.fontWeight = color ? '700' : '';
      // фиолетовая рамка — проблема со страной, синяя — метод не соответствует стране (видны и вместе с другой заливкой)
      f.tCell.style.boxShadow = sh.includes('country') ? 'inset 0 0 0 4px #722ed1' : sh.includes('method') ? 'inset 0 0 0 4px #1677ff' : '';
      f.tCell.title = explain(f, ev.problem ? ev : { fail: {} });
      if (S.badges && sh.length) f.tCell.setAttribute('data-tcb', sh.map((k) => BADGE[k]).join(' · '));
      else f.tCell.removeAttribute('data-tcb');
      f.tr.style.display = S.only && !ev.problem ? 'none' : '';
    });
  }

  // ---------- меню ----------
  const style = document.createElement('style');
  style.textContent = `
    #tc-panel, #tc-panel * { box-sizing: border-box; }
    #tc-panel .tc-sec { font-weight: 700; margin: 8px 0 4px; color: #23909f; font-size: 11px; text-transform: uppercase; letter-spacing: .4px; }
    #tc-panel label { display: block; margin-bottom: 4px; cursor: pointer; }
    #tc-panel .tc-sub { margin-left: 22px; font-size: 12px; color: #444; }
    #tc-panel .tc-chip { display: inline-block; width: 11px; height: 11px; border-radius: 2px; margin: 0 5px 0 3px; vertical-align: -1px; border: 1px solid rgba(0,0,0,.3); }
    #tc-panel button { width: 100%; margin-bottom: 4px; }
    #tc-panel select { width: 100%; margin: 2px 0 4px; }
    #tc-panel .tc-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 4px; margin-bottom: 4px; }
    #tc-panel .tc-grid button { margin: 0; padding: 3px 2px; font-size: 12px; }
    #tc-panel a.tc-link { color: #1677ff; cursor: pointer; font-size: 12px; text-decoration: underline; }
    td[data-tcb]::after { content: attr(data-tcb); display: block; margin-top: 2px; font: 600 10px/1.2 Arial, sans-serif; color: inherit; white-space: normal; }
  `;
  document.head.appendChild(style);

  const launcher = document.createElement('button');
  launcher.id = 'tc-launch';
  launcher.type = 'button';
  launcher.title = 'Меню проверки жалоб';
  launcher.style.cssText =
    'position:fixed;right:276px;bottom:12px;z-index:999999;background:#23909f;color:#fff;border:none;border-radius:20px;' +
    'box-sizing:border-box;width:80px;padding:8px 0;text-align:center;white-space:nowrap;font:700 13px Arial,sans-serif;cursor:pointer;box-shadow:0 2px 10px rgba(0,0,0,.35)';
  launcher.innerHTML = 'Файл <span id="tc-badge" style="display:none;margin-left:6px;background:#ff4d4f;border-radius:9px;padding:0 6px;font-size:11px"></span>';
  document.body.appendChild(launcher);

  const chipRow = (k, label) =>
    `<label><input type="checkbox" id="tc-k-${k}"><span class="tc-chip" style="background:${CHIP[k]}"></span>${label}</label>`;
  const panel = document.createElement('div');
  panel.id = 'tc-panel';
  panel.style.cssText =
    'position:fixed;right:276px;bottom:56px;z-index:999999;background:#fff;border:2px solid #23909f;border-radius:8px;' +
    'padding:10px;font:13px/1.4 Arial,sans-serif;width:300px;max-height:calc(100vh - 80px);overflow:auto;' +
    'box-shadow:0 4px 16px rgba(0,0,0,.3);color:#222;display:none';
  panel.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center"><b>Проверка жалоб</b><span id="tc-x" style="cursor:pointer;font-size:16px;padding:0 4px" title="Закрыть меню">✕</span></div>
    <div id="tc-sets">
      <div class="tc-sec">Что проверять</div>
      ${chipRow('files', 'Файлы: нет ссылок')}
      ${chipRow('country', 'Country не из списка')}
      ${chipRow('method', 'Метод не соответствует стране')}
      ${chipRow('agent', "Agent's wallet пуст")}
      ${chipRow('user', "User's wallet пуст")}
      <div><a class="tc-link" id="tc-selall">Все</a> · <a class="tc-link" id="tc-selnone">Ничего</a></div>
      <div class="tc-sec">Отображение</div>
      <label><input type="checkbox" id="tc-only"> Показать только проблемные</label>
      <label><input type="checkbox" id="tc-badges"> Подписи проблем в ячейке</label>
    </div>
    <div class="tc-sec">Результат</div>
    <div id="tc-stat" style="margin-bottom:6px">—</div>
    <div id="tc-src" style="font-size:12px;color:#555;margin-bottom:6px"></div>
    <button id="tc-run">Проверить страницу</button>
    <button id="tc-reload">Обновить таблицу методов</button>
    <div class="tc-sec">Копировать ID</div>
    <div class="tc-grid">
      <button id="tc-c-files">Файлы</button><button id="tc-c-country">Country</button><button id="tc-c-method">Метод</button>
      <button id="tc-c-agent">Agent</button><button id="tc-c-user">User</button><button id="tc-c-all">Все</button>
    </div>
    <div class="tc-sec">Все страницы</div>
    <label><input type="checkbox" id="tc-first"> Начать с 1-й страницы</label>
    <button id="tc-crawl">Все страницы → собрать ID</button>
    <button id="tc-resume" disabled>Продолжить с текущей</button>
    <button id="tc-stop" disabled>Стоп</button>
    <button id="tc-xlsx" disabled>Скачать Excel</button>
    <div id="tc-msg" style="font-size:12px;color:#555"></div>`;
  document.body.appendChild(panel);
  const $ = (s) => panel.querySelector(s);
  const msg = (t) => ($('#tc-msg').textContent = t);

  // меню открывается кнопкой; состояние запоминается
  const showPanel = (v) => { S.open = v; panel.style.display = v ? 'block' : 'none'; saveSettings(); };
  launcher.onclick = () => showPanel(panel.style.display === 'none');
  $('#tc-x').onclick = () => showPanel(false);
  showPanel(!!S.open);

  // настройки <-> элементы меню
  function syncUI() {
    KEYS.forEach((k) => { $('#tc-k-' + k).checked = !!S[k]; $('#tc-c-' + k).disabled = !S[k]; });
    $('#tc-only').checked = !!S.only; $('#tc-badges').checked = !!S.badges;
    $('#tc-first').checked = !!S.first;
  }
  const changed = () => { saveSettings(); syncUI(); run(); };
  KEYS.forEach((k) => { $('#tc-k-' + k).onchange = (e) => { S[k] = e.target.checked; changed(); }; });
  $('#tc-only').onchange = (e) => { S.only = e.target.checked; changed(); };
  $('#tc-badges').onchange = (e) => { S.badges = e.target.checked; changed(); };
  $('#tc-first').onchange = (e) => { S.first = e.target.checked; saveSettings(); };
  $('#tc-selall').onclick = () => { KEYS.forEach((k) => { S[k] = true; }); changed(); };
  $('#tc-selnone').onclick = () => { KEYS.forEach((k) => { S[k] = false; }); changed(); };

  let last = [];
  let running = false;
  function run() {
    const res = analyze();
    if (!res) { $('#tc-stat').textContent = 'Таблица не найдена'; return; }
    last = res;
    paint(res);
    const evs = res.map((f) => evaluate(f, S, res.cols));
    const bad = evs.filter((e) => e.problem).length;
    const parts = KEYS.filter((k) => S[k]).map((k) =>
      `<span class="tc-chip" style="background:${CHIP[k]}"></span>${BADGE[k]}: <b>${evs.filter((e) => e.fail[k]).length}</b>`);
    const ign = S.useIgnore ? res.filter((f) => f.ignored).length : 0;
    const miss = KEYS.filter((k) => S[k] && !res.cols[k]).map((k) => COLNAME[k]);
    const warn = miss.length
      ? `<div style="color:#d4380d;font-size:12px">⚠ Нет колонки: ${miss.join('; ')} — проверка не работает (включите её в Column settings)</div>` : '';
    const none = KEYS.some((k) => S[k]) ? '' : '<div style="color:#d4380d;font-size:12px">⚠ Не выбрана ни одна проверка</div>';
    $('#tc-stat').innerHTML = `Строк: <b>${res.length}</b> · проблемных: <b style="color:#d4380d">${bad}</b>` +
      (ign ? ` · игнор: <b>${ign}</b>` : '') + (parts.length ? '<br>' + parts.join(' · ') : '') + warn + none;
    const badge = document.getElementById('tc-badge');
    badge.textContent = bad; badge.style.display = bad ? 'inline' : 'none';
  }

  function copy(text) {
    if (typeof GM_setClipboard === 'function') GM_setClipboard(text);
    else navigator.clipboard.writeText(text);
  }
  $('#tc-run').onclick = run;

  // список ID: после обхода всех страниц — по всем страницам, иначе — по текущей; решает текущее меню настроек
  function listFor(kind) {
    const all = collected.size > 0;
    const src = all ? [...collected.entries()].map(([id, f]) => ({ id, f })) : (run(), last.map((f) => ({ id: f.id, f })));
    const ids = src.filter(({ f }) => { const ev = evaluate(f, S, f.cols); return ev.problem && (kind === 'all' || ev.fail[kind]); }).map((x) => x.id);
    return { all, ids };
  }
  const copyKind = (kind) => {
    const { all, ids } = listFor(kind);
    copy(ids.join('\n'));
    msg(`Скопировано (${kind === 'all' ? 'все проблемные' : BADGE[kind]}, ${all ? 'все страницы' : 'эта страница'}): ${ids.length}`);
  };
  [...KEYS, 'all'].forEach((k) => { $('#tc-c-' + k).onclick = () => copyKind(k); });

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
      if (panel.contains(el) || el.closest('table')) return false;
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

  let aborted = false;
  let collected = new Map();   // id -> «сырые» факты строки + page (проблемность считается при копировании/выгрузке по текущим настройкам)
  let visited = new Set();     // «from» каждой пройденной страницы
  let pagesDone = 0;

  function setRunning(v) {
    running = v;
    $('#tc-crawl').disabled = v;
    $('#tc-resume').disabled = v || pagesDone === 0;
    $('#tc-stop').disabled = !v;
    $('#tc-xlsx').disabled = v || collected.size === 0;
    // во время обхода настройки заблокированы, чтобы весь обход шёл по одним правилам
    panel.querySelectorAll('#tc-sets input, #tc-sets select').forEach((el) => { el.disabled = v; });
    if (!v) syncUI();
  }

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
    if (!resume) { collected = new Map(); visited = new Set(); pagesDone = 0; }
    setRunning(true);
    let reason = '';
    let note = '';
    let tally = 0;
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

      // 2) обход
      while (!aborted && pagesDone < 1000) {
        run();
        const t = getTotals();
        if (t) { pageSize = Math.max(pageSize, t.to - t.from + 1); visited.add(t.from); }
        pagesDone++;
        last.forEach((f) => {
          const isNew = !collected.has(f.id);
          const { tr, tCell, ...facts } = f;   // DOM-ссылки не храним
          collected.set(f.id, { ...facts, page: pagesDone });
          if (isNew && evaluate(f, S, f.cols).problem) tally++;
        });
        msg(`Страница ${pagesDone}${t ? ` (${t.to} из ${t.total})` : ''} · найдено ID: ${tally}`);
        if (t && t.to >= t.total) { reason = 'последняя страница'; break; }
        const r = await clickAndWait(findNext);
        if (r !== 'ok') { reason = WHY[r] || r; break; }
      }
      if (aborted) reason = WHY.aborted;
    } catch (e) {
      reason = 'ошибка: ' + e.message;
      console.error('[tc]', e);
    }

    run();
    const evs = [...collected.values()].map((f) => evaluate(f, S, f.cols));
    const cnt = (k) => evs.filter((e) => e.problem && e.fail[k]).length;
    const per = KEYS.filter((k) => S[k]).map((k) => `${BADGE[k]}: ${cnt(k)}`).join(', ');
    msg(`${note}Конец: ${reason || '—'}. Страниц: ${pagesDone}. Проблемных: ${evs.filter((e) => e.problem).length}${per ? ` (${per})` : ''} — копируйте нужный список кнопками.${coverage()}`);
    setRunning(false);
  }

  $('#tc-resume').onclick = () => allPages(true);
  $('#tc-crawl').onclick = () => allPages(false);
  $('#tc-stop').onclick = () => { aborted = true; };

  // Excel: «Сводка» (какие проверки не пройдены) + по листу на каждую включённую проверку; всё считается по текущим настройкам
  $('#tc-xlsx').onclick = () => {
    if (typeof XLSX === 'undefined') { msg('Библиотека XLSX не загрузилась.'); return; }
    if (!KEYS.some((k) => S[k])) { msg('Включите хотя бы одну проверку.'); return; }
    const num = (id) => (/^\d{1,15}$/.test(id) ? Number(id) : id);
    const rows = [...collected.entries()].map(([id, f]) => ({ id, f, ev: evaluate(f, S, f.cols) })).filter((x) => x.ev.problem);
    const wb = XLSX.utils.book_new();
    const add = (name, aoa, wch) => {
      const sh = XLSX.utils.aoa_to_sheet(aoa);
      sh['!cols'] = wch.map((w) => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, sh, name);
    };
    const on = KEYS.filter((k) => S[k]);
    add('Сводка', [
      ['ticket_id', ...on.map((k) => BADGE[k]), 'page'],
      ...rows.map((x) => [num(x.id), ...on.map((k) => (x.ev.fail[k] ? 'ПРОБЛЕМА' : '')), x.f.page]),
    ], [14, ...on.map(() => 16), 6]);
    const sheetCounts = [`Сводка — ${rows.length}`];
    if (S.files) {
      const r = rows.filter((x) => x.ev.fail.files);
      add('Нет ссылок', [['ticket_id', 'support_link', 'user_link', 'page'],
        ...r.map((x) => [num(x.id), x.f.sup ? 'yes' : 'NO', x.f.usr ? 'yes' : 'NO', x.f.page])], [14, 14, 12, 6]);
      sheetCounts.push(`«Нет ссылок» — ${r.length}`);
    }
    if (S.country) {
      const r = rows.filter((x) => x.ev.fail.country);
      add('Country', [['ticket_id', 'country', 'page'], ...r.map((x) => [num(x.id), x.f.country || '(пусто)', x.f.page])], [14, 28, 6]);
      sheetCounts.push(`«Country» — ${r.length}`);
    }
    if (S.method) {
      const r = rows.filter((x) => x.ev.fail.method);
      add('Метод', [['ticket_id', 'country', 'subagent', 'найден метод страны', 'page'],
        ...r.map((x) => [num(x.id), x.f.country, x.f.method, x.f.mHint, x.f.page])], [14, 24, 30, 34, 6]);
      sheetCounts.push(`«Метод» — ${r.length}`);
    }
    if (S.agent) {
      const r = rows.filter((x) => x.ev.fail.agent);
      add('Нет кошелька агента', [['ticket_id', 'subagent', 'page'], ...r.map((x) => [num(x.id), x.f.method, x.f.page])], [14, 40, 6]);
      sheetCounts.push(`«Нет кошелька агента» — ${r.length}`);
    }
    if (S.user) {
      const r = rows.filter((x) => x.ev.fail.user);
      add('Нет кошелька пользователя', [['ticket_id', 'subagent', 'page'], ...r.map((x) => [num(x.id), x.f.method, x.f.page])], [14, 40, 6]);
      sheetCounts.push(`«Нет кошелька пользователя» — ${r.length}`);
    }
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'tickets_check_' + new Date().toISOString().slice(0, 10) + '.xlsx';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    msg('Excel: ' + sheetCounts.join(', '));
  };

  // ---------- таблица методов из Google Таблицы ----------
  const EMBEDDED_METHODS = METHODS;
  const setSrc = (t, ok) => { $('#tc-src').textContent = t; $('#tc-src').style.color = ok === false ? '#d4380d' : '#555'; };
  const hhmm = (ts) => new Date(ts).toTimeString().slice(0, 5);

  function sheetCsvUrl(u) {
    u = (u || '').trim();
    if (!u) return '';
    if (/\/document\/d\//.test(u)) return 'DOC';
    if (/[?&](output|format)=csv/.test(u)) return u;
    const m = u.match(/\/spreadsheets\/d\/(?:e\/)?([\w-]+)/);
    if (!m) return u;
    if (/\/spreadsheets\/d\/e\//.test(u)) return u;   // «опубликовано в интернете» — оставляем как есть
    const g = u.match(/[#&?]gid=(\d+)/);
    return `https://docs.google.com/spreadsheets/d/${m[1]}/export?format=csv&gid=${g ? g[1] : 0}`;
  }
  function parseCsv(t) {
    const rows = [];
    let row = [], f = '', q = false;
    for (let i = 0; i < t.length; i++) {
      const ch = t[i];
      if (q) { if (ch === '"') { if (t[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += ch; }
      else if (ch === '"') q = true;
      else if (ch === ',') { row.push(f); f = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && t[i + 1] === '\n') i++; row.push(f); rows.push(row); row = []; f = ''; }
      else f += ch;
    }
    if (f !== '' || row.length) { row.push(f); rows.push(row); }
    return rows;
  }
  // CSV -> {страна: [методы]}; заголовок = страна (алиасы приводятся к каноническому названию)
  function csvToMethods(text) {
    if (/^\s*</.test(text)) throw new Error('вместо таблицы пришла веб-страница — проверьте доступ «Все, у кого есть ссылка»');
    const rows = parseCsv(text.replace(/^\uFEFF/, ''));
    if (rows.length < 2) throw new Error('таблица пустая');
    const out = {};
    rows[0].forEach((h, c) => {
      h = (h || '').replace(/\s+/g, ' ').trim();
      if (!h) return;
      const country = canonCountry(h) || h;
      const list = out[country] || (out[country] = []);
      rows.slice(1).forEach((r) => {
        const v = (r[c] || '').replace(/\s+/g, ' ').trim();
        if (v && !list.includes(v)) list.push(v);
      });
    });
    const n = Object.keys(out).length;
    if (n < 3 || !Object.values(out).some((l) => l.length)) throw new Error('не похоже на таблицу «страны → методы»');
    return out;
  }
  function applyMethods(obj, label) {
    METHODS = obj;
    buildLookups();
    const total = Object.values(obj).reduce((s, l) => s + l.length, 0);
    setSrc(`Методы: ${label} (${Object.keys(obj).length} стран, ${total} методов)`);
    run();
  }
  function loadMethods(manual) {
    const url = sheetCsvUrl(CFG.methodsUrl);
    if (!url) { if (manual) msg('Ссылка не задана: впишите её в CFG.methodsUrl в скрипте'); return; }
    if (url === 'DOC') { setSrc('Нужна Google Таблица (Sheets), а не Документ', false); return; }
    if (typeof GM_xmlhttpRequest !== 'function') { setSrc('GM_xmlhttpRequest недоступен — используется встроенная таблица', false); return; }
    setSrc('Загружаю таблицу методов…');
    GM_xmlhttpRequest({
      method: 'GET', url, timeout: 20000,
      onload: (r) => {
        try {
          if (r.status !== 200) throw new Error('HTTP ' + r.status + (r.status === 401 || r.status === 403 ? ' — нет доступа к таблице' : ''));
          const obj = csvToMethods(r.responseText);
          try { GM_setValue('mfc_methods', JSON.stringify({ ts: Date.now(), obj })); } catch (e) { /* кэш необязателен */ }
          applyMethods(obj, 'Google, ' + hhmm(Date.now()));
          if (manual) msg('Таблица методов обновлена');
        } catch (e) { fallback(e.message); }
      },
      onerror: () => fallback('ошибка сети'),
      ontimeout: () => fallback('таймаут'),
    });
  }
  // не удалось загрузить: берём последнюю удачную копию, иначе встроенную таблицу
  function fallback(reason) {
    let cached = null;
    try { cached = typeof GM_getValue === 'function' ? JSON.parse(GM_getValue('mfc_methods', 'null')) : null; } catch (e) { /* ignore */ }
    if (cached && cached.obj) { applyMethods(cached.obj, 'копия от ' + hhmm(cached.ts)); setSrc($('#tc-src').textContent + ` · Google: ${reason}`, false); }
    else { applyMethods(EMBEDDED_METHODS, 'встроенная'); setSrc($('#tc-src').textContent + ` · Google: ${reason}`, false); }
  }
  $('#tc-reload').onclick = () => loadMethods(true);
  setSrc('Методы: встроенная таблица');
  if (CFG.methodsUrl) loadMethods(false);

  // ---------- автозапуск при изменении таблицы ----------
  let timer = null;
  new MutationObserver((muts) => {
    if (muts.every((m) => panel.contains(m.target) || launcher.contains(m.target))) return;
    clearTimeout(timer);
    timer = setTimeout(run, 400);
  }).observe(document.body, { childList: true, subtree: true });

  syncUI();
  setTimeout(run, 800);
})();
