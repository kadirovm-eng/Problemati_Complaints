// ==UserScript==
// @name         Tickets: проверка ссылок в Support/User files
// @namespace    local.tickets.checker
// @version      1.8
// @description  Подсвечивает Ticket ID: нет ссылок в Support/User files, неверная страна, метод (Subagent) не соответствует стране
// @match        https://managment.io/en/admin/backoffice/ExtendedPaymentRequestList
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
    both: '#ff4d4f',     // пусты обе колонки
    user: '#ff9f1a',     // пуста только User files
    support: '#ffe066',  // пуста только Support team's files
    country: '#d3adf7',  // проблема только со страной
    method: '#91d5ff',   // метод не соответствует стране
  };

  const norm = (s) => (s || '').replace(/[’`]/g, "'").replace(/\s+/g, ' ').trim().toLowerCase();
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

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
    if (idx.ticket < 0 || idx.support < 0 || idx.user < 0) return null;
    idx.country = heads.findIndex((h) => CFG.countryHeader.test(h));
    idx.method = heads.findIndex((h) => CFG.methodHeader.test(h));
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
  let countryOn = true;

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
  let methodOn = true;

  // ---------- анализ ----------
  let mode = 'any';
  let onlyProblems = false;

  function isProblem(r) {
    switch (mode) {
      case 'both': return !r.sup && !r.usr;
      case 'user': return !r.usr;
      case 'support': return !r.sup;
      default: return !r.sup || !r.usr; // any
    }
  }

  function analyze() {
    const loc = locate();
    if (!loc) return null;
    const { idx, rows } = loc;
    const res = [];
    rows.forEach((tr) => {
      const cells = tr.children;
      if (cells.length <= Math.max(idx.ticket, idx.support, idx.user)) return;
      const id = (cells[idx.ticket].innerText || '').trim();
      if (!id) return;
      const r = { tr, id, tCell: cells[idx.ticket], sup: hasLink(cells[idx.support]), usr: hasLink(cells[idx.user]) };
      r.ignored = CFG.ignore.some((rule, i) => {
        const c = idx.ign[i] >= 0 ? cells[idx.ign[i]] : null;
        return c && rule.match.test(c.textContent || '');
      });
      r.problem = !r.ignored && isProblem(r);
      r.country = idx.country >= 0 ? (cells[idx.country].textContent || '').replace(/\s+/g, ' ').trim() : '';
      r.cProblem = countryOn && !r.ignored && idx.country >= 0 && (countryEmpty(r.country) || !countryOk(r.country));
      // в названии субагента (Subagent) должен присутствовать метод этой страны; без валидной страны / без таблицы для страны — не проверяем
      r.method = idx.method >= 0 ? (cells[idx.method].textContent || '').replace(/\s+/g, ' ').trim() : '';
      r.mCountry = canonCountry(r.country);
      r.mProblem = false;
      r.mHint = '';
      if (methodOn && !r.ignored && idx.method >= 0 && r.mCountry && RES_BY_COUNTRY.has(r.mCountry) && !countryEmpty(r.method)) {
        const txt = mFold(r.method);
        if (!RES_BY_COUNTRY.get(r.mCountry).some((x) => x.re.test(txt))) {
          r.mProblem = true;
          // подсказка: у каких стран найден метод (берём самые длинные совпадения, чтобы не шуметь короткими словами)
          const hits = METHOD_RES.filter((x) => x.re.test(txt));
          const max = Math.max(0, ...hits.map((x) => x.len));
          r.mHint = hits.length ? [...new Set(hits.filter((x) => x.len === max).map((x) => x.country))].join(', ') : 'метод не найден в таблице';
        }
      }
      res.push(r);
    });
    res.missingCols = CFG.ignore.filter((_, i) => idx.ign[i] < 0).map((r) => r.name);
    if (countryOn && idx.country < 0) res.missingCols.push('Country');
    if (methodOn && idx.method < 0 && !res.missingCols.includes('Subagent')) res.missingCols.push('Subagent');
    return res;
  }

  function paint(res) {
    res.forEach((r) => {
      let color = '';
      if (!r.sup && !r.usr) color = COLORS.both;
      else if (!r.usr) color = COLORS.user;
      else if (!r.sup) color = COLORS.support;
      // красим только то, что считается проблемой в текущем режиме
      if (!r.problem) color = '';
      if (!color && r.cProblem) color = COLORS.country;
      if (!color && r.mProblem) color = COLORS.method;
      r.tCell.style.background = color;
      r.tCell.style.color = color === COLORS.both ? '#fff' : '';
      r.tCell.style.fontWeight = color ? '700' : '';
      // фиолетовая рамка — проблема со страной (видна и вместе с проблемой по файлам)
      // синяя рамка — метод не соответствует стране
      r.tCell.style.boxShadow = r.cProblem ? 'inset 0 0 0 4px #722ed1' : r.mProblem ? 'inset 0 0 0 4px #1677ff' : '';
      r.tCell.title = r.mProblem ? `Subagent «${r.method}» не содержит метод страны ${r.mCountry}. Найден метод страны: ${r.mHint}` : '';
      r.tr.style.display = onlyProblems && !r.problem && !r.cProblem && !r.mProblem ? 'none' : '';
    });
  }

  // ---------- панель ----------
  const panel = document.createElement('div');
  panel.style.cssText =
    'position:fixed;right:12px;bottom:12px;z-index:999999;background:#fff;border:2px solid #23909f;' +
    'border-radius:8px;padding:10px;font:13px/1.4 Arial,sans-serif;width:260px;box-shadow:0 4px 16px rgba(0,0,0,.3);color:#222';
  panel.innerHTML = `
    <div style="font-weight:700;margin-bottom:6px">Проверка файлов</div>
    <select id="mfc-mode" style="width:100%;margin-bottom:6px">
      <option value="any">Любая из колонок пуста</option>
      <option value="both">Обе колонки пусты</option>
      <option value="user">Только User files пуст</option>
      <option value="support">Только Support files пуст</option>
    </select>
    <div id="mfc-stat" style="margin-bottom:6px">—</div>
    <div id="mfc-src" style="font-size:12px;color:#555;margin-bottom:6px"></div>
    <label style="display:block;margin-bottom:6px"><input type="checkbox" id="mfc-only"> Показать только проблемные</label>
    <label style="display:block;margin-bottom:6px"><input type="checkbox" id="mfc-country" checked> Проверять Country</label>
    <label style="display:block;margin-bottom:6px"><input type="checkbox" id="mfc-method" checked> Проверять метод по стране</label>
    <button id="mfc-run" style="width:100%;margin-bottom:4px">Проверить страницу</button>
    <button id="mfc-reload" style="width:100%;margin-bottom:4px">Обновить таблицу методов</button>
    <button id="mfc-copy-f" style="width:100%;margin-bottom:4px">Копировать ID: файлы</button>
    <button id="mfc-copy-c" style="width:100%;margin-bottom:4px">Копировать ID: Country</button>
    <button id="mfc-copy-m" style="width:100%;margin-bottom:4px">Копировать ID: Метод</button>
    <label style="display:block;margin-bottom:6px"><input type="checkbox" id="mfc-first" checked> Начать с 1-й страницы</label>
    <button id="mfc-all" style="width:100%;margin-bottom:4px">Все страницы → собрать ID</button>
    <button id="mfc-resume" style="width:100%;margin-bottom:4px" disabled>Продолжить с текущей</button>
    <button id="mfc-stop" style="width:100%;margin-bottom:4px" disabled>Стоп</button>
    <button id="mfc-xlsx" style="width:100%;margin-bottom:4px" disabled>Скачать Excel (3 листа)</button>
    <button id="mfc-csv" style="width:100%;margin-bottom:4px" disabled>Скачать CSV</button>
    <div id="mfc-msg" style="font-size:12px;color:#555"></div>`;
  document.body.appendChild(panel);
  const $ = (s) => panel.querySelector(s);
  const msg = (t) => ($('#mfc-msg').textContent = t);

  let last = [];
  function run() {
    const res = analyze();
    if (!res) { $('#mfc-stat').textContent = 'Таблица не найдена'; return; }
    last = res;
    paint(res);
    const bad = res.filter((r) => r.problem).length;
    const ign = res.filter((r) => r.ignored).length;
    const warn = res.missingCols && res.missingCols.length
      ? `<div style="color:#d4380d;font-size:12px">⚠ Нет колонки: ${res.missingCols.join(', ')} — правило не работает (включите её в Column settings)</div>` : '';
    $('#mfc-stat').innerHTML = `Строк: <b>${res.length}</b> · без ссылок: <b style="color:#d4380d">${bad}</b> · Country: <b style="color:#722ed1">${res.filter((r) => r.cProblem).length}</b> · Метод: <b style="color:#1677ff">${res.filter((r) => r.mProblem).length}</b> · игнор: <b>${ign}</b>${warn}`;
  }

  function copy(text) {
    if (typeof GM_setClipboard === 'function') GM_setClipboard(text);
    else navigator.clipboard.writeText(text);
  }

  $('#mfc-mode').onchange = (e) => { mode = e.target.value; run(); };
  $('#mfc-only').onchange = (e) => { onlyProblems = e.target.checked; run(); };
  $('#mfc-run').onclick = run;
  $('#mfc-country').onchange = (e) => { countryOn = e.target.checked; run(); };
  $('#mfc-method').onchange = (e) => { methodOn = e.target.checked; run(); };
  // список ID: после обхода всех страниц — по всем страницам, иначе — по текущей
  function listFor(kind) {
    const all = collected.size > 0;
    const src = all
      ? [...collected.entries()].map(([id, v]) => ({ id, fProblem: v.fProblem, cProblem: v.cProblem, mProblem: v.mProblem }))
      : (run(), last.map((r) => ({ id: r.id, fProblem: r.problem, cProblem: r.cProblem, mProblem: r.mProblem })));
    return { all, ids: src.filter((x) => (kind === 'c' ? x.cProblem : kind === 'm' ? x.mProblem : x.fProblem)).map((x) => x.id) };
  }
  const copyKind = (kind) => {
    const { all, ids } = listFor(kind);
    copy(ids.join('\n'));
    msg(`Скопировано (${{ c: 'Country', m: 'Метод', f: 'файлы' }[kind]}, ${all ? 'все страницы' : 'эта страница'}): ${ids.length}`);
  };
  $('#mfc-copy-f').onclick = () => copyKind('f');
  $('#mfc-copy-c').onclick = () => copyKind('c');
  $('#mfc-copy-m').onclick = () => copyKind('m');

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
  let running = false;
  let collected = new Map();   // id -> {sup, usr, page}
  let visited = new Set();     // «from» каждой пройденной страницы
  let pagesDone = 0;

  function setRunning(v) {
    running = v;
    $('#mfc-all').disabled = v;
    $('#mfc-resume').disabled = v || pagesDone === 0;
    $('#mfc-stop').disabled = !v;
    $('#mfc-csv').disabled = v || collected.size === 0;
    $('#mfc-xlsx').disabled = v || collected.size === 0;
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
    try {
      // 1) перейти на первую страницу (не при продолжении)
      if (!resume && $('#mfc-first').checked) {
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
        last.filter((r) => r.problem || r.cProblem || r.mProblem).forEach((r) => collected.set(r.id, { sup: r.sup, usr: r.usr, country: r.country, method: r.method, mHint: r.mHint, fProblem: r.problem, cProblem: r.cProblem, mProblem: r.mProblem, page: pagesDone }));
        msg(`Страница ${pagesDone}${t ? ` (${t.to} из ${t.total})` : ''} · найдено ID: ${collected.size}`);
        if (t && t.to >= t.total) { reason = 'последняя страница'; break; }
        const r = await clickAndWait(findNext);
        if (r !== 'ok') { reason = WHY[r] || r; break; }
      }
      if (aborted) reason = WHY.aborted;
    } catch (e) {
      reason = 'ошибка: ' + e.message;
      console.error('[mfc]', e);
    }

    run();
    const vals = [...collected.values()];
    const nf = vals.filter((v) => v.fProblem).length;
    const nc = vals.filter((v) => v.cProblem).length;
    const nm = vals.filter((v) => v.mProblem).length;
    msg(`${note}Конец: ${reason || '—'}. Страниц: ${pagesDone}. Без файлов: ${nf}, Country: ${nc}, Метод: ${nm} — копируйте нужный список кнопками.${coverage()}`);
    setRunning(false);
  }

  $('#mfc-resume').onclick = () => allPages(true);
  $('#mfc-all').onclick = () => allPages(false);
  $('#mfc-stop').onclick = () => { aborted = true; };
  // Excel: листы «Нет ссылок», «Country», «Метод». Тикет с несколькими проблемами попадает на каждый свой лист.
  $('#mfc-xlsx').onclick = () => {
    if (typeof XLSX === 'undefined') { msg('Библиотека XLSX не загрузилась — скачайте CSV.'); return; }
    const num = (id) => (/^\d{1,15}$/.test(id) ? Number(id) : id);
    const rows = [...collected.entries()].map(([id, v]) => ({ id, ...v }));
    const f = rows.filter((v) => v.fProblem);
    const c = rows.filter((v) => v.cProblem);
    const m = rows.filter((v) => v.mProblem);
    const sh1 = XLSX.utils.aoa_to_sheet([
      ['ticket_id', 'support_link', 'user_link', 'page'],
      ...f.map((v) => [num(v.id), v.sup ? 'yes' : 'NO', v.usr ? 'yes' : 'NO', v.page]),
    ]);
    const sh2 = XLSX.utils.aoa_to_sheet([
      ['ticket_id', 'country', 'page'],
      ...c.map((v) => [num(v.id), v.country || '(пусто)', v.page]),
    ]);
    const sh3 = XLSX.utils.aoa_to_sheet([
      ['ticket_id', 'country', 'subagent', 'найден метод страны', 'page'],
      ...m.map((v) => [num(v.id), v.country, v.method, v.mHint, v.page]),
    ]);
    sh3['!cols'] = [{ wch: 14 }, { wch: 24 }, { wch: 30 }, { wch: 34 }, { wch: 6 }];
    sh1['!cols'] = [{ wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 6 }];
    sh2['!cols'] = [{ wch: 14 }, { wch: 28 }, { wch: 6 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sh1, 'Нет ссылок');
    XLSX.utils.book_append_sheet(wb, sh2, 'Country');
    XLSX.utils.book_append_sheet(wb, sh3, 'Метод');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'tickets_check_' + new Date().toISOString().slice(0, 10) + '.xlsx';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    msg(`Excel: «Нет ссылок» — ${f.length}, «Country» — ${c.length}, «Метод» — ${m.length}`);
  };
  $('#mfc-csv').onclick = () => {
    const q = (x) => '"' + String(x == null ? '' : x).replace(/"/g, '""') + '"';
    const lines = ['ticket_id;files_problem;support_link;user_link;country;country_problem;method;method_problem;method_found_in;page'];
    collected.forEach((v, id) =>
      lines.push([id, v.fProblem ? 'YES' : '', v.sup ? 'yes' : 'NO', v.usr ? 'yes' : 'NO', q(v.country), v.cProblem ? 'YES' : '', q(v.method), v.mProblem ? 'YES' : '', q(v.mHint), v.page].join(';')));
    const blob = new Blob(['﻿' + lines.join('\n')], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'tickets_check.csv';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  };

  // ---------- таблица методов из Google Таблицы ----------
  const EMBEDDED_METHODS = METHODS;
  const setSrc = (t, ok) => { $('#mfc-src').textContent = t; $('#mfc-src').style.color = ok === false ? '#d4380d' : '#555'; };
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
    if (cached && cached.obj) { applyMethods(cached.obj, 'копия от ' + hhmm(cached.ts)); setSrc($('#mfc-src').textContent + ` · Google: ${reason}`, false); }
    else { applyMethods(EMBEDDED_METHODS, 'встроенная'); setSrc($('#mfc-src').textContent + ` · Google: ${reason}`, false); }
  }
  $('#mfc-reload').onclick = () => loadMethods(true);
  setSrc('Методы: встроенная таблица');
  if (CFG.methodsUrl) loadMethods(false);

  // ---------- автозапуск при изменении таблицы ----------
  let timer = null;
  new MutationObserver((muts) => {
    if (muts.every((m) => panel.contains(m.target))) return;
    clearTimeout(timer);
    timer = setTimeout(run, 400);
  }).observe(document.body, { childList: true, subtree: true });

  setTimeout(run, 800);
})();
