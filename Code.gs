/**
 * Синхронизация остатков FBS между Google Таблицей, Wildberries, Ozon и Яндекс Маркетом.
 *
 * Как это работает (каждые 15 минут):
 *  1. Скрипт читает текущие остатки на каждом маркетплейсе по артикулу продавца.
 *  2. Сравнивает их с тем, что было на площадке в прошлый раз (колонки «ВБ / Озон / ЯМ»).
 *     Разница — это продажи (или отмены) на этой площадке.
 *  3. Новый остаток = остаток в таблице + сумма изменений со всех площадок (не меньше 0).
 *     Если вы вручную изменили остаток в таблице, за основу берётся ваше число.
 *  4. Новый остаток записывается в таблицу и отправляется на те площадки,
 *     где он отличается. Площадка-источник продажи уже показывает верное число,
 *     поэтому на неё ничего не отправляется.
 */

// ===================== НАСТРОЙКИ ЛИСТОВ =====================

const SHEET_STOCK = 'Остатки';
const SHEET_SETTINGS = 'Настройки';
const SHEET_LOG = 'Журнал';

const COL = {
  ART: 1,     // A  Артикул продавца
  QTY: 2,     // B  Остаток (главный, можно править вручную)
  WB: 3,      // C  Остаток на ВБ (служебная)
  OZON: 4,    // D  Остаток на Озоне (служебная)
  YM: 5,      // E  Остаток на ЯМ (служебная)
  SYNCED: 6,  // F  Последний синхронизированный остаток (служебная)
  TIME: 7,    // G  Время последнего изменения
  STATUS: 8,  // H  Статус / ошибки
  WB_CHRT: 9, // I  chrtId размера на ВБ (заполняется автоматически)
};
const NUM_COLS = 9;
const HEADERS = ['Артикул', 'Остаток', 'ВБ', 'Озон', 'ЯМ', 'Синхр. (не трогать)',
  'Обновлено', 'Статус', 'ВБ chrtId'];

const MPS = ['WB', 'OZON', 'YM'];
const MP_NAMES = { WB: 'ВБ', OZON: 'Озон', YM: 'ЯМ' };
const MP_COL = { WB: COL.WB, OZON: COL.OZON, YM: COL.YM };

const SETTINGS_ROWS = [
  ['Тестовый режим (ДА/НЕТ)', 'ДА', 'В тестовом режиме на маркетплейсы ничего не отправляется, только запись в «Журнал»'],
  ['ВБ включён (ДА/НЕТ)', 'ДА', ''],
  ['ВБ: ID склада продавца', '', 'Узнать: меню «Синхронизация остатков» → «Проверить подключение»'],
  ['Озон включён (ДА/НЕТ)', 'ДА', ''],
  ['Озон: Client-Id', '', 'Личный кабинет Озон → Настройки → Seller API'],
  ['Озон: ID склада FBS', '', 'Узнать: «Проверить подключение»'],
  ['ЯМ включён (ДА/НЕТ)', 'ДА', ''],
  ['ЯМ: ID кампании (магазина)', '', 'Узнать: «Проверить подключение»'],
  ['Макс. изменение за раз (защита)', '50', 'Если остаток по артикулу меняется сильнее, строка пропускается до проверки человеком'],
];

const URL_WB_CONTENT = 'https://content-api.wildberries.ru';
const URL_WB_MARKET = 'https://marketplace-api.wildberries.ru';
const URL_OZON = 'https://api-seller.ozon.ru';
const URL_YM = 'https://api.partner.market.yandex.ru';

// ===================== МЕНЮ =====================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Синхронизация остатков')
    .addItem('1. Создать листы', 'setupSheets')
    .addItem('2. Ввести ключи API', 'enterApiKeys')
    .addItem('3. Проверить подключение', 'testConnections')
    .addSeparator()
    .addItem('Синхронизировать сейчас', 'syncStocks')
    .addItem('Включить автосинхронизацию (15 мин)', 'enableTrigger')
    .addItem('Выключить автосинхронизацию', 'disableTrigger')
    .addToUi();
}

function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_STOCK);
  if (!sh) sh = ss.insertSheet(SHEET_STOCK);
  sh.getRange(1, 1, 1, NUM_COLS).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(1, COL.WB, 1, 4).setBackground('#eeeeee');
  sh.getRange(1, COL.WB_CHRT).setBackground('#eeeeee');
  sh.getRange('A:A').setNumberFormat('@'); // артикулы как текст (чтобы 0012 не стал 12)

  let st = ss.getSheetByName(SHEET_SETTINGS);
  if (!st) {
    st = ss.insertSheet(SHEET_SETTINGS);
    st.getRange(1, 1, 1, 3).setValues([['Параметр', 'Значение', 'Подсказка']]).setFontWeight('bold');
    st.getRange(2, 1, SETTINGS_ROWS.length, 3).setValues(SETTINGS_ROWS);
    st.setColumnWidth(1, 260); st.setColumnWidth(3, 500);
  }

  let lg = ss.getSheetByName(SHEET_LOG);
  if (!lg) {
    lg = ss.insertSheet(SHEET_LOG);
    lg.getRange(1, 1, 1, 6).setValues([['Время', 'Артикул', 'Источник', 'Было', 'Стало', 'Подробности']])
      .setFontWeight('bold');
    lg.setFrozenRows(1);
  }
  SpreadsheetApp.getUi().alert('Листы созданы. Заполните лист «Настройки», затем введите ключи API.');
}

function enterApiKeys() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const keys = [
    ['WB_TOKEN', 'Токен Wildberries (категории: Контент, Маркетплейс)'],
    ['OZON_API_KEY', 'Api-Key Озон'],
    ['YM_TOKEN', 'API-ключ Яндекс Маркета (доступ: управление товарами / остатками)'],
  ];
  for (const [k, title] of keys) {
    const has = props.getProperty(k) ? ' (уже сохранён — оставьте пустым, чтобы не менять)' : '';
    const r = ui.prompt(title + has, ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    const v = r.getResponseText().trim();
    if (v) props.setProperty(k, v);
  }
  ui.alert('Ключи сохранены в защищённом хранилище скрипта (их не видно в таблице).');
}

function enableTrigger() {
  disableTrigger(true);
  ScriptApp.newTrigger('syncStocks').timeBased().everyMinutes(15).create();
  SpreadsheetApp.getUi().alert('Автосинхронизация включена: каждые 15 минут.');
}

function disableTrigger(silent) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncStocks')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (silent !== true) SpreadsheetApp.getUi().alert('Автосинхронизация выключена.');
}

// ===================== КОНФИГ =====================

function loadConfig() {
  const st = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
  if (!st) throw new Error('Нет листа «Настройки». Выполните «1. Создать листы».');
  const vals = st.getRange(2, 1, Math.max(st.getLastRow() - 1, 1), 2).getValues();
  const m = {};
  vals.forEach(r => { m[String(r[0]).trim()] = String(r[1]).trim(); });
  const yes = v => /^(да|yes|1|true)$/i.test(v || '');
  const props = PropertiesService.getScriptProperties();
  return {
    dryRun: yes(m['Тестовый режим (ДА/НЕТ)']),
    maxDelta: Number(m['Макс. изменение за раз (защита)']) || 0,
    WB: {
      enabled: yes(m['ВБ включён (ДА/НЕТ)']),
      token: props.getProperty('WB_TOKEN'),
      warehouseId: m['ВБ: ID склада продавца'],
    },
    OZON: {
      enabled: yes(m['Озон включён (ДА/НЕТ)']),
      clientId: m['Озон: Client-Id'],
      apiKey: props.getProperty('OZON_API_KEY'),
      warehouseId: m['Озон: ID склада FBS'],
    },
    YM: {
      enabled: yes(m['ЯМ включён (ДА/НЕТ)']),
      token: props.getProperty('YM_TOKEN'),
      campaignId: m['ЯМ: ID кампании (магазина)'],
    },
  };
}

function checkMpConfig(mp, c) {
  if (mp === 'WB' && (!c.token || !c.warehouseId)) return 'не заданы токен или ID склада';
  if (mp === 'OZON' && (!c.apiKey || !c.clientId || !c.warehouseId)) return 'не заданы Client-Id, Api-Key или ID склада';
  if (mp === 'YM' && (!c.token || !c.campaignId)) return 'не заданы токен или ID кампании';
  return '';
}

// ===================== HTTP =====================

function http(method, url, headers, body) {
  const opts = {
    method: method,
    headers: headers,
    muteHttpExceptions: true,
    contentType: 'application/json',
  };
  if (body !== undefined) opts.payload = JSON.stringify(body);
  let lastErr = '';
  for (let attempt = 0; attempt < 4; attempt++) {
    const resp = UrlFetchApp.fetch(url, opts);
    const code = resp.getResponseCode();
    const text = resp.getContentText();
    if (code >= 200 && code < 300) return text ? JSON.parse(text) : {};
    lastErr = 'HTTP ' + code + ': ' + text.slice(0, 500);
    if (code === 429 || code >= 500) { Utilities.sleep(2000 * (attempt + 1)); continue; }
    break;
  }
  throw new Error(lastErr);
}

function chunk(arr, n) {
  const out = [];
  for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n));
  return out;
}

// ===================== WILDBERRIES =====================

function wbHeaders(c) { return { Authorization: c.token }; }

/** Карта vendorCode -> chrtId (только для карточек с одним размером) + список многоразмерных. */
function wbLoadChrtMap(c) {
  const map = {}, multi = {};
  let cursor = { limit: 100 };
  for (let guard = 0; guard < 500; guard++) {
    const r = http('post', URL_WB_CONTENT + '/content/v2/get/cards/list', wbHeaders(c),
      { settings: { cursor: cursor, filter: { withPhoto: -1 } } });
    const cards = r.cards || [];
    cards.forEach(card => {
      const sizes = card.sizes || [];
      if (sizes.length === 1) map[String(card.vendorCode)] = sizes[0].chrtID;
      else if (sizes.length > 1) multi[String(card.vendorCode)] = true;
    });
    if (!r.cursor || cards.length < cursor.limit) break;
    cursor = { limit: 100, updatedAt: r.cursor.updatedAt, nmID: r.cursor.nmID };
  }
  return { map: map, multi: multi };
}

/** Возвращает {артикул: остаток}. Для артикулов без chrtId — не возвращает. */
function wbFetch(c, rows) {
  const need = rows.filter(r => !r.chrt);
  if (need.length) {
    const m = wbLoadChrtMap(c);
    need.forEach(r => {
      if (m.map[r.art]) { r.chrt = String(m.map[r.art]); r.chrtChanged = true; }
      else if (m.multi[r.art]) r.notes.push('ВБ: у карточки несколько размеров — впишите chrtId в колонку I');
    });
  }
  const withChrt = rows.filter(r => r.chrt);
  const byChrt = {};
  withChrt.forEach(r => { byChrt[r.chrt] = r.art; });
  const res = {};
  withChrt.forEach(r => { res[r.art] = 0; }); // нет в ответе = 0
  chunk(Object.keys(byChrt), 1000).forEach(part => {
    const r = http('post', URL_WB_MARKET + '/api/v3/stocks/' + c.warehouseId, wbHeaders(c),
      { chrtIds: part.map(Number) });
    (r.stocks || []).forEach(s => {
      const art = byChrt[String(s.chrtId)];
      if (art !== undefined) res[art] = Number(s.amount) || 0;
    });
  });
  return res;
}

/** updates: [{art, qty, chrt}] -> {art: ошибка или ''} */
function wbPush(c, updates) {
  const out = {};
  chunk(updates.filter(u => u.chrt), 1000).forEach(part => {
    try {
      http('put', URL_WB_MARKET + '/api/v3/stocks/' + c.warehouseId, wbHeaders(c),
        { stocks: part.map(u => ({ chrtId: Number(u.chrt), amount: u.qty })) });
      part.forEach(u => { out[u.art] = ''; });
    } catch (e) {
      part.forEach(u => { out[u.art] = String(e.message); });
    }
  });
  return out;
}

// ===================== OZON =====================

function ozHeaders(c) { return { 'Client-Id': c.clientId, 'Api-Key': c.apiKey }; }

function ozFetch(c, rows) {
  const res = {};
  chunk(rows.map(r => r.art), 1000).forEach(part => {
    let cursor = '';
    for (let guard = 0; guard < 100; guard++) {
      const r = http('post', URL_OZON + '/v4/product/info/stocks', ozHeaders(c),
        { cursor: cursor, filter: { offer_id: part, visibility: 'ALL' }, limit: 1000 });
      (r.items || []).forEach(it => {
        let avail = 0;
        (it.stocks || []).forEach(s => {
          if (s.type === 'fbs') avail += (Number(s.present) || 0) - (Number(s.reserved) || 0);
        });
        res[String(it.offer_id)] = Math.max(0, avail);
      });
      if (!r.cursor || !(r.items || []).length || r.cursor === cursor) break;
      cursor = r.cursor;
    }
  });
  return res; // артикулов, которых нет на Озоне, в ответе не будет
}

function ozPush(c, updates) {
  const out = {};
  chunk(updates, 100).forEach((part, i) => {
    if (i > 0) Utilities.sleep(800); // лимит Озона ~80 запросов/мин
    try {
      const r = http('post', URL_OZON + '/v2/products/stocks', ozHeaders(c),
        { stocks: part.map(u => ({ offer_id: u.art, stock: u.qty, warehouse_id: Number(c.warehouseId) })) });
      part.forEach(u => { out[u.art] = 'нет ответа по товару'; });
      (r.result || []).forEach(x => {
        out[String(x.offer_id)] = x.updated ? '' :
          (x.errors || []).map(e => e.code + ' ' + (e.message || '')).join('; ') || 'не обновлено';
      });
    } catch (e) {
      part.forEach(u => { out[u.art] = String(e.message); });
    }
  });
  return out;
}

// ===================== ЯНДЕКС МАРКЕТ =====================

function ymHeaders(c) { return { 'Api-Key': c.token }; }

function ymFetch(c, rows) {
  const res = {};
  chunk(rows.map(r => r.art), 500).forEach(part => {
    let pageToken = '';
    for (let guard = 0; guard < 100; guard++) {
      let url = URL_YM + '/v2/campaigns/' + c.campaignId + '/offers/stocks?limit=200';
      if (pageToken) url += '&page_token=' + encodeURIComponent(pageToken);
      const r = http('post', url, ymHeaders(c), { offerIds: part });
      const result = r.result || {};
      (result.warehouses || []).forEach(w => {
        (w.offers || []).forEach(o => {
          const art = String(o.offerId);
          let avail = 0;
          (o.stocks || []).forEach(s => { if (s.type === 'AVAILABLE') avail += Number(s.count) || 0; });
          res[art] = (res[art] || 0) + avail;
        });
      });
      pageToken = result.paging && result.paging.nextPageToken;
      if (!pageToken) break;
    }
  });
  // Товар без остатков ЯМ может не вернуть — считаем 0
  rows.forEach(r => { if (res[r.art] === undefined) res[r.art] = 0; });
  return res;
}

function ymPush(c, updates) {
  const out = {};
  chunk(updates, 2000).forEach(part => {
    try {
      http('put', URL_YM + '/v2/campaigns/' + c.campaignId + '/offers/stocks', ymHeaders(c),
        { skus: part.map(u => ({ sku: u.art, items: [{ count: u.qty }] })) });
      part.forEach(u => { out[u.art] = ''; });
    } catch (e) {
      part.forEach(u => { out[u.art] = String(e.message); });
    }
  });
  return out;
}

const FETCHERS = { WB: wbFetch, OZON: ozFetch, YM: ymFetch };
const PUSHERS = { WB: wbPush, OZON: ozPush, YM: ymPush };

// ===================== ЛОГИКА РАСЧЁТА (без API) =====================

function toInt(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return Number.isFinite(n) && n >= 0 && Math.floor(n) === n ? n : NaN;
}

/**
 * Считает новый остаток по одной строке.
 * row: {art, qty, synced, base:{WB,OZON,YM}}   (числа или null)
 * cur: {WB: число|undefined|null, ...}  — текущее значение на площадке
 *      (undefined = товара нет на площадке, null = площадка недоступна в этот раз)
 * Возвращает {newQty, deltas, manual, push:{mp: qty}, error}
 */
function planRow(row, cur, maxDelta) {
  const out = { newQty: row.qty, deltas: {}, manual: false, push: {}, error: '' };
  if (Number.isNaN(row.qty)) { out.error = 'Остаток должен быть целым числом ≥ 0'; return out; }
  if (row.qty === null) { out.error = 'Не указан остаток'; return out; }

  const first = row.synced === null || Number.isNaN(row.synced);
  let sum = 0;
  if (!first) {
    out.manual = row.qty !== row.synced;
    MPS.forEach(mp => {
      const c = cur[mp], b = row.base[mp];
      if (typeof c === 'number' && typeof b === 'number' && c !== b) {
        out.deltas[mp] = c - b;
        sum += c - b;
      }
    });
  }
  if (maxDelta > 0 && Math.abs(sum) > maxDelta) {
    out.error = 'Изменение ' + sum + ' больше лимита защиты (' + maxDelta + '). Проверьте и поправьте вручную.';
    out.deltas = {};
    return out;
  }
  out.newQty = Math.max(0, row.qty + sum);
  MPS.forEach(mp => {
    if (typeof cur[mp] === 'number' && cur[mp] !== out.newQty) out.push[mp] = out.newQty;
  });
  return out;
}

// ===================== ГЛАВНАЯ ФУНКЦИЯ =====================

function syncStocks() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // предыдущий запуск ещё идёт
  try {
    runSync();
  } finally {
    lock.releaseLock();
  }
}

function runSync() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_STOCK);
  if (!sh) throw new Error('Нет листа «Остатки». Выполните «1. Создать листы».');
  const cfg = loadConfig();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const range = sh.getRange(2, 1, lastRow - 1, NUM_COLS);
  const data = range.getValues();
  const now = new Date();

  // 1. Строки таблицы
  const rows = [];
  const seen = {};
  data.forEach((d, i) => {
    const art = String(d[COL.ART - 1]).trim();
    if (!art) return;
    const r = {
      idx: i, art: art,
      qty: toInt(d[COL.QTY - 1]),
      synced: toInt(d[COL.SYNCED - 1]),
      base: { WB: toInt(d[COL.WB - 1]), OZON: toInt(d[COL.OZON - 1]), YM: toInt(d[COL.YM - 1]) },
      chrt: String(d[COL.WB_CHRT - 1]).trim(),
      notes: [],
    };
    if (seen[art]) { r.dup = true; r.notes.push('Дубликат артикула — строка пропущена'); }
    seen[art] = true;
    rows.push(r);
  });
  const active = rows.filter(r => !r.dup);

  // 2. Читаем остатки с площадок
  const cur = {};       // mp -> {art: qty} или null при ошибке
  const mpErrors = {};
  MPS.forEach(mp => {
    const c = cfg[mp];
    if (!c.enabled) { cur[mp] = null; return; }
    const bad = checkMpConfig(mp, c);
    if (bad) { cur[mp] = null; mpErrors[mp] = bad; return; }
    try {
      cur[mp] = FETCHERS[mp](c, active);
    } catch (e) {
      cur[mp] = null;
      mpErrors[mp] = String(e.message);
    }
  });

  // 3. Считаем
  const logs = [];
  const pushes = { WB: [], OZON: [], YM: [] };
  active.forEach(r => {
    const rc = {};
    MPS.forEach(mp => {
      if (cur[mp] === null) rc[mp] = null;
      else rc[mp] = cur[mp][r.art]; // undefined = нет на площадке
    });
    r.cur = rc;
    r.plan = planRow(r, rc, cfg.maxDelta);
    const p = r.plan;
    if (p.error) { r.notes.push(p.error); return; }
    const src = Object.keys(p.deltas);
    if (p.manual || src.length) {
      const what = [];
      if (p.manual) what.push('ручная правка ' + r.synced + '→' + r.qty);
      src.forEach(mp => what.push(MP_NAMES[mp] + ' ' + (p.deltas[mp] > 0 ? '+' : '') + p.deltas[mp]));
      logs.push([now, r.art, src.map(m => MP_NAMES[m]).join(', ') || 'Таблица',
        p.manual ? r.synced : r.qty, p.newQty, what.join('; ')]);
    }
    Object.keys(p.push).forEach(mp => pushes[mp].push({ art: r.art, qty: p.push[mp], chrt: r.chrt }));
  });

  // 4. Отправляем на площадки
  const pushRes = { WB: {}, OZON: {}, YM: {} };
  MPS.forEach(mp => {
    if (!pushes[mp].length) return;
    if (cfg.dryRun) {
      pushes[mp].forEach(u => logs.push([now, u.art, 'ТЕСТ', '', u.qty,
        'Было бы отправлено на ' + MP_NAMES[mp] + ' (сейчас там ' + (cur[mp] || {})[u.art] + ')']));
      return;
    }
    pushRes[mp] = PUSHERS[mp](cfg[mp], pushes[mp]);
  });

  // 5. Пишем обратно в таблицу (защита от правок во время работы скрипта)
  const fresh = sh.getRange(2, COL.QTY, data.length, 1).getValues();
  active.forEach(r => {
    const row = data[r.idx];
    if (r.chrtChanged) row[COL.WB_CHRT - 1] = r.chrt;
    const p = r.plan;
    const status = r.notes.slice();
    MPS.forEach(mp => {
      if (mpErrors[mp]) status.push(MP_NAMES[mp] + ': ' + mpErrors[mp]);
      else if (cfg[mp].enabled && r.cur[mp] === undefined) status.push(MP_NAMES[mp] + ': нет товара');
    });
    if (p.error || cfg.dryRun) { row[COL.STATUS - 1] = status.join(' | ') || (cfg.dryRun ? 'Тестовый режим' : ''); return; }

    const editedMeanwhile = String(fresh[r.idx][0]) !== String(data[r.idx][COL.QTY - 1]);
    if (editedMeanwhile) {
      row[COL.QTY - 1] = fresh[r.idx][0];
      status.push('Изменено вручную во время синхронизации — учтётся в следующий раз');
    } else {
      if (row[COL.QTY - 1] !== p.newQty) { row[COL.QTY - 1] = p.newQty; row[COL.TIME - 1] = now; }
      row[COL.SYNCED - 1] = p.newQty;
    }
    MPS.forEach(mp => {
      const c = r.cur[mp];
      if (typeof c !== 'number') return; // площадка недоступна / товара нет — базу не трогаем
      let base = c;
      if (p.push[mp] !== undefined) {
        const err = pushRes[mp][r.art];
        if (err === '') base = p.push[mp];
        else status.push(MP_NAMES[mp] + ': ошибка отправки: ' + err);
      }
      row[MP_COL[mp] - 1] = base;
    });
    if (p.manual || Object.keys(p.deltas).length) row[COL.TIME - 1] = now;
    row[COL.STATUS - 1] = status.join(' | ') || 'OK';
  });
  rows.filter(r => r.dup).forEach(r => { data[r.idx][COL.STATUS - 1] = r.notes.join(' | '); });

  range.setValues(data);
  if (logs.length) {
    const lg = ss.getSheetByName(SHEET_LOG);
    if (lg) lg.getRange(lg.getLastRow() + 1, 1, logs.length, 6).setValues(logs);
  }
}

// ===================== ПРОВЕРКА ПОДКЛЮЧЕНИЯ =====================

function testConnections() {
  const cfg = loadConfig();
  const lines = [];
  if (cfg.WB.token) {
    try {
      const w = http('get', URL_WB_MARKET + '/api/v3/warehouses', wbHeaders(cfg.WB));
      lines.push('ВБ — OK. Склады: ' + (w || []).map(x => x.id + ' «' + x.name + '»').join(', '));
    } catch (e) { lines.push('ВБ — ошибка: ' + e.message); }
  } else lines.push('ВБ — токен не введён');

  if (cfg.OZON.apiKey && cfg.OZON.clientId) {
    try {
      const w = http('post', URL_OZON + '/v1/warehouse/list', ozHeaders(cfg.OZON), {});
      lines.push('Озон — OK. Склады: ' + (w.result || []).map(x => x.warehouse_id + ' «' + x.name + '»').join(', '));
    } catch (e) { lines.push('Озон — ошибка: ' + e.message); }
  } else lines.push('Озон — не введены Client-Id или Api-Key');

  if (cfg.YM.token) {
    try {
      const w = http('get', URL_YM + '/v2/campaigns', ymHeaders(cfg.YM));
      lines.push('ЯМ — OK. Магазины: ' + (w.campaigns || []).map(x => x.id + ' «' + x.domain + '» (' + x.placementType + ')').join(', '));
    } catch (e) { lines.push('ЯМ — ошибка: ' + e.message); }
  } else lines.push('ЯМ — токен не введён');

  SpreadsheetApp.getUi().alert(lines.join('\n\n'));
}

// Для тестов в Node.js (в Google Apps Script не используется)
if (typeof module !== 'undefined') module.exports = { planRow, toInt, runSync, syncStocks };
