/**
 * ============================================================
 * ПРОЕКТ: Заявки на ремонт электродвигателей
 * СЕРВЕРНАЯ ЧАСТЬ (Google Apps Script)
 * ВЕРСИЯ: 3.2
 * ============================================================
 *
 * Изменения в 3.2 (скорость):
 *  • Справочники кэшируются на 10 минут — вход быстрее; кэш сбрасывается
 *    при сохранении из сайта и при ручной правке листов справочников
 *  • История читает только нужные строки, а не всю таблицу дважды
 *
 * Изменения в 3.1:
 *  • Несколько паролей: пользователи — B2, C2, D2 …; администраторы — B3, C3 …
 *
 * Изменения относительно 2.2:
 *  • Весь обмен переведён на POST с телом text/plain (без preflight, CORS не ломается)
 *  • Убрано двойное decodeURIComponent — заявки с символом % больше не теряются
 *  • Идемпотентность: повторная отправка одной и той же заявки не создаёт дубль
 *  • LockService на запись — параллельные отправки не получают одинаковый номер
 *  • Пагинация истории (offset/limit), limit=0 — выдать всё (для печати)
 *  • Дата пишется настоящей датой из времени компьютера пользователя
 *  • Мощность и обороты читаются как отображаемый текст (0,37 / 1,10)
 *  • Кэш пароля и последних ID отправки
 *  • updateLists — один атомарный запрос вместо трёх параллельных
 * ============================================================
 */

var CONFIG = {
  SHEET_APPS: 'Заявки',
  SHEET_SETTINGS: 'Настройки',
  SHEET_DEPARTMENTS: 'Подразделения',
  SHEET_LOCATIONS: 'Место установки',
  SHEET_POWERS: 'Мощность',
  SHEET_SPEEDS: 'Обороты',
  SHEET_MANUFACTURERS: 'Производитель',

  MAX_TEXT_LENGTH: 30,
  INV_NUMBER_LENGTH: 5,
  MAX_ITEMS_PER_REQUEST: 100,

  TOTAL_COLS: 13,
  COL_CLIENT_ID: 13,
  DATE_FORMAT: 'dd.MM.yyyy HH:mm',

  LOCK_TIMEOUT_MS: 25000,
  PASSWORD_CACHE_SEC: 300,
  CLIENT_ID_CACHE_SEC: 21600,

  DEFAULT_LIMIT: 30,
  MAX_LIMIT: 5000,

  NOTIFICATION_EMAIL: 'repair@yourcompany.com',
  EMAIL_ENABLED: false,

  LISTS_CACHE_KEY: 'lists_v1',
  LISTS_CACHE_SEC: 600,

  VERSION: '3.2'
};

/* ============================================================
   БЛОК 1: ТОЧКИ ВХОДА
   ============================================================ */

/**
 * GET используется только как проверка доступности сервиса.
 * Открыв URL в браузере, вы должны увидеть {"success":true,...}
 */
function doGet(e) {
  return jsonResponse({
    success: true,
    service: 'Заявки на ремонт электродвигателей',
    version: CONFIG.VERSION,
    message: 'Сервис работает. Обмен данными выполняется методом POST.'
  });
}

/**
 * Единая точка входа. Тело запроса — JSON в text/plain:
 * { action: '...', password: '...', payload: {...} }
 */
function doPost(e) {
  try {
    var body = readBody_(e);

    if (!body || !body.action) {
      return jsonResponse({ success: false, error: 'Не указано действие', code: 'BAD_REQUEST' });
    }

    var action = String(body.action);
    var password = String(body.password || '');
    var payload = body.payload || {};

    var auth = checkPassword(password);
    if (!auth.success) {
      return jsonResponse({ success: false, error: 'Неверный пароль', code: 'AUTH' });
    }

    switch (action) {
      case 'getSettings':
        var settings = getSettings();
        settings.isAdmin = auth.isAdmin;
        return jsonResponse(settings);

      case 'getHistory':
        return jsonResponse(getHistory(payload));

      case 'submit':
        return jsonResponse(submitApplication(payload));

      case 'updateLists':
        if (!auth.isAdmin) {
          return jsonResponse({ success: false, error: 'Требуются права администратора', code: 'FORBIDDEN' });
        }
        return jsonResponse(updateLists(payload));

      default:
        return jsonResponse({ success: false, error: 'Неизвестное действие: ' + action, code: 'BAD_REQUEST' });
    }

  } catch (error) {
    return jsonResponse({ success: false, error: 'Ошибка сервера: ' + error, code: 'SERVER' });
  }
}

function readBody_(e) {
  if (e && e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents);
  }
  if (e && e.parameter && e.parameter.body) {
    return JSON.parse(e.parameter.body);
  }
  return null;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
   БЛОК 2: АВТОРИЗАЦИЯ
   ============================================================ */

function checkPassword(password) {
  if (!password) return { success: false, isAdmin: false };

  try {
    var cache = CacheService.getScriptCache();
    var key = 'pw_' + Utilities.base64Encode(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8));

    var cached = cache.get(key);
    if (cached === 'admin') return { success: true, isAdmin: true };
    if (cached === 'user') return { success: true, isAdmin: false };
    if (cached === 'no') return { success: false, isAdmin: false };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName(CONFIG.SHEET_SETTINGS);
    if (!settingsSheet) return { success: false, isAdmin: false };

    // Пароли пользователей — строка 2 (B2, C2, D2 …), администраторов — строка 3 (B3, C3 …)
    var userPasswords = readPasswordsRow_(settingsSheet, 2);
    var adminPasswords = readPasswordsRow_(settingsSheet, 3);
    var given = String(password).trim();

    var result;
    if (adminPasswords.indexOf(given) !== -1) {
      result = { success: true, isAdmin: true };
      cache.put(key, 'admin', CONFIG.PASSWORD_CACHE_SEC);
    } else if (userPasswords.indexOf(given) !== -1) {
      result = { success: true, isAdmin: false };
      cache.put(key, 'user', CONFIG.PASSWORD_CACHE_SEC);
    } else {
      result = { success: false, isAdmin: false };
      cache.put(key, 'no', 30);
    }

    return result;

  } catch (error) {
    return { success: false, isAdmin: false };
  }
}

/**
 * Читает все непустые пароли из строки листа «Настройки», начиная со столбца B.
 * Пример: B2 = 1111, C2 = 2222, D2 = 3333 → три пароля пользователя.
 * Количество паролей не ограничено — просто заполняйте ячейки вправо.
 */
function readPasswordsRow_(settingsSheet, row) {
  var lastCol = settingsSheet.getLastColumn();
  if (lastCol < 2) return [];

  var values = settingsSheet.getRange(row, 2, 1, lastCol - 1).getDisplayValues()[0];
  var result = [];
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i]).trim();
    if (v !== '') result.push(v);
  }
  return result;
}

/* ============================================================
   БЛОК 3: СПРАВОЧНИКИ
   ============================================================ */

function getSettings() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = null;
    try { cached = cache.get(CONFIG.LISTS_CACHE_KEY); } catch (e) {}
    if (cached) {
      try {
        var fromCache = JSON.parse(cached);
        fromCache.success = true;
        return fromCache;
      } catch (e) {}
    }

    var lists = {
      departments: getListFromSheet_(CONFIG.SHEET_DEPARTMENTS),
      locations: getListFromSheet_(CONFIG.SHEET_LOCATIONS),
      powers: getListFromSheet_(CONFIG.SHEET_POWERS),
      speeds: getListFromSheet_(CONFIG.SHEET_SPEEDS),
      manufacturers: getListFromSheet_(CONFIG.SHEET_MANUFACTURERS)
    };
    try { cache.put(CONFIG.LISTS_CACHE_KEY, JSON.stringify(lists), CONFIG.LISTS_CACHE_SEC); } catch (e) {}

    lists.success = true;
    return lists;
  } catch (error) {
    return { success: false, error: 'Не удалось прочитать справочники: ' + error };
  }
}

function clearListsCache_() {
  try { CacheService.getScriptCache().remove(CONFIG.LISTS_CACHE_KEY); } catch (e) {}
}

/**
 * Простой триггер: срабатывает сам при ручной правке таблицы.
 * Если поменяли лист справочника — сбрасываем кэш, сайт сразу увидит изменения.
 */
function onEdit(e) {
  try {
    var name = e && e.range ? e.range.getSheet().getName() : '';
    var lists = [CONFIG.SHEET_DEPARTMENTS, CONFIG.SHEET_LOCATIONS, CONFIG.SHEET_POWERS,
                 CONFIG.SHEET_SPEEDS, CONFIG.SHEET_MANUFACTURERS];
    if (lists.indexOf(name) !== -1) clearListsCache_();
  } catch (err) {}
}

/**
 * Читает список ОТОБРАЖАЕМЫМИ значениями — «0,37» приходит строкой,
 * а не числом 0.37, поэтому таблица больше не превращает его в дату.
 */
function getListFromSheet_(sheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    var result = [];
    for (var i = 0; i < values.length; i++) {
      var v = String(values[i][0]).trim();
      if (v !== '') result.push(v);
    }
    return result;

  } catch (error) {
    return [];
  }
}

/**
 * Обновление справочников одним запросом.
 * payload: { lists: { departments:[], locations:[], manufacturers:[], powers:[], speeds:[] }, allowEmpty: false }
 */
function updateLists(payload) {
  var lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      return { success: false, error: 'Таблица занята другим пользователем, повторите через минуту', code: 'LOCKED' };
    }

    var lists = payload.lists || {};
    var allowEmpty = payload.allowEmpty === true;

    var map = {
      departments: CONFIG.SHEET_DEPARTMENTS,
      locations: CONFIG.SHEET_LOCATIONS,
      manufacturers: CONFIG.SHEET_MANUFACTURERS,
      powers: CONFIG.SHEET_POWERS,
      speeds: CONFIG.SHEET_SPEEDS
    };

    var updated = [];

    for (var key in map) {
      if (!lists.hasOwnProperty(key)) continue;

      var items = lists[key];
      if (!Array.isArray(items)) continue;

      if (items.length === 0 && !allowEmpty) {
        return { success: false, error: 'Список «' + map[key] + '» пуст. Сохранение отменено.', code: 'EMPTY_LIST' };
      }

      var res = writeList_(map[key], items);
      if (!res.success) return res;
      updated.push(map[key]);
    }

    clearListsCache_();
    return { success: true, updated: updated };

  } catch (error) {
    return { success: false, error: 'Ошибка сохранения справочников: ' + error };
  } finally {
    clearListsCache_(); // даже если сохранилась только часть списков
    try { lock.releaseLock(); } catch (e) {}
  }
}

function writeList_(sheetName, items) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { success: false, error: 'Лист не найден: ' + sheetName };
  }

  var cleaned = [];
  var seen = {};
  for (var i = 0; i < items.length; i++) {
    var v = validateText(items[i], CONFIG.MAX_TEXT_LENGTH);
    if (v === '' || seen[v]) continue;
    seen[v] = true;
    cleaned.push([v]);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  }

  // Текстовый формат — чтобы «1,10» не стало числом или датой
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');

  if (cleaned.length > 0) {
    sheet.getRange(2, 1, cleaned.length, 1).setValues(cleaned);
  }

  return { success: true };
}

/* ============================================================
   БЛОК 4: СОХРАНЕНИЕ ЗАЯВКИ (с защитой от дублей)
   ============================================================ */

/**
 * payload: {
 *   clientId: '1737890000000-k3f9a2',   // ключ идемпотентности
 *   createdAt: '2026-07-28T09:15:00.000Z',
 *   responsible: '...', phone: '...',
 *   items: [ {department, location, motor, invNumber, manufacturer, power, speed, reason}, ... ]
 * }
 */
function submitApplication(payload) {
  var lock = LockService.getScriptLock();

  try {
    var clientId = validateClientId_(payload.clientId);
    if (!clientId) {
      return { success: false, error: 'Не передан идентификатор отправки', code: 'BAD_REQUEST' };
    }

    var responsible = validateText(payload.responsible, CONFIG.MAX_TEXT_LENGTH);
    var phone = validateText(payload.phone, CONFIG.MAX_TEXT_LENGTH);

    if (!responsible) return { success: false, error: 'Не указано ФИО ответственного', code: 'VALIDATION' };
    if (!phone) return { success: false, error: 'Не указан контактный телефон', code: 'VALIDATION' };

    var rawItems = payload.items || [];
    if (!Array.isArray(rawItems) || rawItems.length === 0) {
      return { success: false, error: 'Нет данных для сохранения', code: 'VALIDATION' };
    }
    if (rawItems.length > CONFIG.MAX_ITEMS_PER_REQUEST) {
      return { success: false, error: 'Слишком много позиций в одной заявке (максимум ' + CONFIG.MAX_ITEMS_PER_REQUEST + ')', code: 'VALIDATION' };
    }

    // Дата — время компьютера пользователя
    var createdAt = parseIsoDate_(payload.createdAt) || new Date();

    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      return { success: false, error: 'Сервер занят, повторите отправку', code: 'LOCKED' };
    }

    // --- Проверка дубля ---
    var existing = findExistingSubmission_(clientId);
    if (existing) {
      return {
        success: true,
        duplicate: true,
        applicationNumber: existing.applicationNumber,
        count: existing.count
      };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_APPS);
    if (!sheet) return { success: false, error: 'Лист «' + CONFIG.SHEET_APPS + '» не найден' };

    ensureColumns_(sheet);

    var nextNumber = getNextNumber_(ss, sheet);
    var year = createdAt.getFullYear();

    var rows = [];
    var firstNumber = '';

    for (var i = 0; i < rawItems.length; i++) {
      var item = rawItems[i] || {};
      var motor = validateText(item.motor, CONFIG.MAX_TEXT_LENGTH);
      if (!motor) continue;

      var applicationNumber = year + '-' + padNumber_(nextNumber, 4);
      if (rows.length === 0) firstNumber = applicationNumber;

      rows.push([
        applicationNumber,
        createdAt,
        validateText(item.department, CONFIG.MAX_TEXT_LENGTH),
        validateText(item.location, CONFIG.MAX_TEXT_LENGTH),
        motor,
        validateInvNumber(item.invNumber),
        validateText(item.manufacturer, CONFIG.MAX_TEXT_LENGTH),
        validateText(item.power, 15),
        validateText(item.speed, 15),
        validateText(item.reason, CONFIG.MAX_TEXT_LENGTH),
        responsible,
        phone,
        clientId
      ]);

      nextNumber++;
    }

    if (rows.length === 0) {
      return { success: false, error: 'Ни в одной строке не указана марка двигателя', code: 'VALIDATION' };
    }

    var startRow = sheet.getLastRow() + 1;
    if (startRow < 2) startRow = 2;

    // Форматы ДО записи — иначе «1,10» снова превратится в дату
    sheet.getRange(startRow, 1, rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 2, rows.length, 1).setNumberFormat(CONFIG.DATE_FORMAT);
    sheet.getRange(startRow, 6, rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 8, rows.length, 2).setNumberFormat('@');
    sheet.getRange(startRow, 13, rows.length, 1).setNumberFormat('@');

    sheet.getRange(startRow, 1, rows.length, CONFIG.TOTAL_COLS).setValues(rows);

    // Счётчик
    var settingsSheet = ss.getSheetByName(CONFIG.SHEET_SETTINGS);
    if (settingsSheet) settingsSheet.getRange('B4').setValue(nextNumber);

    SpreadsheetApp.flush();

    rememberSubmission_(clientId, firstNumber, rows.length);

    if (CONFIG.EMAIL_ENABLED) {
      sendEmailNotification_(firstNumber, rows.length, responsible, phone);
    }

    return {
      success: true,
      duplicate: false,
      applicationNumber: firstNumber,
      count: rows.length
    };

  } catch (error) {
    return { success: false, error: 'Ошибка сохранения: ' + error, code: 'SERVER' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Ищет уже сохранённую отправку по clientId.
 * Сначала кэш (быстро), затем поиск по столбцу M.
 */
function findExistingSubmission_(clientId) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('cid_' + clientId);
    if (cached) {
      var parts = cached.split('|');
      return { applicationNumber: parts[0], count: parseInt(parts[1], 10) || 0 };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_APPS);
    if (!sheet) return null;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    if (sheet.getMaxColumns() < CONFIG.COL_CLIENT_ID) return null;

    var range = sheet.getRange(2, CONFIG.COL_CLIENT_ID, lastRow - 1, 1);
    var found = range.createTextFinder(clientId).matchEntireCell(true).findNext();
    if (!found) return null;

    var row = found.getRow();
    var applicationNumber = String(sheet.getRange(row, 1).getDisplayValue());

    // Считаем, сколько строк с этим же clientId
    var all = range.getDisplayValues();
    var count = 0;
    for (var i = 0; i < all.length; i++) {
      if (String(all[i][0]) === clientId) count++;
    }

    rememberSubmission_(clientId, applicationNumber, count);
    return { applicationNumber: applicationNumber, count: count };

  } catch (error) {
    return null;
  }
}

function rememberSubmission_(clientId, applicationNumber, count) {
  try {
    CacheService.getScriptCache().put(
      'cid_' + clientId,
      applicationNumber + '|' + count,
      CONFIG.CLIENT_ID_CACHE_SEC
    );
  } catch (e) {}
}

/**
 * Следующий номер = максимум из счётчика B4 и (последний номер в таблице + 1).
 * Защищает от сбитого счётчика и от дублирующихся номеров.
 */
function getNextNumber_(ss, sheet) {
  var counter = 0;
  var settingsSheet = ss.getSheetByName(CONFIG.SHEET_SETTINGS);
  if (settingsSheet) {
    var v = parseInt(settingsSheet.getRange('B4').getValue(), 10);
    if (!isNaN(v)) counter = v;
  }

  var fromSheet = 0;
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    // Достаточно посмотреть последние 50 строк — номера растут по порядку
    var startRow = Math.max(2, lastRow - 49);
    var numbers = sheet.getRange(startRow, 1, lastRow - startRow + 1, 1).getDisplayValues();
    for (var i = 0; i < numbers.length; i++) {
      var m = String(numbers[i][0]).match(/(\d+)\s*$/);
      if (m) {
        var n = parseInt(m[1], 10);
        if (!isNaN(n) && n > fromSheet) fromSheet = n;
      }
    }
  }

  var next = Math.max(counter, fromSheet + 1);
  if (next < 1) next = 1;
  return next;
}

function ensureColumns_(sheet) {
  if (sheet.getMaxColumns() < CONFIG.TOTAL_COLS) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), CONFIG.TOTAL_COLS - sheet.getMaxColumns());
  }
  if (String(sheet.getRange(1, CONFIG.COL_CLIENT_ID).getValue()).trim() === '') {
    sheet.getRange(1, CONFIG.COL_CLIENT_ID).setValue('ID отправки');
  }
}

/* ============================================================
   БЛОК 5: ИСТОРИЯ ЗАЯВОК (с пагинацией)
   ============================================================ */

/**
 * payload: { filters: {department, location, power, invNumber}, offset: 0, limit: 30 }
 * limit = 0 → вернуть всё (используется для печати)
 */
function getHistory(payload) {
  try {
    var filters = payload.filters || {};
    var offset = parseInt(payload.offset, 10);
    var limit = parseInt(payload.limit, 10);

    if (isNaN(offset) || offset < 0) offset = 0;
    if (isNaN(limit) || limit < 0) limit = CONFIG.DEFAULT_LIMIT;
    if (limit === 0 || limit > CONFIG.MAX_LIMIT) limit = CONFIG.MAX_LIMIT;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_APPS);
    if (!sheet) return { success: false, error: 'Лист «' + CONFIG.SHEET_APPS + '» не найден' };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: [], total: 0, offset: 0, hasMore: false };
    }

    var rowCount = lastRow - 1;

    var fDept = normalizeFilter_(filters.department);
    var fLoc = normalizeFilter_(filters.location);
    var fPower = normalizeFilter_(filters.power);
    var fInv = normalizeFilter_(filters.invNumber).toUpperCase();
    var hasFilters = !!(fDept || fLoc || fPower || fInv);

    // Шаг 1: найти подходящие строки, читая как можно меньше.
    // Без фильтров достаточно столбца A, с фильтрами — столбцов A..H.
    var scan = sheet.getRange(2, 1, rowCount, hasFilters ? 8 : 1).getDisplayValues();
    var matched = [];

    // Идём с конца — новые заявки сверху
    for (var i = rowCount - 1; i >= 0; i--) {
      var num = String(scan[i][0]).trim();
      if (num === '') continue;

      if (hasFilters) {
        if (fDept && String(scan[i][2]).trim() !== fDept) continue;
        if (fLoc && String(scan[i][3]).trim() !== fLoc) continue;
        if (fPower && String(scan[i][7]).trim() !== fPower) continue;
        if (fInv && String(scan[i][5]).trim().toUpperCase().indexOf(fInv) === -1) continue;
      }

      matched.push(i);
    }

    var total = matched.length;
    var slice = matched.slice(offset, offset + limit);
    var data = [];

    // Шаг 2: прочитать полностью только тот кусок таблицы, где лежат нужные строки
    if (slice.length > 0) {
      var minIdx = slice[slice.length - 1];
      var maxIdx = slice[0];
      var span = maxIdx - minIdx + 1;
      var display = sheet.getRange(2 + minIdx, 1, span, 12).getDisplayValues();
      var dates = sheet.getRange(2 + minIdx, 2, span, 1).getValues();

      for (var j = 0; j < slice.length; j++) {
        var k = slice[j] - minIdx;
        var dateValue = dates[k][0];

        data.push({
          number: String(display[k][0]),
          date: (dateValue instanceof Date) ? dateValue.toISOString() : String(display[k][1]),
          department: String(display[k][2]),
          location: String(display[k][3]),
          motor: String(display[k][4]),
          invNumber: String(display[k][5]),
          manufacturer: String(display[k][6]),
          power: String(display[k][7]),
          speed: String(display[k][8]),
          reason: String(display[k][9]),
          responsible: String(display[k][10]),
          phone: String(display[k][11])
        });
      }
    }

    return {
      success: true,
      data: data,
      total: total,
      offset: offset,
      hasMore: (offset + data.length) < total
    };

  } catch (error) {
    return { success: false, error: 'Ошибка загрузки истории: ' + error };
  }
}

function normalizeFilter_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

/* ============================================================
   БЛОК 6: ВАЛИДАЦИЯ
   ============================================================ */

/**
 * Обрезает по длине и убирает управляющие символы.
 * Кавычки и угловые скобки НЕ вырезаются — экранирование выполняется на клиенте.
 */
function validateText(text, maxLength) {
  if (text === null || text === undefined) return '';
  var cleaned = String(text).replace(/[\u0000-\u001F\u007F]/g, ' ').trim();
  cleaned = cleaned.replace(/\s{2,}/g, ' ');
  if (cleaned.length > maxLength) cleaned = cleaned.substring(0, maxLength);
  return cleaned;
}

function validateInvNumber(invNumber) {
  if (!invNumber) return '';
  var cleaned = String(invNumber).trim().toUpperCase().replace(/[^A-ZА-ЯЁ0-9]/g, '');
  if (cleaned.length > CONFIG.INV_NUMBER_LENGTH) {
    cleaned = cleaned.substring(0, CONFIG.INV_NUMBER_LENGTH);
  }
  return cleaned;
}

function validateClientId_(clientId) {
  if (!clientId) return '';
  var cleaned = String(clientId).trim().replace(/[^A-Za-z0-9\-_]/g, '');
  if (cleaned.length < 8 || cleaned.length > 64) return '';
  return cleaned;
}

function parseIsoDate_(iso) {
  if (!iso) return null;
  try {
    var d = new Date(iso);
    if (isNaN(d.getTime())) return null;
    // Отсекаем заведомо неверные даты (сбитые часы на компьютере)
    var year = d.getFullYear();
    if (year < 2020 || year > 2100) return null;
    return d;
  } catch (e) {
    return null;
  }
}

function padNumber_(num, size) {
  var s = String(num);
  while (s.length < size) s = '0' + s;
  return s;
}

/* ============================================================
   БЛОК 7: EMAIL (выключено по умолчанию)
   ============================================================ */

function sendEmailNotification_(appNumber, count, responsible, phone) {
  try {
    var subject = 'Новая заявка на ремонт: ' + appNumber;
    var body = 'Номер заявки: ' + appNumber + '\n' +
               'Ответственный: ' + responsible + '\n' +
               'Телефон: ' + phone + '\n' +
               'Количество двигателей: ' + count + '\n';
    MailApp.sendEmail(CONFIG.NOTIFICATION_EMAIL, subject, body);
  } catch (error) {
    console.log('Ошибка отправки email: ' + error);
  }
}

/* ============================================================
   БЛОК 8: САМОПРОВЕРКА
   Запустите selfTest из редактора, чтобы убедиться,
   что все листы на месте и структура корректна.
   ============================================================ */

function selfTest() {
  var out = ['=== ПРОВЕРКА НАСТРОЙКИ ==='];
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var required = [
    CONFIG.SHEET_APPS, CONFIG.SHEET_SETTINGS, CONFIG.SHEET_DEPARTMENTS,
    CONFIG.SHEET_LOCATIONS, CONFIG.SHEET_POWERS, CONFIG.SHEET_SPEEDS,
    CONFIG.SHEET_MANUFACTURERS
  ];

  for (var i = 0; i < required.length; i++) {
    var sheet = ss.getSheetByName(required[i]);
    out.push((sheet ? '  OK   ' : '  НЕТ  ') + 'лист «' + required[i] + '»');
  }

  var apps = ss.getSheetByName(CONFIG.SHEET_APPS);
  if (apps) {
    out.push('  Столбцов в «Заявки»: ' + apps.getMaxColumns() + (apps.getMaxColumns() >= 13 ? ' (достаточно)' : ' — МАЛО, нужно 13'));
    out.push('  Заголовок M1: «' + apps.getRange(1, 13).getValue() + '»');
    out.push('  Строк с данными: ' + Math.max(apps.getLastRow() - 1, 0));
  }

  var settings = ss.getSheetByName(CONFIG.SHEET_SETTINGS);
  if (settings) {
    var uCount = readPasswordsRow_(settings, 2).length;
    var aCount = readPasswordsRow_(settings, 3).length;
    out.push('  Паролей пользователей (строка 2): ' + uCount + (uCount === 0 ? ' — НЕТ НИ ОДНОГО' : ''));
    out.push('  Паролей администратора (строка 3): ' + aCount + (aCount === 0 ? ' — НЕТ НИ ОДНОГО' : ''));
    out.push('  Следующий номер (B4): ' + settings.getRange('B4').getValue());
  }

  out.push('  Часовой пояс таблицы: ' + ss.getSpreadsheetTimeZone());
  out.push('  Версия серверного кода: ' + CONFIG.VERSION);

  var text = out.join('\n');
  Logger.log(text);
  try {
    SpreadsheetApp.getUi().alert('Проверка настройки', text, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {}
  return text;
}