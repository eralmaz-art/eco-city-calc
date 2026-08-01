/**
 * Калькулятор рассрочки — ЖК «Эко Сити Айни». Версия по ТЗ «Дополнение, ред. 1.9».
 *
 * ВСЯ математика графика — в файле «calc» (модуль EcoCalc, общий с офлайн-PWA
 * и тестами; в репозитории — docs/calc.js). Здесь и в Index математику
 * НЕ дублировать: интерфейс собирает ввод, считает только EcoCalc.build().
 *
 * Этап 1: режим «Прикидка» — расчёт целиком на клиенте, файл не читается.
 * Этап 2: режим «По квартире» — чтение листа «Объекты» и параметров
 *         из «Справочников» строго read-only (ТЗ, раздел 4).
 * Этап 3: «Зафиксировать» (ТЗ, раздел 6) — единственная запись в файл:
 *         построчный график в лист «График» (№ договора / дата / сумма / тип)
 *         + справочные параметры в строке договора (ПВ, дата 1-го платежа, срок).
 *         Семантика с ред. 1.9: «ПВ» = сумма ВСЕХ стартовых платежей;
 *         «Дата 1-го платежа» = якорь ежемесячной серии (первый ежемесячный);
 *         типы строк графика — из словаря calc.js («Первоначальный взнос»,
 *         «Аванс», «Договорённый», «Ежемесячный», «Финальный»).
 *         «Ежемес. платёж» в «Договорах» — формула, НЕ перезаписывается.
 *         Статус графика — «на согласовании»; утверждение (дату) ставит владелец.
 *
 * Владелец скрипта: eralmaz@gmail.com.
 * Доступ менеджерам — через «Поделиться», без хардкода почт (ТЗ, раздел 3).
 *
 * Рекомендуемый вариант: скрипт привязан к боевому файлу продаж
 * (Расширения → Apps Script) — тогда SPREADSHEET_ID оставить пустым.
 * Для автономного скрипта укажите ID файла продаж.
 */

var SPREADSHEET_ID = '';

/** Значения по умолчанию, если параметры не найдены в «Справочниках». */
var PARAM_DEFAULTS = { basePrice: 0, tailLimitPct: 30 };

var SHEET_OBJECTS = 'Объекты';
var SHEET_REFS = 'Справочники';
var SHEET_CONTRACTS = 'Договоры';
var SHEET_SCHEDULE = 'График';

/* «График»: A № договора, B дата, C сумма — по файлу; тип платежа пишем в D
   (заголовок добавляется при первой фиксации, формулы файла его не используют). */
var SCHEDULE_TYPE_HEADER = 'Тип платежа';
/* Словарь типов платежа = словарь calc.js (ТЗ 7.7): не из словаря — отклоняем. */
var PAYMENT_TYPES = ['Первоначальный взнос', 'Аванс', 'Зачёт брони',
                     'Договорённый', 'Ежемесячный', 'Финальный'];
/* Отметка согласования — в строке договора (решение по открытому вопросу ТЗ 6):
   колонки добавляются в конец «Договоров», если их ещё нет. */
var FIX_STATUS_HEADER = 'График: статус';
var FIX_APPROVED_HEADER = 'График: утверждено (дата)';
/* Формулы файла суммируют «График» только до этой строки — за неё не выходим. */
var SCHEDULE_FORMULA_LIMIT = 4962;

function doGet() {
  /* Index — шаблон: внутри <?!= include('calc') ?> подключает модуль математики */
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('Калькулятор рассрочки — Эко Сити Айни')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Подключение HTML-файла проекта в шаблон (стандартный приём Apps Script). */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/**
 * Данные для режима «По квартире»: квартиры + параметры. Только чтение.
 * Вызывается клиентом через google.script.run.
 */
function getCalcData() {
  var ss = getSpreadsheet_();
  return {
    params: readParams_(ss),
    objects: readObjects_(ss),
    generatedAt: Utilities.formatDate(new Date(), 'Asia/Bishkek', 'dd.MM.yyyy HH:mm')
  };
}

function getSpreadsheet_() {
  var ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActive();
  if (!ss) throw new Error('Файл продаж не найден: привяжите скрипт к файлу или укажите SPREADSHEET_ID.');
  return ss;
}

function norm_(v) {
  return String(v == null ? '' : v).toLowerCase().replace(/\s+/g, ' ').trim();
}

/**
 * «Справочники»: подпись параметра в ячейке, значение — в соседней справа.
 * Ищем по подписи, чтобы не зависеть от адресов ячеек.
 * Сверено с реальным файлом (выгрузка v4.1 от 12.07.26): параметр цены называется
 * «Базовый прайс квартир, $/м²»; «Порог хвоста, %» добавляется владельцем при
 * оживлении прайса (ТЗ, раздел 10, п. 1) — до этого действует значение по умолчанию.
 * Заодно читаем список менеджеров (колонка «Менеджеры…») для подсказки на КП.
 */
function readParams_(ss) {
  var out = {
    basePrice: PARAM_DEFAULTS.basePrice,
    tailLimitPct: PARAM_DEFAULTS.tailLimitPct,
    managers: []
  };
  var sh = ss.getSheetByName(SHEET_REFS);
  if (!sh) return out;
  var vals = sh.getDataRange().getValues();
  var mgrCol = -1;
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      var label = norm_(vals[r][c]);
      if (!label) continue;
      var next = c + 1 < vals[r].length ? vals[r][c + 1] : '';
      if ((label.indexOf('базовый прайс квартир') === 0 || label.indexOf('базовая цена') === 0) &&
          typeof next === 'number' && next > 0) out.basePrice = next;
      if (label.indexOf('порог хвоста') === 0 && typeof next === 'number' && next > 0) out.tailLimitPct = next;
      if (r === 0 && label.indexOf('менеджеры') === 0) mgrCol = c;
    }
  }
  if (mgrCol > -1) {
    for (var i = 1; i < vals.length; i++) {
      var name = String(vals[i][mgrCol] == null ? '' : vals[i][mgrCol]).trim();
      if (name) out.managers.push(name);
    }
  }
  return out;
}

/** Лист «Объекты»: в v1.0 отдаём только тип «Квартира» (ТЗ, раздел 3). */
function readObjects_(ss) {
  var sh = ss.getSheetByName(SHEET_OBJECTS);
  if (!sh) throw new Error('Лист «' + SHEET_OBJECTS + '» не найден в файле продаж.');
  var vals = sh.getDataRange().getValues();

  var headRow = -1, map = {};
  for (var r = 0; r < Math.min(vals.length, 10); r++) {
    var m = mapHeaders_(vals[r]);
    if (m.code !== undefined && m.type !== undefined && m.area !== undefined) { headRow = r; map = m; break; }
  }
  if (headRow < 0) throw new Error('В листе «' + SHEET_OBJECTS + '» не найдена строка заголовков (Код / Тип / S,м² …).');

  var out = [];
  for (var i = headRow + 1; i < vals.length; i++) {
    var row = vals[i];
    var get = function (k) { return map[k] === undefined ? '' : row[map[k]]; };
    if (norm_(get('type')) !== 'квартира') continue;
    var area = Number(get('area')) || 0;
    if (!String(get('code')) && !area) continue;
    out.push({
      code: String(get('code')),
      floor: Number(get('floor')) || 0,
      entrance: Number(get('entrance')) || 0,
      num: String(get('num')),
      rooms: Number(get('rooms')) || 0,
      area: area,
      price: Number(get('price')) || 0,
      status: String(get('status') == null ? '' : get('status')).trim(),
      contract: String(get('contract') == null ? '' : get('contract')).trim(),
      buyer: String(get('buyer') == null ? '' : get('buyer')).trim(),
      note: String(get('note') == null ? '' : get('note')).trim(),
      reservedUntil: get('reservedUntil') instanceof Date
          ? Utilities.formatDate(get('reservedUntil'), 'Asia/Bishkek', 'dd.MM.yyyy')
          : String(get('reservedUntil') == null ? '' : get('reservedUntil')).trim(),
      reservedBy: String(get('reservedBy') == null ? '' : get('reservedBy')).trim()
    });
  }
  return out;
}

/**
 * Колонки ищем по заголовкам. Сверено с реальным файлом (v4.1 от 12.07.26):
 * «Этаж/Уровень», «Подъезд/Зона», «Площадь, м²», «Прайс, $», «Бронь до (дата)»,
 * «Бронь: кто» — поэтому сопоставление по началу текста, не по точному совпадению.
 */
function mapHeaders_(row) {
  var map = {};
  for (var c = 0; c < row.length; c++) {
    var h = norm_(row[c]);
    if (!h) continue;
    if (h === 'код') map.code = c;
    else if (h === 'тип') map.type = c;
    else if (h.indexOf('этаж') === 0) map.floor = c;
    else if (h.indexOf('подъезд') === 0) map.entrance = c;
    else if (h === '№' || h === 'номер') map.num = c;
    else if (h.indexOf('комнат') === 0) map.rooms = c;
    else if ((h.charAt(0) === 's' && h.indexOf('м') > -1) || h.indexOf('площадь') === 0) map.area = c;
    else if (h.indexOf('прайс') === 0) map.price = c;
    else if (h === 'статус') map.status = c;
    else if (h.indexOf('текущий договор') === 0) map.contract = c;
    else if (h.indexOf('покупатель') === 0) map.buyer = c;
    else if (h.indexOf('примечание') === 0) map.note = c;
    else if (h.indexOf('бронь до') === 0) map.reservedUntil = c;
    else if (h.indexOf('бронь') === 0) map.reservedBy = c;
  }
  return map;
}

/* ==================== ЭТАП 3: «ЗАФИКСИРОВАТЬ» ==================== */

/**
 * Договоры по коду объекта — для диалога подтверждения фиксации.
 * Возвращает № договора, покупателя, статус, форму оплаты, менеджера,
 * сумму договора и число уже записанных строк графика.
 */
function findContract(objectCode) {
  var ss = getSpreadsheet_();
  var cd = readContracts_(ss);
  var target = norm_(objectCode);
  if (!target) throw new Error('Не указан код объекта.');
  var list = [];
  for (var i = cd.headRow + 1; i < cd.vals.length; i++) {
    var row = cd.vals[i];
    var num = String(row[cd.map.num] == null ? '' : row[cd.map.num]).trim();
    if (!num || norm_(row[cd.map.code]) !== target) continue;
    list.push({
      num: num,
      buyer: cell_(row, cd.map.buyer),
      status: cell_(row, cd.map.status),
      form: cell_(row, cd.map.form),
      manager: cell_(row, cd.map.manager),
      amount: typeof row[cd.map.amount] === 'number' ? row[cd.map.amount] : null,
      scheduleRows: 0
    });
  }
  if (list.length) {
    var counts = scheduleCounts_(ss);
    list.forEach(function (c) { c.scheduleRows = counts[c.num] || 0; });
  }
  return { contracts: list };
}

/**
 * Запись согласованного графика. payload:
 * { contractNum, manager, overwrite, total,
 *   down, firstDate 'yyyy-mm-dd', months,
 *   rows: [{ date: 'yyyy-mm-dd', sum, type }] }
 * Возвращает { ok, written, replaced, planToday, overdue, overLimit }
 * или { ok:false, reason:'exists', count } — если график уже есть и overwrite не задан.
 */
function fixSchedule(p) {
  if (!p || !p.contractNum || !p.rows || !p.rows.length) throw new Error('Пустой запрос фиксации.');
  var contractNum = String(p.contractNum).trim();

  /* валидация до захвата блокировки (правила листа «График», ТЗ 7.4/8) */
  var rows = [], sumRows = 0;
  for (var i = 0; i < p.rows.length; i++) {
    var r = p.rows[i];
    var d = parseYmd_(r.date);
    if (!d || d.getFullYear() < 2024 || d.getFullYear() > 2036) {
      throw new Error('Дата платежа «' + r.date + '» вне диапазона 2024–2036.');
    }
    var s = Math.round(Number(r.sum));
    if (!(s > 0)) throw new Error('Сумма платежа должна быть числом больше нуля.');
    var t = String(r.type || '').trim();
    if (PAYMENT_TYPES.indexOf(t) < 0) {
      throw new Error('Тип платежа «' + t + '» не из словаря (' + PAYMENT_TYPES.join(', ') + ') — график отклонён.');
    }
    rows.push([contractNum, d, s, t]);
    sumRows += s;
  }
  if (p.total && Math.abs(sumRows - Math.round(p.total)) > 0) {
    throw new Error('Контрольная сумма графика (' + sumRows + ') не сходится со стоимостью (' + Math.round(p.total) + ').');
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Файл занят другой записью — повторите через минуту.');
  try {
    var ss = getSpreadsheet_();
    var cd = readContracts_(ss);
    var sheetRow = -1;
    for (var j = cd.headRow + 1; j < cd.vals.length; j++) {
      if (String(cd.vals[j][cd.map.num] == null ? '' : cd.vals[j][cd.map.num]).trim() === contractNum) {
        sheetRow = j + 1;
        break;
      }
    }
    if (sheetRow < 0) throw new Error('Договор «' + contractNum + '» не найден в «Договорах».');

    var gsh = ss.getSheetByName(SHEET_SCHEDULE);
    if (!gsh) throw new Error('Лист «' + SHEET_SCHEDULE + '» не найден.');

    /* существующие строки договора */
    var existing = [];
    var colA = gsh.getRange(1, 1, Math.max(gsh.getLastRow(), 1), 1).getValues();
    for (var k = 1; k < colA.length; k++) {
      if (String(colA[k][0]).trim() === contractNum) existing.push(k + 1);
    }
    if (existing.length && !p.overwrite) return { ok: false, reason: 'exists', count: existing.length };
    for (var m = existing.length - 1; m >= 0; m--) gsh.deleteRow(existing[m]);

    if (!String(gsh.getRange(1, 4).getValue()).trim()) gsh.getRange(1, 4).setValue(SCHEDULE_TYPE_HEADER);

    /* последняя занятая строка по колонке A (в F живёт текст ПРАВИЛ — getLastRow ей не верим) */
    colA = gsh.getRange(1, 1, Math.max(gsh.getLastRow(), 1), 1).getValues();
    var last = 1;
    for (var n = colA.length - 1; n >= 0; n--) {
      if (String(colA[n][0]).trim() !== '') { last = n + 1; break; }
    }
    gsh.getRange(last + 1, 1, rows.length, 4).setValues(rows);
    var overLimit = last + rows.length > SCHEDULE_FORMULA_LIMIT;

    /* строка договора: ПВ (P), дата 1-го платежа (Q), срок (R).
       «Ежемес. платёж» (S) — формула файла, не трогаем.
       Менеджер — только если ячейка пуста (менеджер договора главнее автора расчёта). */
    var sh = cd.sh;
    if (cd.map.down !== undefined) sh.getRange(sheetRow, cd.map.down + 1).setValue(Math.round(p.down || 0));
    if (cd.map.firstDate !== undefined && p.firstDate) {
      sh.getRange(sheetRow, cd.map.firstDate + 1).setValue(parseYmd_(p.firstDate));
    }
    if (cd.map.term !== undefined) sh.getRange(sheetRow, cd.map.term + 1).setValue(Number(p.months) || rows.length);
    if (cd.map.manager !== undefined && p.manager) {
      if (!String(sh.getRange(sheetRow, cd.map.manager + 1).getValue()).trim()) {
        sh.getRange(sheetRow, cd.map.manager + 1).setValue(String(p.manager));
      }
    }
    var fix = ensureFixColumns_(sh, cd);
    sh.getRange(sheetRow, fix.status + 1).setValue('на согласовании');
    if (existing.length) sh.getRange(sheetRow, fix.approved + 1).setValue(''); /* новый график — новое утверждение */

    SpreadsheetApp.flush();

    /* приёмочный критерий ТЗ: просрочка должна ожить — возвращаем живые числа */
    var planToday = cd.map.planToday !== undefined ? sh.getRange(sheetRow, cd.map.planToday + 1).getValue() : '';
    var overdue = cd.map.overdue !== undefined ? sh.getRange(sheetRow, cd.map.overdue + 1).getValue() : '';
    return {
      ok: true,
      written: rows.length,
      replaced: existing.length,
      contract: contractNum,
      planToday: typeof planToday === 'number' ? planToday : null,
      overdue: typeof overdue === 'number' ? overdue : null,
      overLimit: overLimit
    };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- служебные для этапа 3 ---------- */

function cell_(row, idx) {
  return idx === undefined ? '' : String(row[idx] == null ? '' : row[idx]).trim();
}

function parseYmd_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}

function readContracts_(ss) {
  var sh = ss.getSheetByName(SHEET_CONTRACTS);
  if (!sh) throw new Error('Лист «' + SHEET_CONTRACTS + '» не найден.');
  var vals = sh.getDataRange().getValues();
  for (var r = 0; r < Math.min(vals.length, 5); r++) {
    var map = mapContractHeaders_(vals[r]);
    if (map.num !== undefined && map.code !== undefined && map.amount !== undefined) {
      return { sh: sh, vals: vals, headRow: r, map: map };
    }
  }
  throw new Error('В «Договорах» не найдена строка заголовков (№ Дог / Код объекта / Сумма договора …).');
}

/** Заголовки «Договоров» по файлу от 22.07.26 (переносы строк в шапке гасятся norm_). */
function mapContractHeaders_(row) {
  var map = {};
  for (var c = 0; c < row.length; c++) {
    var h = norm_(row[c]);
    if (!h) continue;
    if (h.indexOf('№ дог') === 0) map.num = c;
    else if (h.indexOf('код объекта') === 0) map.code = c;
    else if (h.indexOf('покупатель') === 0) map.buyer = c;
    else if (h.indexOf('сумма договора') === 0) map.amount = c;
    else if (h === 'менеджер') map.manager = c;
    else if (h === 'статус') map.status = c;
    else if (h.indexOf('форма оплаты') === 0) map.form = c;
    else if (h.indexOf('пв') === 0) map.down = c;
    else if (h.indexOf('дата 1-го платежа') === 0) map.firstDate = c;
    else if (h.indexOf('срок') === 0) map.term = c;
    else if (h.indexOf('план на сегодня') === 0) map.planToday = c;
    else if (h.indexOf('просрочка') === 0) map.overdue = c;
    else if (h.indexOf(norm_(FIX_STATUS_HEADER)) === 0) map.fixStatus = c;
    else if (h.indexOf(norm_(FIX_APPROVED_HEADER)) === 0) map.fixApproved = c;
  }
  return map;
}

function scheduleCounts_(ss) {
  var gsh = ss.getSheetByName(SHEET_SCHEDULE);
  var counts = {};
  if (!gsh) return counts;
  var colA = gsh.getRange(1, 1, Math.max(gsh.getLastRow(), 1), 1).getValues();
  for (var i = 1; i < colA.length; i++) {
    var num = String(colA[i][0]).trim();
    if (num) counts[num] = (counts[num] || 0) + 1;
  }
  return counts;
}

/** Колонки отметки согласования в «Договорах»: найти или создать в конце шапки. */
function ensureFixColumns_(sh, cd) {
  if (cd.map.fixStatus !== undefined && cd.map.fixApproved !== undefined) {
    return { status: cd.map.fixStatus, approved: cd.map.fixApproved };
  }
  var headRowNum = cd.headRow + 1;
  var lastCol = sh.getLastColumn();
  var statusCol = cd.map.fixStatus, approvedCol = cd.map.fixApproved;
  if (statusCol === undefined) {
    statusCol = lastCol;
    sh.getRange(headRowNum, statusCol + 1).setValue(FIX_STATUS_HEADER);
    lastCol++;
  }
  if (approvedCol === undefined) {
    approvedCol = lastCol;
    sh.getRange(headRowNum, approvedCol + 1).setValue(FIX_APPROVED_HEADER);
  }
  return { status: statusCol, approved: approvedCol };
}
