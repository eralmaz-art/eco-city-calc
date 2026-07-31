/**********************************************************************
 * ECO CITY AINI — СИСТЕМА ПРОДАЖ v4.1
 * Скрипт 1: НАСТРОИТЬ_ЗАЩИТУ — защита листов по схеме из Инструкции
 * Скрипт 2: прописью(x) — сумма прописью для ордеров КО-1/КО-2
 *
 * УСТАНОВКА: Расширения → Apps Script → удалить содержимое Code.gs →
 * вставить ВЕСЬ этот текст → сохранить (значок дискеты).
 * Затем обновить страницу таблицы: появится меню «⚙ Сервис».
 * Запустить «⚙ Сервис → Настроить защиту» (один раз разрешить доступ).
 * Запускать защиту должен ВЛАДЕЛЕЦ файла.
 **********************************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙ Сервис')
    .addItem('Настроить защиту', 'НАСТРОИТЬ_ЗАЩИТУ')
    .addItem('Снять всю защиту', 'СНЯТЬ_ВСЮ_ЗАЩИТУ')
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu('📄 ДДУ')
    .addItem('Сгенерировать договор', 'СГЕНЕРИРОВАТЬ_ДДУ')
    .addToUi();
}

/*============================ ЗАЩИТА ================================*/

function НАСТРОИТЬ_ЗАЩИТУ() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];

  СНЯТЬ_ВСЮ_ЗАЩИТУ(true); // начинаем с чистого листа (идемпотентно)

  // 1) Листы целиком — редактирует только владелец
  ['Шахматка кв', 'Шахматка паркинг', 'Коммерция', 'Дашборд',
   'Справочники', 'Инструкция', 'Бартер'].forEach(function (name) {
    var sh = findSheet_(ss, name);
    if (!sh) { log.push('НЕ НАЙДЕН лист: ' + name); return; }
    lockSheet_(sh, 'Только владелец: ' + name, null);
    log.push('Целиком: ' + sh.getName());
  });

  // 2) Объекты — целиком, кроме L (примечание) и M:N (бронь)
  protectExcept_(ss, 'Объекты', ['L2:L', 'M2:N'], log);

  // 3) Договоры — целиком, кроме «синих» столбцов ввода менеджера
  protectExcept_(ss, 'Договоры',
    ['A2:A','B2:B','C2:C','G2:H','J2:K','M2:R','T2:T','Z2:Z','AD2:AE'],
    log);

  // 4) Клиенты — целиком, кроме A:M (ввод менеджера) и O (примечание)
  protectExcept_(ss, 'Клиенты', ['A2:M', 'O2:O'], log);

  // 5) Платежи — открыт для кассира, закрыты ТОЛЬКО формулы C, F, M
  var pl = findSheet_(ss, 'Платежи');
  if (pl) {
    ['C:C', 'F:F', 'M:M'].forEach(function (a1) {
      var p = pl.getRange(a1).protect()
        .setDescription('Платежи: формулы (' + a1 + ') — не редактировать');
      ownerOnly_(p);
    });
    log.push('Платежи: закрыты колонки C, F, M');
  } else { log.push('НЕ НАЙДЕН лист: Платежи'); }

  // 6) КО-1 и КО-2 — целиком, кроме № документа и полей подписи
  ['КО-1', 'КО-2'].forEach(function (name) {
    protectExcept_(ss, name, ['E12:F12', 'B24:G24', 'B25:G25'], log);
  });

  SpreadsheetApp.getUi().alert('Защита настроена:\n\n' + log.join('\n'));
}

function СНЯТЬ_ВСЮ_ЗАЩИТУ(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [SpreadsheetApp.ProtectionType.SHEET, SpreadsheetApp.ProtectionType.RANGE]
    .forEach(function (type) {
      ss.getSheets().forEach(function (sh) {
        sh.getProtections(type).forEach(function (p) {
          if (p.canEdit()) p.remove();
        });
      });
    });
  if (silent !== true) SpreadsheetApp.getUi().alert('Вся защита снята.');
}

/*----------------------- служебные функции -------------------------*/

// Ищет лист по имени, терпимо к лишним пробелам («Шахматка кв » и т.п.)
function findSheet_(ss, name) {
  var target = String(name).trim();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim() === target) return sheets[i];
  }
  return null;
}

// Защитить лист целиком; unprotected — массив A1-диапазонов-исключений
function lockSheet_(sh, description, unprotected) {
  var p = sh.protect().setDescription(description);
  if (unprotected && unprotected.length) {
    p.setUnprotectedRanges(unprotected.map(function (a1) {
      return sh.getRange(a1);
    }));
  }
  ownerOnly_(p);
  return p;
}

function protectExcept_(ss, name, unprotected, log) {
  var sh = findSheet_(ss, name);
  if (!sh) { log.push('НЕ НАЙДЕН лист: ' + name); return; }
  lockSheet_(sh, 'Защита: ' + name + ' (кроме полей ввода)', unprotected);
  log.push(name + ': открыто только ' + unprotected.join(', '));
}

// Оставить право редактирования только владельцу
function ownerOnly_(p) {
  p.addEditor(Session.getEffectiveUser());
  var eds = p.getEditors();
  if (eds && eds.length) p.removeEditors(eds);
  if (p.canDomainEdit && p.canDomainEdit()) p.setDomainEdit(false);
}

/*========================= СУММА ПРОПИСЬЮ ===========================
 * Использование в ячейке:  =прописью($F$17)
 * Пример: 1 234 567,5  →  «Один миллион двести тридцать четыре тысячи
 * пятьсот шестьдесят семь сомов 50 тыйынов»
 *====================================================================*/

function прописью(sum) {
  if (sum === '' || sum === null || sum === undefined) return '';
  var n = Number(sum);
  if (isNaN(n)) return '';
  var minus = n < 0;
  n = Math.abs(Math.round(n * 100) / 100);

  var som = Math.floor(n);
  var tyiyn = Math.round((n - som) * 100);
  if (tyiyn === 100) { som += 1; tyiyn = 0; }

  var words = числоСловами_(som);
  var text = words + ' ' + форма_(som, ['сом', 'сома', 'сомов']) + ' ' +
             ('0' + tyiyn).slice(-2) + ' ' +
             форма_(tyiyn, ['тыйын', 'тыйына', 'тыйынов']);
  if (minus) text = 'минус ' + text;
  return text.charAt(0).toUpperCase() + text.slice(1);
}

function числоСловами_(n) {
  if (n === 0) return 'ноль';
  var ЕД_М = ['', 'один', 'два', 'три', 'четыре', 'пять', 'шесть',
              'семь', 'восемь', 'девять'];
  var ЕД_Ж = ['', 'одна', 'две', 'три', 'четыре', 'пять', 'шесть',
              'семь', 'восемь', 'девять'];
  var ОТ10 = ['десять', 'одиннадцать', 'двенадцать', 'тринадцать',
              'четырнадцать', 'пятнадцать', 'шестнадцать', 'семнадцать',
              'восемнадцать', 'девятнадцать'];
  var ДЕС  = ['', '', 'двадцать', 'тридцать', 'сорок', 'пятьдесят',
              'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто'];
  var СОТ  = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот',
              'шестьсот', 'семьсот', 'восемьсот', 'девятьсот'];
  var РАЗРЯДЫ = [
    null,                                            // единицы
    ['тысяча', 'тысячи', 'тысяч'],                   // женский род
    ['миллион', 'миллиона', 'миллионов'],
    ['миллиард', 'миллиарда', 'миллиардов'],
    ['триллион', 'триллиона', 'триллионов']
  ];

  var parts = [];
  var idx = 0;
  while (n > 0) {
    var t = n % 1000;
    if (t > 0) {
      var ед = (idx === 1) ? ЕД_Ж : ЕД_М; // тысячи — женский род
      var s = [];
      if (СОТ[Math.floor(t / 100)]) s.push(СОТ[Math.floor(t / 100)]);
      var д = t % 100;
      if (д >= 10 && д <= 19) {
        s.push(ОТ10[д - 10]);
      } else {
        if (ДЕС[Math.floor(д / 10)]) s.push(ДЕС[Math.floor(д / 10)]);
        if (ед[д % 10]) s.push(ед[д % 10]);
      }
      if (idx > 0) s.push(форма_(t, РАЗРЯДЫ[idx]));
      parts.unshift(s.join(' '));
    }
    n = Math.floor(n / 1000);
    idx++;
  }
  return parts.join(' ');
}

// Выбор формы слова: 1 сом / 2 сома / 5 сомов
function форма_(n, f) {
  n = Math.abs(n) % 100;
  if (n >= 11 && n <= 14) return f[2];
  n = n % 10;
  if (n === 1) return f[0];
  if (n >= 2 && n <= 4) return f[1];
  return f[2];
}
/* ============ ГЕНЕРАЦИЯ ДДУ ============ */
const ШАБЛОН_ID = '17k-UW2TczbXqdLfZqGrfL5UDF63QPhqMBIdAlNfIQEI';
const ПАПКА_ID  = '1qZApR8LBcNKoTBNHoN-krLCjJFl11lOx';

function СГЕНЕРИРОВАТЬ_ДДУ() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Генерация ДДУ', 'Введите № договора:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const num = String(resp.getResponseText()).trim();
  if (!num) { ui.alert('Не указан номер договора.'); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dg = ss.getSheetByName('Договоры');
  const cl = ss.getSheetByName('Клиенты');
  const ob = ss.getSheetByName('Объекты');
  const gr = ss.getSheetByName('График');

  const dgData = dg.getDataRange().getValues();
  let d = null;
  for (let i = 1; i < dgData.length; i++) {
    if (String(dgData[i][0]).trim() === num) { d = dgData[i]; break; }
  }
  if (!d) { ui.alert('Договор № ' + num + ' не найден в листе «Договоры».'); return; }

  const кодОбъекта = d[1];
  const idКлиента  = d[2];
  const датаДог    = d[6];
  const ценаМ2     = d[7];
  const суммаДог   = d[11];
  const пв         = d[15];
  const дата1      = d[16];
  const срок       = d[17];
  const ежемес     = d[18];

  const clData = cl.getDataRange().getValues();
  let c = null;
  for (let i = 1; i < clData.length; i++) {
    if (String(clData[i][0]).trim() === String(idКлиента).trim()) { c = clData[i]; break; }
  }
  if (!c) { ui.alert('Клиент ' + idКлиента + ' не найден в листе «Клиенты».'); return; }

  const obData = ob.getDataRange().getValues();
  let o = null;
  for (let i = 1; i < obData.length; i++) {
    if (String(obData[i][0]).trim() === String(кодОбъекта).trim()) { o = obData[i]; break; }
  }
  if (!o) { ui.alert('Объект ' + кодОбъекта + ' не найден.'); return; }

  const этаж     = o[2];
  const подъезд  = o[3];
  const номерКв  = o[4];
  const площадь  = o[6];

  let график = [];
  const grData = gr.getDataRange().getValues();
  for (let i = 1; i < grData.length; i++) {
    if (String(grData[i][0]).trim() === num && grData[i][1]) {
      график.push({ дата: grData[i][1], сумма: grData[i][2] });
    }
  }
  if (график.length === 0 && срок && ежемес && дата1) {
    const d1 = new Date(дата1);
    for (let m = 0; m < Number(срок); m++) {
      const dt = new Date(d1.getFullYear(), d1.getMonth() + m, d1.getDate());
      график.push({ дата: dt, сумма: Number(ежемес) });
    }
  }

  let пвИтог = Number(пв || 0);
  if (!пвИтог && график.length > 0) {
    пвИтог = Number(график[0].сумма || 0);
  }
  const остаток = Number(суммаДог) - пвИтог;
  const итогоГрафик = график.reduce((s, x) => s + Number(x.сумма || 0), 0);

  const папка = DriveApp.getFolderById(ПАПКА_ID);
  const имя = 'ДДУ-' + num + ' ' + (c[1] || '') + ' ' + кодОбъекта;
  const копия = DriveApp.getFileById(ШАБЛОН_ID).makeCopy(имя, папка);
  const doc = DocumentApp.openById(копия.getId());
  const body = doc.getBody();

  const fmtD = dt => dt ? Utilities.formatDate(new Date(dt), 'GMT+6', 'dd.MM.yyyy') : '____________';
  const fmtM = n => n ? Number(n).toLocaleString('ru-RU') : '____';

  const map = {
    '{{НОМЕР_ДОГОВОРА}}': num,
    '{{ДАТА_ДОГОВОРА_ПРОПИСЬЮ}}': fmtD(датаДог),
    '{{ДЕНЬ}}': датаДог ? new Date(датаДог).getDate() : '__',
    '{{МЕСЯЦ}}': датаДог ? (new Date(датаДог).getMonth() + 1) : '__',
    '{{ГОД}}': датаДог ? new Date(датаДог).getFullYear() : '____',
    '{{ФИО}}': c[1] || '',
    '{{ФИО_КРАТКО}}': c[1] || '',
    '{{ДАТА_РОЖДЕНИЯ}}': fmtD(c[2]),
    '{{ПИН}}': c[3] || '',
    '{{ID_КАРТА}}': c[4] || '',
    '{{КЕМ_ВЫДАНА}}': c[5] || '',
    '{{ДАТА_ВЫДАЧИ}}': fmtD(c[6]),
    '{{ПРОПИСКА}}': c[7] || '',
    '{{ТЕЛЕФОН}}': c[8] || '',
    '{{ДОМ}}': '77',
    '{{БЛОК}}': подъезд || '',
    '{{ЭТАЖ}}': этаж || '',
    '{{НОМЕР_КВАРТИРЫ}}': номерКв || '',
    '{{ПЛОЩАДЬ}}': площадь || '',
    '{{ПОДЪЕЗД}}': подъезд || '',
    '{{ЦЕНА_ДОГОВОРА_USD}}': fmtM(суммаДог),
    '{{ЦЕНА_ПРОПИСЬЮ}}': fmtM(суммаДог),
    '{{ЦЕНА_ЗА_М2}}': fmtM(ценаМ2),
    '{{ПЕРВОНАЧАЛЬНЫЙ_ВЗНОС}}': fmtM(пвИтог),
    '{{ОСТАТОК}}': fmtM(остаток),
    '{{ГРАФИК_ИТОГО}}': fmtM(итогоГрафик)
  };

  for (const k in map) body.replaceText(escapeRe_(k), String(map[k]));

  const tables = body.getTables();
  let tGraf = null;
  for (const t of tables) {
    if (t.getText().indexOf('{{ГРАФИК_СТРОКИ}}') !== -1) { tGraf = t; break; }
  }
  if (tGraf) {
    for (let i = 0; i < график.length; i++) {
      const r = tGraf.insertTableRow(2 + i);
      r.appendTableCell(String(i + 1));
      r.appendTableCell(fmtD(график[i].дата));
      r.appendTableCell(fmtM(график[i].сумма) + ' USD');
    }
    tGraf.removeRow(1);
  }

  doc.saveAndClose();
  ui.alert('Готово!\n\nДоговор ДДУ-' + num + ' создан в папке «ДДУ Айни».\n\n' +
           'Клиент: ' + (c[1] || '') + '\nОбъект: ' + кодОбъекта +
           '\nСумма: ' + fmtM(суммаДог) + ' USD\nПлатежей в графике: ' + график.length);
}

function escapeRe_(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}



















