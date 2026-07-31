/**
 * ЭРА Констракшн — файл «Эко Сити Айни Продажи v1.0»
 * ОДНОРАЗОВЫЙ скрипт настройки: пункты 2.2–2.3 ТЗ «ПРИЁМ ДОКУМЕНТОВ v2.2»
 * Версия 1.4 — 19.07.2026: иностранные номера — «+» достраивается по коду страны
 * (11–15 цифр = международный номер; «+» при вводе съедает сама таблица).
 * Версия 1.3 — 19.07.2026: автонормализация телефонов к формату +996/+7 при вводе
 * (функция phoneOnEdit + триггер), валидация телефона допускает пробелы и скобки.
 * Версия 1.2 — исправление 19.07.2026: надёжное определение разделителя формул.
 * Проба теперь применяется принудительно (SpreadsheetApp.flush), чтобы ошибка
 * ловилась внутри пробы, и первым пробуется «;» — родной для локали файла.
 * Замените прежний код целиком этим.
 *
 * Что делает:
 *  1) Чинит страж согласия супруга: приводит значения колонки J (семейное положение)
 *     к справочнику SP_MARITAL («женат» → «женат/замужем» и т.д.).
 *  2) Включает строгую валидацию (отклонение неверного ввода):
 *     C дата рождения, D ПИН, E № документа, G дата выдачи, I телефон,
 *     J семейное положение, L форма согласия, Q срок действия.
 *  3) Создаёт колонки P «Статус ввода» и Q «Срок действия документа».
 *  4) Условное форматирование: красный — не сверено/ошибка, жёлтый — возможный дубль,
 *     оранжевая ячейка ПИН — если позиции 2–9 ПИН не совпадают с датой рождения.
 *
 * КАК ЗАПУСТИТЬ: как раньше — вставить целиком вместо прежнего кода, сохранить,
 * выбрать функцию setupValidationClients → Выполнить. Повторный запуск безопасен:
 * то, что применилось при первом (прерванном) запуске, корректно перезапишется.
 *
 * ПОСЛЕ ЗАПУСКА ЭТО НОРМАЛЬНО:
 *  - Старые мусорные значения останутся с пометками «неверные данные» — эти строки
 *    всё равно уходят по чек-листу перезапуска (раздел 2.1 ТЗ).
 *  - У клиентов в браке без заполненного согласия колонка N покажет
 *    «БРАК: нет согласия супруга» — это страж заработал, а не сломался.
 *  - После запуска проверьте глазами, что подсветка работает: впишите в любую
 *    пустую строку в P слово РАСПОЗНАНО — строка должна стать красной; удалите после проверки.
 */

function setupValidationClients() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Клиенты');
  const sp = ss.getSheetByName('Справочники');
  if (!sh || !sp) throw new Error('Не найден лист «Клиенты» или «Справочники» — проверьте названия листов.');

  const LAST = 1001; // рабочий диапазон строк 2..1001, как в именованных диапазонах файла
  const N_ROWS = LAST - 1;
  const today = new Date();
  const summary = [];

  // ── 0. Определяем разделитель аргументов формул под локаль файла ──
  const SEP = detectFormulaSep_(sh);
  const f = (tpl) => tpl.split('%S%').join(SEP);
  summary.push('Разделитель формул: «' + SEP + '»');

  // ── 1. Нормализация существующих значений J (семейное положение) ──
  const jRange = sh.getRange(2, 10, N_ROWS, 1); // колонка J
  const jVals = jRange.getValues();
  const MAP = {
    'женат': 'женат/замужем',
    'замужем': 'женат/замужем',
    'женат/замужем': 'женат/замужем',
    'не женат': 'холост/не замужем',
    'не замужем': 'холост/не замужем',
    'холост': 'холост/не замужем',
    'холостой': 'холост/не замужем',
    'холост/не замужем': 'холост/не замужем',
    'разведен': 'разведён(а)',
    'разведён': 'разведён(а)',
    'разведена': 'разведён(а)',
    'разведён(а)': 'разведён(а)',
    'вдовец': 'вдовец/вдова',
    'вдова': 'вдовец/вдова',
    'вдовец/вдова': 'вдовец/вдова'
  };
  let fixed = 0;
  for (let i = 0; i < jVals.length; i++) {
    const raw = String(jVals[i][0] || '').trim();
    if (!raw) continue;
    const mapped = MAP[raw.toLowerCase()];
    if (mapped && mapped !== raw) { jVals[i][0] = mapped; fixed++; }
  }
  if (fixed > 0) jRange.setValues(jVals);
  summary.push('Семейное положение: нормализовано значений — ' + fixed);

  // ── 2. Заголовки новых колонок P и Q (не затирая, если уже есть) ──
  if (!sh.getRange('P1').getValue()) sh.getRange('P1').setValue('Статус ввода');
  if (!sh.getRange('Q1').getValue()) sh.getRange('Q1').setValue('Срок действия документа');
  sh.getRange('O1').copyTo(sh.getRange('P1:Q1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  summary.push('Колонки P и Q: готовы');

  // ── 3. Текстовый формат для ПИН, № документа и телефона (защита «+» и ведущих нулей) ──
  sh.getRange(2, 4, N_ROWS, 1).setNumberFormat('@'); // D ПИН
  sh.getRange(2, 5, N_ROWS, 1).setNumberFormat('@'); // E № документа
  sh.getRange(2, 9, N_ROWS, 1).setNumberFormat('@'); // I телефон

  // ── 4. Валидация (строгая: неверный ввод отклоняется) ──
  const v = (b) => b.setAllowInvalid(false).build();

  // J — семейное положение: строго из справочника SP_MARITAL
  sh.getRange(2, 10, N_ROWS, 1).setDataValidation(v(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(sp.getRange('O2:O5'), true)
      .setHelpText('Только из справочника: женат/замужем, холост/не замужем, разведён(а), вдовец/вдова')
  ));

  // L — форма согласия: строго из справочника SP_CONSENT
  sh.getRange(2, 12, N_ROWS, 1).setDataValidation(v(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(sp.getRange('Q2:Q4'), true)
      .setHelpText('Только из справочника форм согласия')
  ));

  // C — дата рождения: 01.01.1930 … (сегодня минус 18 лет)
  const maxBirth = new Date(today.getFullYear() - 18, today.getMonth(), today.getDate());
  sh.getRange(2, 3, N_ROWS, 1).setDataValidation(v(
    SpreadsheetApp.newDataValidation()
      .requireDateBetween(new Date(1930, 0, 1), maxBirth)
      .setHelpText('Дата рождения: от 1930 года, клиент не младше 18 лет')
  ));

  // G — дата выдачи: 01.01.2004 … сегодня
  sh.getRange(2, 7, N_ROWS, 1).setDataValidation(v(
    SpreadsheetApp.newDataValidation()
      .requireDateBetween(new Date(2004, 0, 1), today)
      .setHelpText('Дата выдачи: от 2004 года до сегодня')
  ));

  // Q — срок действия документа: дата после 01.01.2004 (может быть в будущем)
  sh.getRange(2, 17, N_ROWS, 1).setDataValidation(v(
    SpreadsheetApp.newDataValidation()
      .requireDateAfter(new Date(2004, 0, 1))
      .setHelpText('Срок действия документа — дата')
  ));

  // D — ПИН: пусто (иностранцы) или ровно 14 цифр
  sh.getRange(2, 4, N_ROWS, 1).setDataValidation(v(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied(f('=OR($D2=""%S%REGEXMATCH(TO_TEXT($D2)%S%"^\\d{14}$"))'))
      .setHelpText('ПИН: ровно 14 цифр; для иностранцев без ПИН — оставить пустым')
  ));

  // E — № документа: без пробелов
  sh.getRange(2, 5, N_ROWS, 1).setDataValidation(v(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied(f('=OR($E2=""%S%NOT(REGEXMATCH(TO_TEXT($E2)%S%" ")))'))
      .setHelpText('Номер документа единым словом без пробелов, например ID3039209')
  ));

  // I — телефон: цифры, допускаются «+», пробелы, скобки, дефисы —
  // автонормализация к формату +996… выполняется триггером phoneOnEdit
  sh.getRange(2, 9, N_ROWS, 1).setDataValidation(v(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied(f('=OR($I2=""%S%REGEXMATCH(TO_TEXT($I2)%S%"^\\+?[0-9()\\-\\s]{7,20}$"))'))
      .setHelpText('Телефон: цифры (можно с «+», пробелами, скобками) — система сама приведёт к виду +996…')
  ));

  // триггер автонормализации телефонов (установка идемпотентна)
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'phoneOnEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('phoneOnEdit').forSpreadsheet(ss).onEdit().create();

  summary.push('Валидация C, D, E, G, I, J, L, Q: включена (строгая)');

  // ── 5. Условное форматирование (без задвоения при повторном запуске) ──
  const T_RED = '=AND($P2<>""%S%$P2<>"СВЕРЕНО"%S%NOT(ISNUMBER(SEARCH("ДУБЛЬ"%S%$P2))))';
  const T_YEL = '=ISNUMBER(SEARCH("ДУБЛЬ"%S%$P2))';
  const T_PIN = '=AND($D2<>""%S%$C2<>""%S%MID(TO_TEXT($D2)%S%2%S%8)<>TEXT($C2%S%"ddmmyyyy"))';
  // для фильтра повторного запуска учитываем оба варианта разделителя:
  const mine = [];
  [T_RED, T_YEL, T_PIN].forEach(function (t) {
    mine.push(t.split('%S%').join(','));
    mine.push(t.split('%S%').join(';'));
  });

  const kept = sh.getConditionalFormatRules().filter(function (r) {
    const bc = r.getBooleanCondition();
    if (!bc) return true;
    const cv = bc.getCriteriaValues();
    return !(cv && cv.length && mine.indexOf(String(cv[0])) !== -1);
  });

  const fullRange = sh.getRange('A2:Q' + LAST);
  kept.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(f(T_YEL)).setBackground('#fff2cc') // жёлтый: возможный дубль
      .setRanges([fullRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(f(T_RED)).setBackground('#f4cccc') // красный: не сверено / ошибка
      .setRanges([fullRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(f(T_PIN)).setBackground('#f9cb9c') // оранжевый: ПИН не совпал с датой рождения
      .setRanges([sh.getRange('D2:D' + LAST)]).build()
  );
  sh.setConditionalFormatRules(kept);
  summary.push('Подсветка: не сверено (красный), дубль (жёлтый), ПИН≠дата рождения (оранжевый)');

  // ── Итог ──
  const msg = summary.join(' • ');
  Logger.log(msg);
  ss.toast(msg, 'Настройка листа «Клиенты» завершена', 10);
}

/**
 * Определяет разделитель аргументов формул («;» или «,») пробными правилами
 * валидации на служебной ячейке Z1. Первым пробуется «;» (локаль файла).
 * SpreadsheetApp.flush() принудительно применяет операцию сразу — без этого
 * Apps Script ставит её в очередь, и ошибка всплывает позже, мимо try/catch.
 */
function detectFormulaSep_(sh) {
  const probe = sh.getRange('Z1');
  const saved = probe.getDataValidation();

  function works(formula) {
    try {
      probe.setDataValidation(
        SpreadsheetApp.newDataValidation().requireFormulaSatisfied(formula).build()
      );
      SpreadsheetApp.flush();
      return true;
    } catch (e) {
      return false;
    } finally {
      probe.setDataValidation(null);
    }
  }

  let sep = null;
  if (works('=OR(TRUE;TRUE)')) {
    sep = ';';
  } else if (works('=OR(TRUE,TRUE)')) {
    sep = ',';
  }

  probe.setDataValidation(saved);
  SpreadsheetApp.flush();

  if (!sep) {
    throw new Error('Не удалось определить разделитель формул: обе пробы отклонены. Пришлите этот текст Клоду.');
  }
  return sep;
}

/**
 * Автонормализация телефонов в колонке I листа «Клиенты» при вводе.
 * «0555 25 11 00» → «+996555251100»; 9 цифр → «+996…»; «8…» (11 цифр) → «+7…»;
 * «996…», «7…» → добавляется «+». Непонятные форматы не трогаются.
 * Триггер устанавливается функцией setupValidationClients.
 */
function phoneOnEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== 'Клиенты') return;
    if (e.range.getColumn() !== 9 || e.range.getRow() < 2) return;
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return; // только одиночный ввод
    const raw = String(e.range.getValue() || '').trim();
    if (!raw) return;
    const norm = normalizePhone_(raw);
    if (norm && norm !== raw) e.range.setValue(norm);
  } catch (err) { /* не мешаем вводу */ }
}

function normalizePhone_(raw) {
  const hasPlus = raw.charAt(0) === '+';
  const digits = raw.replace(/\D/g, '');
  if (!digits) return null;
  if (hasPlus) return '+' + digits;
  if (digits.indexOf('996') === 0 && digits.length === 12) return '+' + digits;
  if (digits.charAt(0) === '0' && digits.length === 10) return '+996' + digits.substring(1);
  if (digits.length === 9) return '+996' + digits;
  if (digits.charAt(0) === '8' && digits.length === 11) return '+7' + digits.substring(1);
  if (digits.charAt(0) === '7' && digits.length === 11) return '+' + digits;
  if (digits.charAt(0) === '1' && digits.length === 11) return '+' + digits; // США/Канада
  // Общий случай: 11–15 цифр — международный номер, введённый без «+»
  // (или «+» съела сама таблица при вводе — она трактует его как начало формулы)
  if (digits.length >= 11 && digits.length <= 15) return '+' + digits;
  return null; // формат не распознан — оставляем как ввели
}
