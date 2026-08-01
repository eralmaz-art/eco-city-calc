/**
 * Тесты модуля математики графика (docs/calc.js) по кейсам ТЗ ред. 1.9.
 *
 * Запуск:
 *   Node:    node tests/calc.test.js   (код возврата 0 — все тесты прошли)
 *   Браузер: открыть tests/run.html
 */
(function (root) {
'use strict';

var EcoCalc = (typeof module === 'object' && module.exports && typeof require === 'function')
  ? require('../docs/calc.js')
  : root.EcoCalc;

var results = [];

function test(name, fn) {
  try { fn(); results.push({ name: name, ok: true }); }
  catch (e) { results.push({ name: name, ok: false, error: String(e && e.message || e) }); }
}
function ok(v, msg) { if (!v) throw new Error(msg || 'условие не выполнено'); }
function eq(got, want, msg) {
  if (got !== want) throw new Error((msg || 'значение') + ': ожидалось ' + JSON.stringify(want) + ', получено ' + JSON.stringify(got));
}
function hasWarning(res, code) {
  return res.warnings.some(function (w) { return w.code === code; });
}
function rowsOf(res, type) {
  return res.rows.filter(function (r) { return r.type === type; });
}
function T() { return EcoCalc.TYPES; }

/* ============ 1. Кейс-образец владельца (ТЗ 7): $1 000 при подписании
   + $20 000 через 15 дней + серия с удобного числа ============ */
test('Кейс ТЗ 7: 1000 сегодня + 20000 через 15 дней + серия с 5-го числа', function () {
  var res = EcoCalc.build({
    total: 48360,
    contractDate: '2026-08-01',
    startPayments: [
      { date: '2026-08-01', sum: 1000 },                       // при подписании
      { date: '2026-08-16', sum: 20000, type: T().advance }    // основной взнос через 15 дней
    ],
    anchorDate: '2026-09-05',                                  // удобное число дольщика
    mode: 'payment',
    monthlyPayment: 1500
  });
  ok(res.ok, 'расчёт должен быть корректен: ' + JSON.stringify(res.warnings));
  eq(res.rows[0].date, '2026-08-01', 'первая строка — платёж при подписании');
  eq(res.rows[0].sum, 1000);
  eq(res.rows[0].type, T().down);
  eq(res.rows[1].date, '2026-08-16', 'вторая строка — основной взнос');
  eq(res.rows[1].sum, 20000);
  eq(res.rows[1].type, T().advance);

  var monthly = rowsOf(res, T().monthly);
  eq(monthly.length, 18, 'число ежемесячных платежей');          // 27360 = 18×1500 + 360
  eq(monthly[0].date, '2026-09-05', 'серия начинается с якоря');
  monthly.forEach(function (r) {
    eq(r.sum, 1500, 'ежемесячный платёж');
    eq(r.date.slice(8), '05', 'все ежемесячные — 5-го числа');
  });

  var fin = rowsOf(res, T().final);
  eq(fin.length, 1, 'один финальный платёж');
  eq(fin[0].sum, 360, 'хвост');
  eq(fin[0].date, '2028-03-05', 'финальный — 19-й месяц серии');
  eq(res.months, 19, 'срок = 19 (18 полных + финальный)');
  ok(res.checksum.ok, 'контрольная сумма сходится до доллара');
});

/* ============ 2. «Выровнять взносом» — пример из ТЗ §3 и обе стороны ============ */
test('Выровнять взносом: подсказки как в примере ТЗ §3 (48 360 / 20 000 / 1 500)', function () {
  var res = EcoCalc.build({
    total: 48360,
    startPayments: [{ date: '2026-08-01', sum: 20000 }],
    anchorDate: '2026-09-01',
    mode: 'payment',
    monthlyPayment: 1500
  });
  ok(res.levelSuggestions, 'хвост есть — подсказки должны быть');
  eq(res.levelSuggestions.tail, 1360, 'хвост');
  eq(res.levelSuggestions.up.startSum, 21360, 'взнос вверх: ПВ + хвост');
  eq(res.levelSuggestions.up.months, 18, 'вверх → ровно 18 × 1500');
  eq(res.levelSuggestions.down.startSum, 19860, 'взнос вниз: ПВ − (M − хвост)');
  eq(res.levelSuggestions.down.months, 19, 'вниз → ровно 19 × 1500');
});

test('Выровнять взносом вверх: после применения график идеально круглый', function () {
  var res = EcoCalc.build({
    total: 48360,
    startPayments: [{ date: '2026-08-01', sum: 21360 }],   // применили подсказку «вверх»
    anchorDate: '2026-09-01',
    mode: 'payment',
    monthlyPayment: 1500
  });
  ok(res.ok);
  eq(res.finalPayment, 0, 'финального платежа нет');
  eq(rowsOf(res, T().final).length, 0);
  eq(rowsOf(res, T().monthly).length, 18, 'ровно 18 × 1500');
  eq(res.levelSuggestions, null, 'делится ровно — подсказок нет');
  ok(res.checksum.ok);
});

test('Выровнять взносом вниз: после применения график идеально круглый', function () {
  var res = EcoCalc.build({
    total: 48360,
    startPayments: [{ date: '2026-08-01', sum: 19860 }],   // применили подсказку «вниз»
    anchorDate: '2026-09-01',
    mode: 'payment',
    monthlyPayment: 1500
  });
  ok(res.ok);
  eq(res.finalPayment, 0);
  eq(rowsOf(res, T().monthly).length, 19, 'ровно 19 × 1500');
  eq(res.levelSuggestions, null);
  ok(res.checksum.ok);
});

test('Выровнять взносом при договорённых: формула учитывает Σ договорённых (ТЗ §3)', function () {
  var res = EcoCalc.build({
    total: 48360,
    startPayments: [{ date: '2026-08-01', sum: 20000 }],
    agreedPayments: [{ date: '2026-11-20', sum: 5000 }],
    anchorDate: '2026-09-05',
    mode: 'payment',
    monthlyPayment: 1500
  });
  // R = 48360 − 20000 − 5000 = 23360; 23360 % 1500 = 860
  ok(res.levelSuggestions);
  eq(res.levelSuggestions.tail, 860);
  eq(res.levelSuggestions.up.startSum, 20860);
  eq(res.levelSuggestions.down.startSum, 20000 - (1500 - 860));
  // применяем «вверх» — делится нацело
  var res2 = EcoCalc.build({
    total: 48360,
    startPayments: [{ date: '2026-08-01', sum: 20860 }],
    agreedPayments: [{ date: '2026-11-20', sum: 5000 }],
    anchorDate: '2026-09-05',
    mode: 'payment',
    monthlyPayment: 1500
  });
  eq(res2.finalPayment, 0, 'после выравнивания хвоста нет');
  ok(res2.checksum.ok);
});

/* ============ 3. Гарантия по построению: хвост всегда меньше M ============ */
test('Гарантия «хвост < M» на сетке параметров (режим «От платежа»)', function () {
  var checked = 0;
  for (var total = 30000; total <= 48360; total += 3731) {
    [0, 5000, 20000].forEach(function (down) {
      [700, 1500, 2100].forEach(function (M) {
        var res = EcoCalc.build({
          total: total,
          startPayments: down > 0 ? [{ date: '2026-08-01', sum: down }] : [],
          anchorDate: '2026-09-05',
          mode: 'payment',
          monthlyPayment: M,
          maxMonths: 120
        });
        if (res.finalPayment > 0) {
          ok(res.finalPayment < M, 'хвост ' + res.finalPayment + ' ≥ M ' + M +
            ' при total=' + total + ', down=' + down);
        }
        ok(res.checksum.ok, 'checksum при total=' + total + ', down=' + down + ', M=' + M);
        checked++;
      });
    });
  }
  ok(checked > 40, 'сетка прогнана: ' + checked + ' комбинаций');
});

/* ============ 4. Якорь 29–31 числа при переходе через февраль ============ */
test('Якорь 31-е: февраль поджимается, март возвращается на 31-е (ТЗ 7.2)', function () {
  var res = EcoCalc.build({
    total: 10000,
    anchorDate: '2026-12-31',
    mode: 'payment',
    monthlyPayment: 1500
  });
  ok(res.ok, JSON.stringify(res.warnings));
  var dates = res.rows.map(function (r) { return r.date; });
  eq(JSON.stringify(dates), JSON.stringify([
    '2026-12-31', '2027-01-31', '2027-02-28', '2027-03-31',
    '2027-04-30', '2027-05-31', '2027-06-30'
  ]), 'серия с якорем 31-го');
  eq(res.rows[2].sum, 1500, 'февральский платёж обычный');
  eq(res.finalPayment, 1000, 'финальный 1000 = 10000 − 6×1500');
  ok(res.checksum.ok);
});

test('slotDate: високосный февраль (31-е → 29.02.2028) и 29-е число', function () {
  var d = EcoCalc.slotDate(new Date(2027, 11, 31), 2);   // дек 2027 + 2 = фев 2028
  eq(d.getFullYear(), 2028); eq(d.getMonth(), 1); eq(d.getDate(), 29, '2028 — високосный');
  var d2 = EcoCalc.slotDate(new Date(2026, 0, 29), 1);   // янв 29 + 1 мес = фев 2026
  eq(d2.getDate(), 28, '29-е в невисокосном феврале → 28-е');
  var d3 = EcoCalc.slotDate(new Date(2026, 0, 29), 3);   // апрель
  eq(d3.getDate(), 29, 'после февраля серия возвращается на своё число');
});

/* ============ 5. Контрольная сумма сходится до доллара ============ */
test('Контрольная сумма при некруглом делении («От срока», равными долями)', function () {
  var res = EcoCalc.build({
    total: 50000,
    startPayments: [{ date: '2026-08-20', sum: 7777 }],
    anchorDate: '2026-09-10',
    mode: 'term',
    months: 12,
    restMode: 'equal'
  });
  ok(res.ok, JSON.stringify(res.warnings));
  eq(res.monthly, 3519, 'равный платёж: round(42223 / 12)');
  var monthly = rowsOf(res, T().monthly);
  eq(monthly.length, 12);
  eq(monthly[11].sum, 42223 - 3519 * 11, 'последний равный сводит до доллара');
  eq(res.checksum.scheduled, 50000, 'сумма всех строк = стоимость');
  ok(res.checksum.ok);
});

/* ============ Режим «От срока», хвостом в конце + правило 10% ============ */
test('«От срока» хвостом в конце: договорённый учтён, правило 10% срабатывает', function () {
  var res = EcoCalc.build({
    total: 48000,
    startPayments: [{ date: '2026-08-10', sum: 20000 }],
    agreedPayments: [{ date: '2026-11-20', sum: 5000 }],
    anchorDate: '2026-09-05',
    mode: 'term',
    months: 12,
    restMode: 'tail'
  });
  eq(res.finalPayment, 23000, 'хвост = 48000 − 20000 − 5000');
  var fin = rowsOf(res, T().final);
  eq(fin[0].date, '2027-08-05', 'хвост в последнем месяце срока');
  ok(hasWarning(res, 'final-over-10pct'), 'финальный 23000 > 10% от 48000 — красное предупреждение');
  ok(res.checksum.ok, 'контрольная сумма сходится');
});

/* ============ «Влить в последний» (ТЗ §3.2) ============ */
test('«Влить в последний»: 17 × 1500, финальный 18-й — 2860 (пример ТЗ §3)', function () {
  var res = EcoCalc.build({
    total: 48360,
    startPayments: [{ date: '2026-08-01', sum: 20000 }],
    anchorDate: '2026-09-05',
    mode: 'payment',
    monthlyPayment: 1500,
    tailMode: 'merge'
  });
  ok(res.ok, JSON.stringify(res.warnings));
  eq(rowsOf(res, T().monthly).length, 17, '17 полных платежей');
  var fin = rowsOf(res, T().final);
  eq(fin.length, 1);
  eq(fin[0].sum, 2860, 'финальный = 1500 + 1360');
  eq(fin[0].date, '2028-02-05', 'финальный — в 18-м месяце серии');
  eq(res.months, 18, 'срок 18 месяцев');
  eq(res.finalPayment, 2860);
  ok(res.checksum.ok);
});

test('«Хвост отдельным платежом» (по умолчанию): 18 × 1500 + 19-й 1360 (пример ТЗ §3)', function () {
  var res = EcoCalc.build({
    total: 48360,
    startPayments: [{ date: '2026-08-01', sum: 20000 }],
    anchorDate: '2026-09-05',
    mode: 'payment',
    monthlyPayment: 1500
  });
  eq(rowsOf(res, T().monthly).length, 18);
  eq(res.finalPayment, 1360);
  eq(res.months, 19);
  ok(res.checksum.ok);
});

/* ============ Договорённые платежи внутри серии («От платежа») ============ */
test('«От платежа» с договорённым: месяц договорённого — без ежемесячного (ТЗ §2)', function () {
  var res = EcoCalc.build({
    total: 48360,
    startPayments: [{ date: '2026-08-10', sum: 20000 }],
    agreedPayments: [{ date: '2026-11-20', sum: 5000 }],
    anchorDate: '2026-09-05',
    mode: 'payment',
    monthlyPayment: 1500
  });
  ok(res.ok, JSON.stringify(res.warnings));
  var novPayments = res.rows.filter(function (r) { return r.date.slice(0, 7) === '2026-11'; });
  eq(novPayments.length, 1, 'в ноябре только договорённый');
  eq(novPayments[0].type, T().agreed);
  eq(novPayments[0].date, '2026-11-20', 'договорённый — своей реальной датой');
  eq(rowsOf(res, T().monthly).length, 15, '2 до ноября + 13 после');
  eq(res.finalPayment, 860, 'хвост: 48360−20000−5000−15×1500');
  eq(rowsOf(res, T().final)[0].date, '2028-01-05');
  ok(res.checksum.ok);
});

/* ============ Защиты (ТЗ §4) ============ */
test('Защита: M ≥ остатка → подсказка «разовый платёж», график из одного финального', function () {
  var res = EcoCalc.build({
    total: 10000,
    startPayments: [{ date: '2026-08-01', sum: 5000 }],
    anchorDate: '2026-09-01',
    mode: 'payment',
    monthlyPayment: 6000
  });
  ok(hasWarning(res, 'single-payment'));
  eq(res.finalPayment, 5000, 'весь остаток одним платежом');
  eq(res.months, 1);
  ok(hasWarning(res, 'final-over-10pct'), '5000 > 10% от 10000');
  ok(res.checksum.ok);
});

test('Защита: M ≤ 0 → ошибка', function () {
  var res = EcoCalc.build({
    total: 10000, anchorDate: '2026-09-01', mode: 'payment', monthlyPayment: 0
  });
  ok(!res.ok);
  ok(hasWarning(res, 'invalid-monthly'));
});

test('Защита: срок больше максимального → красное + минимальный платёж', function () {
  var res = EcoCalc.build({
    total: 100000, anchorDate: '2026-09-01', mode: 'payment', monthlyPayment: 1000
  });
  var w = res.warnings.filter(function (x) { return x.code === 'over-max-term'; })[0];
  ok(w, 'предупреждение о превышении срока');
  eq(w.level, 'red');
  ok(w.message.indexOf('2778') > -1, 'подсказка M_min = ceil(100000/36) = 2778');
  ok(res.checksum.ok, 'график всё равно построен и сходится');
});

test('Защита: стартовые больше стоимости → ошибка (ТЗ 7.5)', function () {
  var res = EcoCalc.build({
    total: 10000,
    startPayments: [{ date: '2026-08-01', sum: 12000 }],
    anchorDate: '2026-09-01',
    mode: 'payment', monthlyPayment: 1500
  });
  ok(!res.ok);
  ok(hasWarning(res, 'start-exceeds-total'));
});

test('Правило 10% в режиме «От платежа»: хвост 2000 при стоимости 10000', function () {
  var res = EcoCalc.build({
    total: 10000, anchorDate: '2026-09-01', mode: 'payment', monthlyPayment: 8000
  });
  eq(res.finalPayment, 2000);
  ok(hasWarning(res, 'final-over-10pct'));
  ok(res.checksum.ok);
});

/* ============ Валидации дат (ТЗ 7.5) ============ */
test('Якорь раньше последнего стартового → предупреждение', function () {
  var res = EcoCalc.build({
    total: 20000,
    startPayments: [{ date: '2026-09-10', sum: 5000 }],
    anchorDate: '2026-09-05',
    mode: 'payment', monthlyPayment: 1500
  });
  ok(hasWarning(res, 'anchor-before-start'));
});

test('Нет стартового на дату договора → жёлтая пометка «квартира без денег»', function () {
  var res = EcoCalc.build({
    total: 48360,
    contractDate: '2026-08-01',
    startPayments: [{ date: '2026-08-16', sum: 20000 }],
    anchorDate: '2026-09-05',
    mode: 'payment', monthlyPayment: 1500
  });
  ok(hasWarning(res, 'no-payment-on-contract-date'));
});

/* ============ Словарь типов (ТЗ 7.7, 8.4) ============ */
test('Словарь стартовых типов: без «Задатка», «Зачёт брони» — только системный', function () {
  var manual = EcoCalc.START_PAYMENT_TYPES;
  eq(manual.length, 2);
  eq(manual[0], 'Первоначальный взнос');
  eq(manual[1], 'Аванс');
  ok(manual.indexOf('Задаток') === -1, '«Задаток» исключён решением владельца');
  ok(manual.indexOf('Зачёт брони') === -1, '«Зачёт брони» руками не выбирается');
  eq(EcoCalc.TYPES.booking, 'Зачёт брони', 'но в словаре типов существует — для автоподстановки');
});

/* ---------- отчёт ---------- */
var passed = results.filter(function (r) { return r.ok; }).length;
var failed = results.length - passed;

if (typeof module === 'object' && module.exports && typeof process !== 'undefined') {
  results.forEach(function (r) {
    console.log((r.ok ? '  ✓ ' : '  ✗ ') + r.name + (r.ok ? '' : '\n      ' + r.error));
  });
  console.log('\n' + (failed === 0 ? 'ВСЕ ТЕСТЫ ПРОШЛИ' : 'ЕСТЬ ПАДЕНИЯ') +
    ': ' + passed + ' из ' + results.length);
  process.exitCode = failed === 0 ? 0 : 1;
} else {
  root.CALC_TEST_RESULTS = { results: results, passed: passed, failed: failed };
}

})(typeof self !== 'undefined' ? self : this);
