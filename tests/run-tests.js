/* 判定ロジックのテスト。実行: node tests/run-tests.js
 * 検証ケースの ID は docs/test-plan.md に対応する。 */
var assert = require('assert');
var L = require('../src/logic.js');

var results = [];
function test(id, name, fn) {
  try { fn(); results.push({ id: id, ok: true }); }
  catch (e) { results.push({ id: id, ok: false, name: name, msg: e.message }); }
}
function judge(billed, payments, due, today, y, r) {
  return L.judgePayment({ billed: billed, payments: payments, dueDate: due, today: today, yellowDays: y, redDays: r });
}

// A. 入金の状態（R-01〜R-03）
test('T-01', 'R-01 満額入金', function () {
  var j = judge(80000, [80000], '2026-11-27', '2026-11-20');
  assert.strictEqual(j.status, 'PAID'); assert.strictEqual(j.level, 'NONE');
});
test('T-02', 'R-01 期日後の満額入金', function () {
  var j = judge(80000, [80000], '2026-11-27', '2026-12-05');
  assert.strictEqual(j.status, 'PAID'); assert.strictEqual(j.level, 'NONE'); assert.strictEqual(j.overdueDays, 0);
});
test('T-03', 'R-01 複数行の合計', function () {
  assert.strictEqual(judge(80000, [40000, 40000], '2026-11-27', '2026-11-27').status, 'PAID');
});
test('T-04', 'R-02 一部入金', function () {
  var j = judge(80000, [50000], '2026-11-27', '2026-11-20');
  assert.strictEqual(j.status, 'PARTIAL'); assert.strictEqual(j.shortage, 30000); assert.strictEqual(j.level, 'NONE');
});
test('T-05', 'R-02 1円だけ入金', function () {
  var j = judge(80000, [1], '2026-11-27', '2026-11-20');
  assert.strictEqual(j.status, 'PARTIAL'); assert.strictEqual(j.shortage, 79999);
});
test('T-06', 'R-03 過入金', function () {
  var j = judge(80000, [85000], '2026-11-27', '2026-11-27');
  assert.strictEqual(j.status, 'OVERPAID'); assert.strictEqual(j.excess, 5000); assert.strictEqual(j.level, 'NONE');
});

// B. 期日前・滞納の段階（R-04〜R-06）
test('T-07', 'R-04 期日前', function () {
  var j = judge(80000, [], '2026-11-27', '2026-11-20');
  assert.strictEqual(j.status, 'UNPAID'); assert.strictEqual(j.level, 'NONE');
});
test('T-08', 'R-04 期日当日', function () {
  var j = judge(80000, [], '2026-11-27', '2026-11-27');
  assert.strictEqual(j.status, 'UNPAID'); assert.strictEqual(j.overdueDays, 0); assert.strictEqual(j.level, 'NONE');
});
test('T-09', 'R-05 超過1日は黄', function () {
  var j = judge(80000, [], '2026-11-27', '2026-11-28');
  assert.strictEqual(j.overdueDays, 1); assert.strictEqual(j.level, 'YELLOW');
});
test('T-10', 'R-05 超過7日は黄', function () {
  assert.strictEqual(judge(80000, [], '2026-11-27', '2026-12-04').level, 'YELLOW');
});
test('T-11', 'R-05 超過8日は赤', function () {
  assert.strictEqual(judge(80000, [], '2026-11-27', '2026-12-05').level, 'RED');
});
test('T-12', 'R-05 一部入金で超過10日は赤', function () {
  var j = judge(80000, [30000], '2026-11-27', '2026-12-07');
  assert.strictEqual(j.status, 'PARTIAL'); assert.strictEqual(j.level, 'RED'); assert.strictEqual(j.shortage, 50000);
});
test('T-13', 'R-05 一部入金で超過3日は黄', function () {
  var j = judge(80000, [30000], '2026-11-27', '2026-11-30');
  assert.strictEqual(j.status, 'PARTIAL'); assert.strictEqual(j.level, 'YELLOW');
});
test('T-14', 'R-06 黄3・赤10: 超過2日', function () {
  assert.strictEqual(judge(80000, [], '2026-11-27', '2026-11-29', 3, 10).level, 'NONE');
});
test('T-15', 'R-06 黄3・赤10: 超過3日', function () {
  assert.strictEqual(judge(80000, [], '2026-11-27', '2026-11-30', 3, 10).level, 'YELLOW');
});
test('T-16', 'R-06 黄3・赤10: 超過10日', function () {
  assert.strictEqual(judge(80000, [], '2026-11-27', '2026-12-07', 3, 10).level, 'YELLOW');
});
test('T-17', 'R-06 黄3・赤10: 超過11日', function () {
  assert.strictEqual(judge(80000, [], '2026-11-27', '2026-12-08', 3, 10).level, 'RED');
});

// C. 期日の計算（R-07、R-08）
test('T-18', 'R-07 31日の2月', function () { assert.strictEqual(L.dueDateFor(2027, 2, 31), '2027-02-28'); });
test('T-19', 'R-07 31日のうるう年の2月', function () { assert.strictEqual(L.dueDateFor(2028, 2, 31), '2028-02-29'); });
test('T-20', 'R-07 31日の4月', function () { assert.strictEqual(L.dueDateFor(2026, 4, 31), '2026-04-30'); });
test('T-21', 'R-07 月末', function () { assert.strictEqual(L.dueDateFor(2026, 12, '月末'), '2026-12-31'); });
test('T-22', 'R-07 27日', function () { assert.strictEqual(L.dueDateFor(2026, 11, 27), '2026-11-27'); });
test('T-23', 'R-08 土曜でも繰り下げない', function () {
  assert.strictEqual(L.dueDateFor(2026, 11, 28), '2026-11-28');
  assert.strictEqual(new Date(L.parseDate('2026-11-28')).getUTCDay(), 6); // 土曜日
});

// D. 保証会社（R-09、R-10）
test('T-24', 'R-09 割り当て一致', function () {
  var c = L.checkAllocation(240000, [80000, 80000, 80000]);
  assert.strictEqual(c.ok, true); assert.strictEqual(c.diff, 0);
});
test('T-25', 'R-09 割り当て不足', function () {
  var c = L.checkAllocation(240000, [80000, 80000, 70000]);
  assert.strictEqual(c.ok, false); assert.strictEqual(c.diff, 10000);
});
test('T-26', 'R-09 割り当て超過', function () {
  var c = L.checkAllocation(240000, [80000, 80000, 90000]);
  assert.strictEqual(c.ok, false); assert.strictEqual(c.diff, -10000);
});
function ar(today, type, level) {
  return L.accidentReport({ guaranteeType: type || '一般保証型', level: level || 'RED', dueDate: '2026-11-27', days: 14, today: today });
}
test('T-27', 'R-10 残り6日', function () {
  var r = ar('2026-12-05'); assert.strictEqual(r.deadline, '2026-12-11'); assert.strictEqual(r.remainingDays, 6); assert.strictEqual(r.overdue, false);
});
test('T-28', 'R-10 報告期限超過', function () {
  var r = ar('2026-12-12'); assert.strictEqual(r.overdue, true);
});
test('T-29', 'R-10 期限当日は超過ではない', function () {
  var r = ar('2026-12-11'); assert.strictEqual(r.remainingDays, 0); assert.strictEqual(r.overdue, false);
});
test('T-30', 'R-10 保証なし・支払委託型は出さない', function () {
  assert.strictEqual(ar('2026-12-05', '保証なし'), null);
  assert.strictEqual(ar('2026-12-05', '支払委託型'), null);
});

// E. 期限管理（R-11）
test('T-31', 'R-11 91日は通常', function () { assert.strictEqual(L.deadlineLevel(91).level, 'NONE'); });
test('T-32', 'R-11 90日は黄', function () { assert.strictEqual(L.deadlineLevel(90).level, 'YELLOW'); });
test('T-33', 'R-11 31日は黄', function () { assert.strictEqual(L.deadlineLevel(31).level, 'YELLOW'); });
test('T-34', 'R-11 30日は赤', function () { assert.strictEqual(L.deadlineLevel(30).level, 'RED'); });
test('T-35', 'R-11 0日は赤', function () {
  var d = L.deadlineLevel(0); assert.strictEqual(d.level, 'RED'); assert.strictEqual(d.expired, false);
});
test('T-36', 'R-11 期限超過は赤', function () {
  var d = L.deadlineLevel(-5); assert.strictEqual(d.level, 'RED'); assert.strictEqual(d.expired, true);
});

// F. 年間収支と翌月の作成（R-12、R-13）
test('T-37', 'R-12 入金日ベース', function () {
  var s = L.summarizeIncome([{ billed: 80000, payments: [{ date: '2026-12-03', amount: 80000 }] }], 2026);
  assert.strictEqual(s.income[11], 80000); assert.strictEqual(s.income[10], 0);
});
test('T-38', 'R-12 過入金は預り', function () {
  var s = L.summarizeIncome([{ billed: 80000, payments: [{ date: '2026-11-25', amount: 85000 }] }], 2026);
  assert.strictEqual(s.income[10], 80000); assert.strictEqual(s.deposit[10], 5000);
});
test('T-39', 'R-12 年間の収入・経費・収支', function () {
  var recs = [];
  for (var m = 1; m <= 12; m++) {
    var d = '2026-' + ('0' + m).slice(-2) + '-25';
    recs.push({ billed: 80000, payments: [{ date: d, amount: 80000 }] });
  }
  var s = L.summarizeIncome(recs, 2026);
  var exp = [0, 0, 30000, 0, 0, 0, 0, 0, 0, 0, 0, 120000];
  var a = L.annualSummary(s.income, exp);
  assert.strictEqual(a.income, 960000); assert.strictEqual(a.expense, 150000); assert.strictEqual(a.net, 810000);
});
test('T-40', 'R-13 入居中だけ作る', function () {
  var ledger = [
    { room: '101', tenant: 'A', rent: 70000, fee: 5000, dueDay: 27, status: '入居中' },
    { room: '102', tenant: 'B', rent: 65000, fee: 5000, dueDay: 27, status: '入居中' },
    { room: '103', tenant: 'C', rent: 60000, fee: 0, dueDay: '月末', status: '入居中' },
    { room: '104', tenant: 'D', rent: 60000, fee: 0, dueDay: 27, status: '退去済' }
  ];
  var rows = L.nextMonthRows(ledger, 2026, 12);
  assert.strictEqual(rows.length, 3);
  assert.strictEqual(rows[0].billed, 75000);
  assert.strictEqual(rows[2].dueDate, '2026-12-31');
});
test('T-41', 'R-13 請求額の上書き', function () {
  var ledger = [{ room: '101', tenant: 'A', rent: 70000, fee: 10000, dueDay: 27, status: '入居中' }];
  var rows = L.nextMonthRows(ledger, 2026, 12, { '101': 40000 });
  assert.strictEqual(rows[0].billed, 40000);
});

// G. 入力の不正
test('T-42', '請求額が0以下', function () {
  assert.throws(function () { judge(0, [], '2026-11-27', '2026-11-28'); });
  assert.throws(function () { judge(-1, [], '2026-11-27', '2026-11-28'); });
});
test('T-43', '入金額が負', function () {
  assert.throws(function () { judge(80000, [-100], '2026-11-27', '2026-11-28'); });
});
test('T-44', '日付が不正', function () {
  assert.throws(function () { judge(80000, [], '2026-13-40', '2026-11-28'); });
  assert.throws(function () { judge(80000, [], '2026-11-27', 'abc'); });
});

// ===== 追加機能（R-14〜R-22） =====
// I. 設備の履歴（R-14）
test('T-52', 'R-14 月末の月加算', function () { assert.strictEqual(L.addMonths('2026-01-31', 1), '2026-02-28'); });
test('T-53', 'R-14 うるう日の12か月後', function () { assert.strictEqual(L.addMonths('2024-02-29', 12), '2025-02-28'); });
test('T-54', 'R-14 次回の目安日', function () {
  var e = L.equipmentStatus({ date: '2024-08-20', years: 10, today: '2026-11-23' });
  assert.strictEqual(e.nextDate, '2034-08-20'); assert.strictEqual(e.level, 'NONE');
});
test('T-55', 'R-14 目安を過ぎている', function () {
  var e = L.equipmentStatus({ date: '2016-06-10', years: 10, today: '2026-11-23' });
  assert.strictEqual(e.level, 'EXPIRED'); assert.strictEqual(e.remainingDays < 0, true);
});
test('T-56', 'R-14 より新しい記録があれば更新済', function () {
  var e = L.equipmentStatus({ date: '2016-06-10', years: 10, today: '2026-11-23', laterExists: true });
  assert.strictEqual(e.superseded, true); assert.strictEqual(e.level, 'NONE');
});
function eqLevel(remaining) {
  var next = '2030-01-01';
  var today = L.addDays(next, -remaining);
  return L.equipmentStatus({ date: '2020-01-01', years: 10, today: today }).level;
}
test('T-57', 'R-14 段階の境界', function () {
  assert.strictEqual(eqLevel(181), 'NONE'); assert.strictEqual(eqLevel(180), 'YELLOW');
  assert.strictEqual(eqLevel(61), 'YELLOW'); assert.strictEqual(eqLevel(60), 'RED');
  assert.strictEqual(eqLevel(0), 'RED'); assert.strictEqual(eqLevel(-1), 'EXPIRED');
});
test('T-58', 'R-14 目安の年数がなければ判定しない', function () {
  var e = L.equipmentStatus({ date: '2026-05-10', years: null, today: '2026-11-23' });
  assert.strictEqual(e.nextDate, null); assert.strictEqual(e.level, 'NONE');
});

// J. 空室（R-15）
function vac(days, rent) {
  var today = '2026-11-23';
  return L.vacancyStatus({ startDate: L.addDays(today, -days), rent: rent, today: today });
}
test('T-59', 'R-15 空室日数の段階', function () {
  assert.strictEqual(vac(29).level, 'NONE'); assert.strictEqual(vac(30).level, 'YELLOW');
  assert.strictEqual(vac(59).level, 'YELLOW'); assert.strictEqual(vac(60).level, 'RED');
});
test('T-60', 'R-15 空室損失の概算', function () {
  assert.strictEqual(vac(45, 68000).loss, 102000); assert.strictEqual(vac(10, 68000).loss, 22667);
});
test('T-61', 'R-15 開始日が基準日より後なら0日', function () {
  var v = L.vacancyStatus({ startDate: '2026-12-01', rent: 68000, today: '2026-11-23' });
  assert.strictEqual(v.days, 0); assert.strictEqual(v.level, 'NONE'); assert.strictEqual(v.loss, 0);
});

// K. 敷金・保証金（R-16）
function dep(o) { var b = { deposit: 70000, repair: 30000, other: 0, moveOut: '2026-10-31', today: '2026-11-30', targetDays: 30, status: '預り中' }; for (var k in o) b[k] = o[k]; return L.depositSettlement(b); }
test('T-62', 'R-16 返還額', function () {
  var d = dep({}); assert.strictEqual(d.deducted, 30000); assert.strictEqual(d.refund, 40000); assert.strictEqual(d.extra, 0);
});
test('T-63', 'R-16 追加請求額', function () {
  var d = dep({ repair: 60000, other: 20000 }); assert.strictEqual(d.refund, 0); assert.strictEqual(d.extra, 10000);
});
test('T-64', 'R-16 退去日がなければ預り中', function () { assert.strictEqual(dep({ moveOut: null }).state, '預り中'); });
test('T-65', 'R-16 精算の目安日の当日は期間内', function () {
  var d = dep({}); assert.strictEqual(d.settleBy, '2026-11-30'); assert.strictEqual(d.state, '精算期間内');
});
test('T-66', 'R-16 目安日の翌日は精算遅れ', function () { assert.strictEqual(dep({ today: '2026-12-01' }).state, '精算遅れ'); });
test('T-67', 'R-16 精算済は遅れにしない', function () { assert.strictEqual(dep({ today: '2027-03-01', status: '精算済' }).state, '精算済'); });

// L. 家賃改定（R-17）
test('T-68', 'R-17 増減額・改定率・年間', function () {
  var r = L.rentRevision({ current: 70000, market: 74000, proposed: 72000, churn: 0.05, vacancyDays: 60 });
  assert.strictEqual(r.delta, 2000); assert.strictEqual(r.annual, 24000);
  assert.ok(Math.abs(r.rate - 2000 / 70000) < 1e-12);
});
test('T-69', 'R-17 相場との差', function () {
  var r = L.rentRevision({ current: 70000, market: 74000, proposed: 72000, churn: 0.05, vacancyDays: 60 });
  assert.ok(Math.abs(r.gapPct - (-4000 / 74000)) < 1e-12);
});
test('T-70', 'R-17 値上げの期待値', function () {
  var r = L.rentRevision({ current: 70000, market: 74000, proposed: 72000, churn: 0.05, vacancyDays: 60 });
  assert.ok(Math.abs(r.expected - 15800) < 1e-6);
});
test('T-71', 'R-17 値下げは増減額の年換算', function () {
  var r = L.rentRevision({ current: 70000, market: 66000, proposed: 68000, churn: 0.05, vacancyDays: 60 });
  assert.strictEqual(r.annual, -24000); assert.strictEqual(r.expected, -24000);
});

// M. 税の納付（R-18）
function tax(remaining, paid) {
  var due = '2026-12-28';
  return L.taxStatus({ due: due, paidDate: paid || null, today: L.addDays(due, -remaining) });
}
test('T-72', 'R-18 納付済', function () { assert.strictEqual(tax(100, '2026-12-01').status, '納付済'); });
test('T-73', 'R-18 段階の境界', function () {
  assert.strictEqual(tax(31).level, 'NONE'); assert.strictEqual(tax(30).level, 'YELLOW');
  assert.strictEqual(tax(8).level, 'YELLOW'); assert.strictEqual(tax(7).level, 'RED');
  assert.strictEqual(tax(0).level, 'RED'); assert.strictEqual(tax(-1).level, 'OVERDUE');
});
test('T-74', 'R-18 今後90日の納付予定額と超過額', function () {
  var items = [
    { due: '2026-11-10', amount: 10000, paidDate: null },
    { due: '2026-12-28', amount: 37000, paidDate: null },
    { due: '2027-02-21', amount: 37000, paidDate: null },
    { due: '2027-02-22', amount: 5000, paidDate: null },
    { due: '2026-12-01', amount: 99999, paidDate: '2026-11-20' }
  ];
  var r = L.taxReserve(items, '2026-11-23', 90);
  assert.strictEqual(r.overdue, 10000);
  assert.strictEqual(r.upcoming, 74000); // 2027-02-21 は基準日から90日目、02-22 は91日目
});

// N. 経費と請求額（R-19、R-21）
test('T-75', 'R-19 修繕費は「含める」だけ、実施日の月に計上', function () {
  var e = L.summarizeExpenses({ repairs: [{ date: '2026-03-15', amount: 30000, include: true }, { date: '2026-03-20', amount: 180000, include: false }], taxes: [], other: [] }, 2026);
  assert.strictEqual(e.repair[2], 30000);
});
test('T-76', 'R-19 税金は納付日の月に計上', function () {
  var e = L.summarizeExpenses({ repairs: [], taxes: [{ paidDate: '2026-04-28', amount: 38000 }, { paidDate: null, amount: 37000 }], other: [] }, 2026);
  assert.strictEqual(e.tax[3], 38000); assert.strictEqual(e.tax[11], 0);
});
test('T-77', 'R-19 経費の合計', function () {
  var other = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 20000];
  var e = L.summarizeExpenses({ repairs: [{ date: '2026-12-05', amount: 5000, include: true }], taxes: [{ paidDate: '2026-12-28', amount: 37000 }], other: other }, 2026);
  assert.strictEqual(e.total[11], 62000);
});
test('T-78', 'R-21 改定前の月は改定前の家賃', function () {
  var b = { rent: 72000, fee: 5000, prevRent: 70000, revisedFrom: '2026-07-01', year: 2026, month: 6 };
  assert.strictEqual(L.billedFor(b), 75000);
  b.month = 7; assert.strictEqual(L.billedFor(b), 77000);
});
test('T-79', 'R-21 改定日がなければ現在の家賃', function () {
  assert.strictEqual(L.billedFor({ rent: 70000, fee: 5000, year: 2026, month: 1 }), 75000);
});
test('T-80', 'R-21 請求調整が優先', function () {
  assert.strictEqual(L.billedFor({ rent: 72000, fee: 5000, prevRent: 70000, revisedFrom: '2026-07-01', year: 2026, month: 6, adjustment: 40000 }), 40000);
});

test('T-81', 'R-23 評価額が未入力なら価格から推定（特例なし）', function () {
  var r = L.estimatePropertyTax({ price: 30000000, landRatio: 0.3, landSpecial: 'なし' });
  // 土地 30,000,000×0.3×0.7=6,300,000 / 建物 30,000,000×0.7×0.6=12,600,000
  assert.strictEqual(r.landAssessed, 6300000); assert.strictEqual(r.buildingAssessed, 12600000);
  assert.strictEqual(r.fixed, 264600); assert.strictEqual(r.city, 56700); assert.strictEqual(r.total, 321300);
});
test('T-82', 'R-23 小規模住宅用地の特例（固定1/6・都市1/3）', function () {
  var r = L.estimatePropertyTax({ price: 30000000, landRatio: 0.3, landSpecial: '小規模' });
  assert.strictEqual(r.fixed, 14700 + 176400); assert.strictEqual(r.city, 6300 + 37800);
});
test('T-83', 'R-23 一般住宅用地の特例（固定1/3・都市2/3）', function () {
  var r = L.estimatePropertyTax({ landAssessed: 9000000, buildingAssessed: 0, landSpecial: '一般' });
  assert.strictEqual(r.fixed, 42000); assert.strictEqual(r.city, 18000);
});
test('T-84', 'R-23 評価額を入力した場合は推定より優先', function () {
  var r = L.estimatePropertyTax({ price: 99999999, landAssessed: 5000000, buildingAssessed: 10000000, landSpecial: 'なし' });
  assert.strictEqual(r.landAssessed, 5000000); assert.strictEqual(r.fixed, 210000);
});
test('T-85', 'R-24 年間の予定収支＝家賃×(1−空室率)−税−その他', function () {
  var r = L.propertyForecast({ monthlyRent: 300000, vacancyRate: 0.05, tax: 321300, otherExpense: 100000, price: 30000000 });
  assert.strictEqual(r.annualRent, 3600000); assert.strictEqual(r.expectedRent, 3420000);
  assert.strictEqual(r.profit, 3420000 - 321300 - 100000);
  assert.ok(Math.abs(r.grossYield - 0.12) < 1e-9);
});
test('T-86', 'R-24 取得価格が0・空欄なら利回りは出さない', function () {
  assert.strictEqual(L.propertyForecast({ monthlyRent: 100000, tax: 0 , price: 0 }).grossYield, null);
});

test('T-87', 'R-25 税カレンダーに登録があれば入力値を優先', function () {
  var r = L.adoptTax({ registeredCount: 2, registeredSum: 150000, estimated: 235200 });
  assert.strictEqual(r.amount, 150000); assert.strictEqual(r.source, '入力値');
});
test('T-88', 'R-25 登録がなければ推定額を使う', function () {
  var r = L.adoptTax({ registeredCount: 0, registeredSum: 0, estimated: 235200 });
  assert.strictEqual(r.amount, 235200); assert.strictEqual(r.source, '推定');
});
test('T-89', 'R-25 登録額が0円でも、登録があれば入力値を優先', function () {
  var r = L.adoptTax({ registeredCount: 1, registeredSum: 0, estimated: 235200 });
  assert.strictEqual(r.amount, 0); assert.strictEqual(r.source, '入力値');
});

test('T-90', 'R-26 税制が変わったとき、特例の割合と税率を差し替えられる', function () {
  var r = L.estimatePropertyTax({ landAssessed: 6000000, buildingAssessed: 0, landSpecial: '小規模',
    fixedRate: 0.015, cityRate: 0.002, specialRates: { '小規模': [1 / 4, 1 / 2] } });
  assert.strictEqual(r.fixed, 22500); assert.strictEqual(r.city, 6000);
});

test('T-91', 'R-23 割合の小数誤差で1円ずれない', function () {
  var r = L.estimatePropertyTax({ landAssessed: 6300000, buildingAssessed: 12600000, landSpecial: '小規模',
    specialRates: { '小規模': [0.166666666666667, 0.333333333333333] } });
  assert.strictEqual(r.city, 44100);
});

var failed = results.filter(function (r) { return !r.ok; });
results.forEach(function (r) { console.log((r.ok ? 'PASS ' : 'FAIL ') + r.id + (r.ok ? '' : '  ' + r.name + ' : ' + r.msg)); });
console.log('\n' + (results.length - failed.length) + ' / ' + results.length + ' 件が成功');
process.exit(failed.length ? 1 : 0);
