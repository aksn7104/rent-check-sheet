/**
 * 家賃入金チェック表 判定ロジック
 * シート（SpreadsheetApp）に依存しない関数だけを置く。Node.js でテストでき、GAS でもそのまま使える。
 * 仕様は docs/requirements.md の R-01〜R-13 に対応する。
 */

var DAY_MS = 24 * 60 * 60 * 1000;

/** 'YYYY-MM-DD' を UTC の日付（ミリ秒）に変換する。不正な日付は例外にする。 */
function parseDate(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s));
  if (!m) throw new Error('日付の形式が不正です: ' + s);
  var y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  var t = Date.UTC(y, mo - 1, d);
  var chk = new Date(t);
  if (chk.getUTCFullYear() !== y || chk.getUTCMonth() !== mo - 1 || chk.getUTCDate() !== d) {
    throw new Error('存在しない日付です: ' + s);
  }
  return t;
}

function formatDate(t) {
  var d = new Date(t);
  var mm = ('0' + (d.getUTCMonth() + 1)).slice(-2);
  var dd = ('0' + d.getUTCDate()).slice(-2);
  return d.getUTCFullYear() + '-' + mm + '-' + dd;
}

/** b - a の日数 */
function daysBetween(a, b) {
  return Math.round((parseDate(b) - parseDate(a)) / DAY_MS);
}

function addDays(dateStr, n) {
  return formatDate(parseDate(dateStr) + n * DAY_MS);
}

/** R-07: 指定の年月の支払期日。存在しない日付はその月の末日にする。dueDay は 1〜31 または '月末'。 */
function dueDateFor(year, month, dueDay) {
  var last = new Date(Date.UTC(year, month, 0)).getUTCDate();
  var day;
  if (dueDay === '月末') {
    day = last;
  } else {
    day = Number(dueDay);
    if (!(day >= 1 && day <= 31) || Math.floor(day) !== day) throw new Error('支払期日が不正です: ' + dueDay);
    if (day > last) day = last;
  }
  return formatDate(Date.UTC(year, month - 1, day));
}

function sum(arr) {
  var s = 0;
  for (var i = 0; i < arr.length; i++) s += arr[i];
  return s;
}

/** R-05, R-06: 超過日数から滞納の段階を決める */
function overdueLevel(overdueDays, yellowDays, redDays) {
  if (overdueDays > redDays) return 'RED';
  if (overdueDays >= yellowDays && overdueDays > 0) return 'YELLOW';
  return 'NONE';
}

/**
 * R-01〜R-06: 入金の状態を判定する。
 * input: { billed, payments: [金額...], dueDate, today, yellowDays=1, redDays=7 }
 * 戻り値: { status: PAID|PARTIAL|OVERPAID|UNPAID, level: NONE|YELLOW|RED, overdueDays, shortage, excess }
 */
function judgePayment(input) {
  var billed = input.billed;
  if (typeof billed !== 'number' || !isFinite(billed) || billed <= 0) throw new Error('請求額は0より大きい数にしてください');
  var payments = input.payments || [];
  for (var i = 0; i < payments.length; i++) {
    if (typeof payments[i] !== 'number' || !isFinite(payments[i]) || payments[i] < 0) throw new Error('入金額は0以上の数にしてください');
  }
  var yellowDays = input.yellowDays == null ? 1 : input.yellowDays;
  var redDays = input.redDays == null ? 7 : input.redDays;
  var overdue = Math.max(0, daysBetween(input.dueDate, input.today)); // 日付の検証も兼ねる
  var paid = sum(payments);

  if (paid === billed) return { status: 'PAID', level: 'NONE', overdueDays: 0, shortage: 0, excess: 0 };
  if (paid > billed) return { status: 'OVERPAID', level: 'NONE', overdueDays: 0, shortage: 0, excess: paid - billed };
  return {
    status: paid > 0 ? 'PARTIAL' : 'UNPAID',
    level: overdueLevel(overdue, yellowDays, redDays),
    overdueDays: overdue,
    shortage: billed - paid,
    excess: 0
  };
}

/** R-09: 保証会社からの入金を、入居者へ割り当てた合計と突き合わせる。diff = 入金額 − 割り当ての合計 */
function checkAllocation(received, allocations) {
  var diff = received - sum(allocations);
  return { ok: diff === 0, diff: diff };
}

/**
 * R-10: 一般保証型で滞納が赤のとき、事故報告の期限までの残り日数を返す。それ以外は null。
 * input: { guaranteeType, level, dueDate, days, today }
 */
function accidentReport(input) {
  if (input.guaranteeType !== '一般保証型' || input.level !== 'RED') return null;
  var deadline = addDays(input.dueDate, input.days);
  var remaining = daysBetween(input.today, deadline);
  return { deadline: deadline, remainingDays: remaining, overdue: remaining < 0 };
}

/** R-11: 期限までの残り日数から段階を決める（更新・保険・保証会社の契約期限） */
function deadlineLevel(remainingDays, yellowDays, redDays) {
  var y = yellowDays == null ? 90 : yellowDays;
  var r = redDays == null ? 30 : redDays;
  var level = 'NONE';
  if (remainingDays <= r) level = 'RED';
  else if (remainingDays <= y) level = 'YELLOW';
  return { level: level, expired: remainingDays < 0 };
}

/**
 * R-12: 入金日ベースで月別の家賃収入を集計する。請求額を超えた分は「預り」とし、収入に含めない。
 * records: [{ billed, payments: [{ date, amount }] }]（1件＝入居者1人の1か月分）
 * 戻り値: { income: [12か月分], deposit: [12か月分] }
 */
function summarizeIncome(records, year) {
  var income = [], deposit = [];
  for (var m = 0; m < 12; m++) { income.push(0); deposit.push(0); }
  records.forEach(function (rec) {
    var remaining = rec.billed;
    var pays = rec.payments.slice().sort(function (a, b) { return parseDate(a.date) - parseDate(b.date); });
    pays.forEach(function (p) {
      var t = new Date(parseDate(p.date));
      var take = Math.min(p.amount, remaining);
      remaining -= take;
      if (t.getUTCFullYear() === year) {
        income[t.getUTCMonth()] += take;
        deposit[t.getUTCMonth()] += p.amount - take;
      }
    });
  });
  return { income: income, deposit: deposit };
}

/** 年間収支: 月別の収入・経費（12か月分）から合計を出す */
function annualSummary(incomeByMonth, expenseByMonth) {
  var inc = sum(incomeByMonth), exp = sum(expenseByMonth);
  return { income: inc, expense: exp, net: inc - exp };
}

/**
 * R-13: 翌月の入金チェックの行を、台帳の「入居中」の契約から作る。
 * ledger: [{ room, tenant, rent, fee, dueDay, status }]。overrides: { 号室: 請求額 }（退去月などの上書き）
 */
function nextMonthRows(ledger, year, month, overrides) {
  var ov = overrides || {};
  return ledger.filter(function (c) { return c.status === '入居中'; }).map(function (c) {
    var billed = ov[c.room] != null ? ov[c.room] : c.rent + (c.fee || 0);
    return { room: c.room, tenant: c.tenant, billed: billed, dueDate: dueDateFor(year, month, c.dueDay) };
  });
}

/* ===== 追加機能（R-14〜R-22） ===== */

/** 月を足す（Excel の EDATE と同じ。日が存在しない月は月末にする） */
function addMonths(dateStr, n) {
  var t = new Date(parseDate(dateStr));
  var m = t.getUTCMonth() + n, d = t.getUTCDate();
  var ny = t.getUTCFullYear() + Math.floor(m / 12);
  var nm = ((m % 12) + 12) % 12;
  var last = new Date(Date.UTC(ny, nm + 1, 0)).getUTCDate();
  return formatDate(Date.UTC(ny, nm, Math.min(d, last)));
}

/**
 * R-14: 設備の次回の目安。同じ号室・設備に、より新しい記録があれば「更新済」とする。
 * input: { date, years, today, laterExists, yellowDays=180, redDays=60 }
 * 戻り値: { superseded, nextDate, remainingDays, level: NONE|YELLOW|RED|EXPIRED }
 */
function equipmentStatus(input) {
  var yellow = input.yellowDays == null ? 180 : input.yellowDays;
  var red = input.redDays == null ? 60 : input.redDays;
  if (input.laterExists) return { superseded: true, nextDate: null, remainingDays: null, level: 'NONE' };
  if (input.years == null) return { superseded: false, nextDate: null, remainingDays: null, level: 'NONE' };
  var next = addMonths(input.date, 12 * input.years);
  var remaining = daysBetween(input.today, next);
  var level = remaining < 0 ? 'EXPIRED' : remaining <= red ? 'RED' : remaining <= yellow ? 'YELLOW' : 'NONE';
  return { superseded: false, nextDate: next, remainingDays: remaining, level: level };
}

/**
 * R-15: 空室日数と段階、空室損失（概算）。日数が黄・赤の開始日数以上で段階を上げる。
 * input: { startDate, rent, today, yellowDays=30, redDays=60 }
 */
function vacancyStatus(input) {
  var yellow = input.yellowDays == null ? 30 : input.yellowDays;
  var red = input.redDays == null ? 60 : input.redDays;
  var days = Math.max(0, daysBetween(input.startDate, input.today));
  var level = days >= red ? 'RED' : days >= yellow ? 'YELLOW' : 'NONE';
  var loss = input.rent == null ? null : Math.round(input.rent * days / 30);
  return { days: days, level: level, loss: loss };
}

/**
 * R-16: 敷金・保証金の精算。
 * input: { deposit, repair, other, moveOut, today, targetDays=30, status('預り中'|'精算済') }
 */
function depositSettlement(input) {
  var deducted = (input.repair || 0) + (input.other || 0);
  var target = input.targetDays == null ? 30 : input.targetDays;
  var settleBy = input.moveOut ? addDays(input.moveOut, target) : null;
  var state;
  if (input.status === '精算済') state = '精算済';
  else if (!input.moveOut) state = '預り中';
  else state = daysBetween(settleBy, input.today) > 0 ? '精算遅れ' : '精算期間内';
  return {
    deducted: deducted,
    refund: Math.max(0, input.deposit - deducted),
    extra: Math.max(0, deducted - input.deposit),
    settleBy: settleBy,
    state: state
  };
}

/**
 * R-17: 家賃改定の試算。
 * input: { current, market, proposed, churn, vacancyDays }
 * 値上げのときは、想定退去率で、増額の期待値から退去時の空室損失を引く。値下げは増減額の年換算。
 */
function rentRevision(input) {
  var delta = input.proposed - input.current;
  var annual = delta * 12;
  var expected = delta > 0
    ? annual * (1 - input.churn) - input.churn * input.current * input.vacancyDays / 30
    : annual;
  return {
    gapPct: input.market ? (input.current - input.market) / input.market : null,
    delta: delta,
    rate: delta / input.current,
    annual: annual,
    expected: expected
  };
}

/**
 * R-18: 税の納付の状態。
 * input: { due, paidDate, today, yellowDays=30, redDays=7 }
 */
function taxStatus(input) {
  var yellow = input.yellowDays == null ? 30 : input.yellowDays;
  var red = input.redDays == null ? 7 : input.redDays;
  if (input.paidDate) return { status: '納付済', remainingDays: null, level: 'NONE' };
  var remaining = daysBetween(input.today, input.due);
  var level = remaining < 0 ? 'OVERDUE' : remaining <= red ? 'RED' : remaining <= yellow ? 'YELLOW' : 'NONE';
  return { status: '未納', remainingDays: remaining, level: level };
}

/** R-18: 今後 windowDays 日以内の納付予定額と、納期限を過ぎた未納額 */
function taxReserve(items, today, windowDays) {
  var upcoming = 0, overdue = 0;
  items.forEach(function (it) {
    if (it.paidDate) return;
    var r = daysBetween(today, it.due);
    if (r < 0) overdue += it.amount;
    else if (r <= windowDays) upcoming += it.amount;
  });
  return { upcoming: upcoming, overdue: overdue };
}

/**
 * R-19: 月別の経費。修繕費（「含める」のもの、実施日の月）、税金（納付日の月）、その他（手入力）の合計。
 * input: { repairs: [{date, amount, include}], taxes: [{paidDate, amount}], other: [12か月分] }
 */
function summarizeExpenses(input, year) {
  var repair = [], tax = [], total = [];
  for (var m = 0; m < 12; m++) { repair.push(0); tax.push(0); }
  input.repairs.forEach(function (r) {
    var t = new Date(parseDate(r.date));
    if (r.include && t.getUTCFullYear() === year) repair[t.getUTCMonth()] += r.amount;
  });
  input.taxes.forEach(function (x) {
    if (!x.paidDate) return;
    var t = new Date(parseDate(x.paidDate));
    if (t.getUTCFullYear() === year) tax[t.getUTCMonth()] += x.amount;
  });
  for (var i = 0; i < 12; i++) total.push(repair[i] + tax[i] + (input.other[i] || 0));
  return { repair: repair, tax: tax, total: total };
}

/**
 * R-21: ある月の請求額。その月の請求調整があればそれを使う。なければ、家賃の改定日より前の月は改定前の家賃を使う。
 * input: { rent, fee, prevRent, revisedFrom, year, month, adjustment }
 */
function billedFor(input) {
  if (input.adjustment != null) return input.adjustment;
  var useOld = input.revisedFrom && Date.UTC(input.year, input.month - 1, 1) < parseDate(input.revisedFrom);
  return (useOld ? input.prevRent : input.rent) + (input.fee || 0);
}

/**
 * 物件の固定資産税・都市計画税の推定（目安）。
 * 評価額が未入力なら、取得価格×土地割合×評価割合で推定する。住宅用地の特例は
 * 小規模(固定1/6・都市1/3)、一般(固定1/3・都市2/3)、なし(特例なし)。
 */
function estimatePropertyTax(input) {
  var landRatio = input.landRatio == null ? 0.3 : input.landRatio;
  var landRate = input.landAssessRate == null ? 0.7 : input.landAssessRate;
  var bldgRate = input.buildingAssessRate == null ? 0.6 : input.buildingAssessRate;
  var fixedRate = input.fixedRate == null ? 0.014 : input.fixedRate;
  var cityRate = input.cityRate == null ? 0.003 : input.cityRate;
  var land = input.landAssessed != null ? input.landAssessed : Math.round(input.price * landRatio * landRate);
  var bldg = input.buildingAssessed != null ? input.buildingAssessed : Math.round(input.price * (1 - landRatio) * bldgRate);
  var sp = input.specialRates || {};
  var table = { '小規模': sp['小規模'] || [1 / 6, 1 / 3], '一般': sp['一般'] || [1 / 3, 2 / 3], 'なし': [1, 1] };
  var f = table[input.landSpecial || 'なし'] || [1, 1];
  // 1/3 などの割合で生じる小数の誤差（例：44099.9999…）で1円ずれないよう、小数第4位で丸めてから切り捨てる
  var floor4 = function (x) { return Math.floor(Math.round(x * 1e4) / 1e4); };
  var fixed = floor4(land * f[0] * fixedRate + bldg * fixedRate);
  var city = floor4(land * f[1] * cityRate + bldg * cityRate);
  return { landAssessed: land, buildingAssessed: bldg, fixed: fixed, city: city, total: fixed + city };
}
/** 物件の年間の予定収支（目安）。年間見込み家賃＝現在の年間家賃×(1−想定空室率)。 */
function propertyForecast(input) {
  var annualRent = input.monthlyRent * 12;
  var expected = Math.round(annualRent * (1 - (input.vacancyRate || 0)));
  var other = input.otherExpense || 0;
  return {
    annualRent: annualRent,
    expectedRent: expected,
    profit: expected - input.tax - other,
    grossYield: input.price ? annualRent / input.price : null
  };
}

/** 物件の税の採用額。納付予定が登録されていれば登録額を優先し、なければ推定額を使う。 */
function adoptTax(input) {
  var has = (input.registeredCount || 0) > 0;
  return { amount: has ? input.registeredSum : input.estimated, source: has ? '入力値' : '推定' };
}

var RentLogic = {
  addMonths: addMonths,
  equipmentStatus: equipmentStatus,
  vacancyStatus: vacancyStatus,
  depositSettlement: depositSettlement,
  rentRevision: rentRevision,
  taxStatus: taxStatus,
  taxReserve: taxReserve,
  summarizeExpenses: summarizeExpenses,
  billedFor: billedFor,
  estimatePropertyTax: estimatePropertyTax,
  propertyForecast: propertyForecast,
  adoptTax: adoptTax,
  parseDate: parseDate,
  formatDate: formatDate,
  daysBetween: daysBetween,
  addDays: addDays,
  dueDateFor: dueDateFor,
  overdueLevel: overdueLevel,
  judgePayment: judgePayment,
  checkAllocation: checkAllocation,
  accidentReport: accidentReport,
  deadlineLevel: deadlineLevel,
  summarizeIncome: summarizeIncome,
  annualSummary: annualSummary,
  nextMonthRows: nextMonthRows
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = RentLogic;
}
