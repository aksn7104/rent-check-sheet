# -*- coding: utf-8 -*-
"""ひな形（.xlsx）の計算式の結果を、src/logic.js の結果と突き合わせる（v2）。
実行: RECALC=<recalc.py のパス> python3 tools/verify_template.py rent-check-sheet-manual.xlsx
（<xlsx>.addr.json に設定セルの位置が入っている）"""
import os, sys, json, shutil, subprocess, datetime as dt
from openpyxl import load_workbook

SRC = sys.argv[1]
RECALC = os.environ['RECALC']
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ADDR = json.load(open(SRC + '.addr.json'))
S, GU0 = ADDR['S'], ADDR['GU0']

def node(expr):
    code = "const L=require('%s/src/logic.js');process.stdout.write(JSON.stringify((%s)));" % (ROOT, expr)
    return json.loads(subprocess.check_output(['node', '-e', code]))

def run(case):
    shutil.copy(SRC, '/tmp/verify.xlsx')
    wb = load_workbook('/tmp/verify.xlsx')
    s = wb['設定']
    s[S['today_in']] = case['today']
    s[S['y']] = case.get('yellow', 1); s[S['r']] = case.get('red', 7)
    wb['入金チェック']['B2'] = case['year']; wb['入金チェック']['B3'] = case['month']
    wb.save('/tmp/verify.xlsx')
    out = subprocess.check_output(['python3', RECALC, '/tmp/verify.xlsx', '180'])
    j = json.loads(out)
    assert j['status'] == 'success', out
    return load_workbook('/tmp/verify.xlsx', data_only=True)

def D(v):
    if isinstance(v, dt.datetime): return v.strftime('%Y-%m-%d')
    if isinstance(v, dt.date): return v.isoformat()
    return v
def iso(v): return D(v) if v else None
total = bad = 0
def check(label, got, exp):
    global total, bad
    total += 1
    if got != exp:
        bad += 1; print('NG', label, '\n  シート :', got, '\n  logic.js:', exp)

cases = [
 {'name': '2026年10月分・基準日 11/23', 'today': dt.date(2026, 11, 23), 'year': 2026, 'month': 10},
 {'name': '2026年11月分・基準日 11/28', 'today': dt.date(2026, 11, 28), 'year': 2026, 'month': 11},
 {'name': '2026年11月分・黄3/赤10・基準日 12/07', 'today': dt.date(2026, 12, 7), 'year': 2026, 'month': 11, 'yellow': 3, 'red': 10},
 {'name': '2026年6月分（改定前の家賃）・基準日 7/10', 'today': dt.date(2026, 7, 10), 'year': 2026, 'month': 6},
 {'name': '2026年12月分・基準日 12/10', 'today': dt.date(2026, 12, 10), 'year': 2026, 'month': 12},
]
SJ = {'PAID': '入金済', 'OVERPAID': '過入金', 'PARTIAL': '一部入金'}
LJ = {'NONE': '', 'YELLOW': '黄', 'RED': '赤'}
wbs = {}
for case in cases:
    wb = run(case); wbs[case['name']] = wb
    c, led, pay, st, adj = wb['入金チェック'], wb['台帳'], wb['入金記録'], wb['設定'], wb['請求調整']
    ty = case['today'].isoformat()
    yl, rd = case.get('yellow', 1), case.get('red', 7)
    gu = {st[f'B{r}'].value: st[f'C{r}'].value for r in range(GU0, GU0 + 20) if st[f'B{r}'].value}
    adjs = [a for a in adj.iter_rows(min_row=5, max_row=504, values_only=True) if a[0]]
    pays_all = [p for p in pay.iter_rows(min_row=5, max_row=3004, max_col=8, values_only=True) if p[0]]
    for r in range(7, 107):
        k = r - 2
        room = led[f'A{k}'].value
        if not room or led[f'O{k}'].value != '入居中':
            check(f'{case["name"]} 行{r} 非表示', c[f'A{r}'].value or '', ''); continue
        rent, fee, dd = led[f'F{k}'].value, led[f'G{k}'].value or 0, led[f'H{k}'].value
        a = [x[3] for x in adjs if x[0] == case['year'] and x[1] == case['month'] and x[2] == room]
        bi = {'year': case['year'], 'month': case['month'], 'rent': rent, 'fee': fee,
              'prevRent': led[f'P{k}'].value, 'revisedFrom': iso(led[f'Q{k}'].value), 'adjustment': a[0] if a else None}
        billed = node('L.billedFor(%s)' % json.dumps(bi))
        pays = [p[4] for p in pays_all if p[1] == room and p[2] == case['year'] and p[3] == case['month'] and p[4]]
        due = node('L.dueDateFor(%d,%d,%s)' % (case['year'], case['month'], json.dumps(dd)))
        j = node("L.judgePayment({billed:%d,payments:%s,dueDate:'%s',today:'%s',yellowDays:%d,redDays:%d})" % (billed, json.dumps(pays), due, ty, yl, rd))
        es = SJ.get(j['status']) or ('期日前' if ty <= due else '未入金')
        got = dict(請求額=c[f'C{r}'].value, 期日=D(c[f'D{r}'].value), 入金額=c[f'E{r}'].value, 状態=c[f'G{r}'].value,
                   不足=c[f'H{r}'].value or 0, 超過=c[f'I{r}'].value or 0, 超過日数=c[f'J{r}'].value, 滞納=c[f'K{r}'].value or '')
        exp = dict(請求額=billed, 期日=due, 入金額=sum(pays), 状態=es, 不足=j['shortage'], 超過=j['excess'], 超過日数=j['overdueDays'], 滞納=LJ[j['level']])
        check(f'{case["name"]} {room}', got, exp)
        days = gu.get(led[f'K{k}'].value)
        ar = node("L.accidentReport({guaranteeType:%s,level:'%s',dueDate:'%s',days:%s,today:'%s'})" % (json.dumps(led[f'L{k}'].value), j['level'], due, days if days is not None else 0, ty))
        exp_mn = ('', '') if ar is None else (ar['deadline'], '報告期限超過' if ar['overdue'] else '残り%d日' % ar['remainingDays'])
        check(f'{case["name"]} {room} 事故報告', (D(c[f'M{r}'].value) or '', c[f'N{r}'].value or ''), exp_mn)
    print('入金チェック 確認済み:', case['name'])

wb = wbs[cases[0]['name']]
T = dt.date(2026, 11, 23); Ts = T.isoformat()
# 期限管理
dl, led = wb['期限管理'], wb['台帳']
for r in range(5, 105):
    if led[f'O{r}'].value != '入居中' or not led[f'A{r}'].value: continue
    for dc, nc, jc in (('C', 'D', 'E'), ('F', 'G', 'H'), ('I', 'J', 'K')):
        v = dl[f'{dc}{r}'].value
        if not v: continue
        rem = (v.date() - T).days
        e = node('L.deadlineLevel(%d,90,30)' % rem)
        check(f'期限管理 {r}{dc}', dl[f'{jc}{r}'].value or '', '期限超過' if e['expired'] else LJ[e['level']])
# 敷金台帳
dp = wb['敷金台帳']
for r in range(5, 205):
    if not dp[f'D{r}'].value: continue
    inp = {'deposit': dp[f'D{r}'].value, 'moveOut': iso(dp[f'F{r}'].value), 'repair': dp[f'G{r}'].value or 0,
           'other': dp[f'H{r}'].value or 0, 'status': dp[f'L{r}'].value, 'today': Ts, 'targetDays': 30}
    e = node('L.depositSettlement(%s)' % json.dumps(inp))
    check(f'敷金 {r}', (dp[f'I{r}'].value, dp[f'J{r}'].value, dp[f'K{r}'].value or 0, dp[f'N{r}'].value),
          (e['deducted'], e['refund'], e['extra'], e['state']))
# 空室
vc = wb['空室・募集']
for r in range(5, 105):
    if not vc[f'A{r}'].value or not vc[f'C{r}'].value: continue
    e = node("L.vacancyStatus({startDate:'%s',today:'%s',rent:%s,yellowDays:30,redDays:60})" % (iso(vc[f'C{r}'].value), Ts, vc[f'E{r}'].value or 'null'))
    check(f'空室 {r}', (vc[f'J{r}'].value, vc[f'K{r}'].value or '', vc[f'L{r}'].value), (e['days'], LJ[e['level']], e['loss']))
# 修繕
rp = wb['修繕履歴']
rows = [(r, rp[f'A{r}'].value, rp[f'B{r}'].value, rp[f'D{r}'].value) for r in range(5, 505) if rp[f'A{r}'].value]
for r, a, room, eq in rows:
    yrs = rp[f'I{r}'].value
    if yrs is None: continue
    later = any(x[2] == room and x[3] == eq and x[1] > a for x in rows)
    e = node("L.equipmentStatus({date:'%s',years:%d,today:'%s',laterExists:%s,yellowDays:180,redDays:60})" % (iso(a), yrs, Ts, 'true' if later else 'false'))
    exp = '更新済' if e['superseded'] else ('目安超過' if e['level'] == 'EXPIRED' else LJ[e['level']])
    check(f'修繕 {r}', (D(rp[f'J{r}'].value), rp[f'L{r}'].value or ''), (e['nextDate'], exp) if not e['superseded'] else (D(rp[f'J{r}'].value), exp))
# 税
tx = wb['税カレンダー']
items = []
for r in range(5, 205):
    if not tx[f'E{r}'].value: continue
    e = node("L.taxStatus({due:'%s',paidDate:%s,today:'%s',yellowDays:30,redDays:7})" % (iso(tx[f'E{r}'].value), json.dumps(iso(tx[f'G{r}'].value)), Ts))
    exp = '納付済' if e['status'] == '納付済' else {'OVERDUE': '納期限超過', 'NONE': '', 'YELLOW': '黄', 'RED': '赤'}[e['level']]
    check(f'税 {r}', (tx[f'H{r}'].value, tx[f'J{r}'].value or ''), (e['status'], exp))
    items.append({'due': iso(tx[f'E{r}'].value), 'amount': tx[f'F{r}'].value, 'paidDate': iso(tx[f'G{r}'].value)})
res = node("L.taxReserve(%s,'%s',90)" % (json.dumps(items), Ts))
print('taxReserve ->', res)
# 家賃改定試算
rv = wb['家賃改定試算']
for r in range(5, 105):
    if not rv[f'A{r}'].value or rv[f'F{r}'].value is None: continue
    e = node('L.rentRevision({current:%d,proposed:%d,market:%d,churn:0.05,vacancyDays:60})' % (rv[f'C{r}'].value, rv[f'F{r}'].value, rv[f'D{r}'].value))
    check(f'改定 {r}', (rv[f'G{r}'].value, round(rv[f'I{r}'].value), round(rv[f'J{r}'].value), round(rv[f'E{r}'].value, 6)),
          (e['delta'], round(e['annual']), round(e['expected']), round(e['gapPct'], 6)))
print('各シート 確認済み')
# 物件台帳（税の推定・予定収支）
pr = wb['物件台帳']; stt = wb['設定']; tx = wb['税カレンダー']
def sv(k): return stt[S[k]].value
plan_props = []
for r in range(5, 25):
    if not pr[f'A{r}'].value: continue
    lr = pr[f'F{r}'].value if pr[f'F{r}'].value is not None else sv('lsh')
    inp = {'price': pr[f'C{r}'].value or 0, 'landRatio': lr, 'landAssessRate': sv('lr'), 'buildingAssessRate': sv('br'),
           'fixedRate': sv('fix'), 'cityRate': sv('city'), 'landSpecial': pr[f'I{r}'].value or 'なし',
           'landAssessed': pr[f'G{r}'].value, 'buildingAssessed': pr[f'H{r}'].value,
           'specialRates': {'小規模': [sv('smf'), sv('smc')], '一般': [sv('gef'), sv('gec')]}}
    e = node('L.estimatePropertyTax(%s)' % json.dumps(inp))
    check(f'物件税 {r}', (pr[f'L{r}'].value, pr[f'M{r}'].value, pr[f'N{r}'].value, pr[f'O{r}'].value, pr[f'P{r}'].value),
          (e['landAssessed'], e['buildingAssessed'], e['fixed'], e['city'], e['total']))
    name = pr[f'A{r}'].value
    mr = sum((x[0] or 0) + (x[1] or 0) for x in [(led[f'F{k}'].value, led[f'G{k}'].value) for k in range(5, 105)
             if led[f'R{k}'].value == name and led[f'O{k}'].value == '入居中'])
    vac = pr[f'J{r}'].value if pr[f'J{r}'].value is not None else sv('pv')
    tx_rows = [(tx[f'C{q}'].value, tx[f'E{q}'].value, tx[f'F{q}'].value) for q in range(5, 205) if tx[f'E{q}'].value]
    reg = [a[2] or 0 for a in tx_rows if a[0] == name and a[1].year == wb['年間収支']['B2'].value]
    ad = node('L.adoptTax(%s)' % json.dumps({'registeredCount': len(reg), 'registeredSum': sum(reg), 'estimated': e['total']}))
    check(f'物件税採用 {r}', (pr[f'X{r}'].value, pr[f'Y{r}'].value), (ad['amount'], ad['source']))
    plan_props.append((name, ad, mr, vac))
    f = node('L.propertyForecast(%s)' % json.dumps({'monthlyRent': mr, 'vacancyRate': vac, 'tax': ad['amount'], 'otherExpense': pr[f'K{r}'].value or 0, 'price': pr[f'C{r}'].value or 0}))
    check(f'物件予定収支 {r}', (pr[f'S{r}'].value, pr[f'T{r}'].value, pr[f'U{r}'].value, round(pr[f'V{r}'].value, 9) if pr[f'V{r}'].value != '' else None),
          (mr, f['expectedRent'], f['profit'], round(f['grossYield'], 9) if f['grossYield'] is not None else None))
# 予定年間収支（月別）
an = wb['年間収支']
est_only = sum(ad_['amount'] for (_, ad_, _, _) in plan_props if ad_['source'] == '推定')
exp_rent = sum(round(mr_ * 12 * (1 - v_)) for (_, _, mr_, v_) in plan_props)
for m in range(1, 13):
    r = 25 + m
    tx_m = sum((x[2] or 0) for x in tx_rows if x[1].year == 2026 and x[1].month == m)
    check(f'予定 {m}月 見込み家賃', an[f'B{r}'].value, round(exp_rent / 12))
    check(f'予定 {m}月 税', an[f'C{r}'].value, tx_m + round(est_only / 12))
print('物件台帳 確認済み')


# 年間収支（収入・預り）
ac, pay = wb['年間収支'], wb['入金記録']
recs = {}
for p in pay.iter_rows(min_row=5, max_row=3004, max_col=12, values_only=True):
    if p[0] and p[4]:
        recs.setdefault((p[1], p[2], p[3]), []).append(p)
for m in range(1, 13):
    inc = dep = 0
    for k, pl in recs.items():
        billed = pl[0][8]
        paid = 0
        for p in sorted(pl, key=lambda x: x[0]):
            take = min(p[4], max(0, billed - paid)); paid += take
            if p[0].year == 2026 and p[0].month == m:
                inc += take; dep += p[4] - take
    check(f'年間 {m}月 収入', ac[f'B{5+m}'].value or 0, inc)
    check(f'年間 {m}月 預り', ac[f'C{5+m}'].value or 0, dep)
print('年間収支 確認済み')
print(f'一致 {total - bad} / {total}')
sys.exit(1 if bad else 0)
