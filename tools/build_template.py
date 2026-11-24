# -*- coding: utf-8 -*-
"""家賃入金チェック表（手入力版）のひな形（.xlsx）を作る。
実行: python3 tools/build_template.py [出力パス]
サンプルデータはすべて架空。仕様は docs/requirements.md（R-01〜R-22）。"""
import sys, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

OUT = sys.argv[1] if len(sys.argv) > 1 else 'rent-check-sheet-manual.xlsx'
FONT = 'Meiryo'
D = dt.date

# ---- 容量 ----
LED0, LED1 = 5, 104          # 台帳（契約100件）
PAY0, PAY1 = 5, 3004         # 入金記録
ADJ0, ADJ1 = 5, 504          # 請求調整
GUA0, GUA1 = 5, 304          # 保証会社入金
DEP0, DEP1 = 5, 204          # 敷金台帳
VAC0, VAC1 = 5, 104          # 空室・募集
REP0, REP1 = 5, 504          # 修繕履歴
TAX0, TAX1 = 5, 204          # 税カレンダー
REV0, REV1 = 5, 104          # 家賃改定試算
PRO0, PRO1 = 5, 24           # 物件台帳（20物件）
CHK0 = 7                     # 入金チェックの先頭行（台帳の5行目に対応）
CHK1 = CHK0 + (LED1 - LED0)  # 106

# ---- シート名 ----
SH = dict(guide='使い方', dash='ダッシュボード', set='設定', led='台帳', pay='入金記録', adj='請求調整',
          chk='入金チェック', gua='保証会社入金', due='期限管理', dep='敷金台帳', vac='空室・募集',
          rep='修繕履歴', tax='税カレンダー', rev='家賃改定試算', ann='年間収支', pro='物件台帳')
def q(key, ref):
    return "'%s'!%s" % (SH[key], ref)
def rng(key, col, r0, r1):
    return "'%s'!$%s$%d:$%s$%d" % (SH[key], col, r0, col, r1)

# ---- 色と書式 ----
GREEN = '1F5A44'; GREEN_L = 'E4F1E9'; CREAM = 'FBF6E9'; ORANGE = 'E8872B'; ORANGE_L = 'FFF1DC'
f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_title = Font(name=FONT, size=14, bold=True, color=GREEN)
f_input = Font(name=FONT, size=10, color='0000FF')
f_head = Font(name=FONT, size=10, bold=True, color='FFFFFF')
f_note = Font(name=FONT, size=9, color='555555')
fill_head = PatternFill('solid', fgColor=GREEN)
fill_input = PatternFill('solid', fgColor='FFF9DB')
fill_band = PatternFill('solid', fgColor=GREEN_L)
fill_cream = PatternFill('solid', fgColor=CREAM)
fill_warn = PatternFill('solid', fgColor=ORANGE_L)
fill_red = PatternFill('solid', bgColor='F8CBCB', fgColor='F8CBCB')
fill_yel = PatternFill('solid', bgColor='FFF0A8', fgColor='FFF0A8')
fill_grn = PatternFill('solid', bgColor='D5EBD9', fgColor='D5EBD9')
fill_gray = PatternFill('solid', fgColor='F1F3F4')
thin = Side(style='thin', color='C9D1D6')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
DATE = 'yyyy-mm-dd'
YEN = '#,##0;-#,##0;0'
YEN0 = '#,##0;-#,##0;;'
PCT = '0.0%'
CENTER = Alignment(horizontal='center', vertical='center')
WRAP = Alignment(wrap_text=True, vertical='top')
WRAPC = Alignment(wrap_text=True, vertical='center')

def setc(ws, ref, value=None, font=f_base, fmt=None, fill=None, align=None, b=True):
    c = ws[ref]
    if value is not None: c.value = value
    c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if align: c.alignment = align
    if b: c.border = border
    return c

def header(ws, row, labels, widths=None):
    for i, t in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=t)
        c.font = f_head; c.fill = fill_head; c.border = border
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[row].height = 32
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[ws.cell(row=row, column=i).column_letter].width = w

def input_cell(ws, ref, value=None, fmt=None, text=False):
    return setc(ws, ref, value, font=f_input, fmt=('@' if text else fmt), fill=fill_input)

def auto_cell(ws, ref, formula, fmt=None, align=None):
    return setc(ws, ref, formula, fmt=fmt, align=align, fill=fill_gray)

def title(ws, text, note=None, note_row=2):
    ws['A1'].value = text; ws['A1'].font = f_title
    if note:
        ws.cell(row=note_row, column=1, value=note).font = f_note

def dv_list(ws, formula, cells, allow_blank=True):
    dv = DataValidation(type='list', formula1=formula, allow_blank=allow_blank)
    ws.add_data_validation(dv); dv.add(cells)

def cf(ws, cells, formula, fill):
    ws.conditional_formatting.add(cells, FormulaRule(formula=[formula], fill=fill))

wb = Workbook()
ws_guide = wb.active; ws_guide.title = SH['guide']
ws_dash = wb.create_sheet(SH['dash'])
ws_set = wb.create_sheet(SH['set'])

# ============================================================ 設定
ws = ws_set
ws.column_dimensions['A'].width = 2; ws.column_dimensions['B'].width = 38; ws.column_dimensions['C'].width = 16
ws.column_dimensions['D'].width = 7; ws.column_dimensions['E'].width = 72; ws.column_dimensions['F'].width = 30
ws['B1'].value = '設定'; ws['B1'].font = f_title
ws.merge_cells('B2:F2')
ws['B2'].value = '黄色い背景・青い文字のセルだけを書き換えます。値を変えると、すべてのシートの判定にすぐ反映されます。税制や法制が変わったときは、⑩を書き換え、右端の欄に確認した日と出典をメモします（手順は「使い方」シート）。'
ws['B2'].font = f_base; ws['B2'].alignment = WRAPC; ws['B2'].fill = fill_cream
ws.merge_cells('B3:F3')
ws['B3'].value = '注意：自動入金確認は、導入後の最初の数か月は、通帳や銀行の明細と必ず突き合わせて確認してください。'
ws['B3'].font = Font(name=FONT, size=10, bold=True, color='9A4B00'); ws['B3'].alignment = WRAPC; ws['B3'].fill = fill_warn
ws.row_dimensions[2].height = 34; ws.row_dimensions[3].height = 24

sections = [
 ('① 基準日（「今日」をいつにするか）', [
  ('today_in', '基準日を指定する（空欄でOK）', None, '日付', DATE, 'ふだんは空欄のままにします（自動で今日の日付を使います）。過去や未来の日付で確認したいときだけ、日付を入れてください。'),
  ('today', '実際に使われる基準日（自動）', 'AUTO', '日付', DATE, '上の欄が空欄なら今日、入力があればその日付です。このセルは書き換えません。')]),
 ('② 家賃の滞納（入金チェックの「滞納」の色）', [
  ('y', '黄色にする日数', 1, '日', '0', '支払期日からこの日数がたつと、黄色の警告になります。例：1 → 期日の翌日から黄色。'),
  ('r', '赤にする日数', 7, '日', '0', '期日からの日数がこの値を超えると、赤の警告になります。例：7 → 8日目から赤。')]),
 ('③ 契約更新・火災保険・保証会社の契約の期限（期限管理）', [
  ('dy', '黄色にする残り日数', 90, '日', '0', '期限までの残りがこの日数以下になると黄色。例：90 → 約3か月前から。'),
  ('dr', '赤にする残り日数', 30, '日', '0', '残りがこの日数以下、または期限を過ぎると赤。例：30 → 約1か月前から。')]),
 ('④ 設備の交換の目安（修繕履歴）', [
  ('ey', '黄色にする残り日数', 180, '日', '0', '次回の目安日までの残りがこの日数以下になると黄色。例：180 → 約半年前から。'),
  ('er', '赤にする残り日数', 60, '日', '0', '残りがこの日数以下になると赤。目安日を過ぎると「目安超過」。')]),
 ('⑤ 空室（空室・募集）', [
  ('vy', '黄色にする空室日数', 30, '日', '0', '空室になってからこの日数以上たつと黄色。例：30 → 1か月以上空いたら黄色。'),
  ('vr', '赤にする空室日数', 60, '日', '0', '空室になってからこの日数以上たつと赤。例：60 → 2か月以上空いたら赤。')]),
 ('⑥ 敷金・保証金の精算（敷金台帳）', [
  ('dt', '精算の目安日数', 30, '日', '0', '退去日からこの日数を、精算の目安にします。目安を過ぎても未精算だと「精算遅れ」と表示します。ご自身の運用の目安であり、法律上の期限ではありません。')]),
 ('⑦ 固定資産税などの納付（税カレンダー）', [
  ('ty', '黄色にする残り日数', 30, '日', '0', '納期限までの残りがこの日数以下になると黄色。'),
  ('tr', '赤にする残り日数', 7, '日', '0', '残りがこの日数以下になると赤。納期限を過ぎると「納期限超過」。'),
  ('tw', '納付予定額を見る日数', 90, '日', '0', '今日から何日先までの納付予定額を合計するか。ダッシュボードと税カレンダーに出ます。')]),
 ('⑩ 税制・法制の値（制度が変わったら、ここを書き換えます）', [
  ('fix', '固定資産税の税率', 0.014, '', '0.00%', '標準税率は1.4%です。自治体の条例で異なる場合や、法改正があった場合は書き換えます。'),
  ('city', '都市計画税の税率', 0.003, '', '0.00%', '制限税率（上限）は0.3%で、自治体ごとに決まります。課税されない地域では0%にします。'),
  ('smf', '小規模住宅用地：固定資産税の課税標準の割合', 1/6, '', '# ?/?', '住宅1戸あたり200㎡以下の部分。現行は1/6です。変えるときは「=1/6」のように、イコールから入力します。'),
  ('smc', '小規模住宅用地：都市計画税の課税標準の割合', 1/3, '', '# ?/?', '現行は1/3です。'),
  ('gef', '一般住宅用地：固定資産税の課税標準の割合', 1/3, '', '# ?/?', '200㎡を超える部分。現行は1/3です。'),
  ('gec', '一般住宅用地：都市計画税の課税標準の割合', 2/3, '', '# ?/?', '現行は2/3です。'),
  ('sqm', '小規模住宅用地の面積の上限（1戸あたり）', 200, '㎡', '0', '物件台帳の説明文に使います（計算には使いません）。現行は200㎡です。')]),
 ('⑪ 税の推定の前提（制度ではなく、推定のための仮定）', [
  ('lr', '土地の評価割合（価格に対する）', 0.7, '', '0%', '評価額を入力しない物件で、取得価格から評価額を推定するときの割合です。例：70%。'),
  ('br', '建物の評価割合（価格に対する）', 0.6, '', '0%', '同じく建物の割合です。例：60%。築年数で大きく変わるため、通知書の評価額を入力する方が正確です。'),
  ('lsh', '土地の割合の初期値', 0.3, '', '0%', '物件台帳で土地の割合を空欄にした物件に使います。例：30%。'),
  ('pv', '想定空室率の初期値', 0.05, '', '0.0%', '物件台帳で想定空室率を空欄にした物件に使います。例：5%。')]),
 ('⑧ 家賃改定の試算（家賃改定試算）', [
  ('churn', '想定退去率', 0.05, '', '0.0%', '値上げしたときに、退去する人の割合の見込みです。例：5%。試算の結果は、この仮定に左右されます。'),
  ('vdays', '退去時の空室想定日数', 60, '日', '0', '値上げで退去した場合に、次の入居者が決まるまでの日数の見込みです。')]),
]
S = {}
row = 5
header(ws, row, ['', '項目', '値', '単位', '説明', '確認日・出典（メモ）'])
ws.cell(row=row, column=1).fill = PatternFill(fill_type=None); ws.cell(row=row, column=1).border = Border()
row += 1
for sec_title, items in sections:
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
    c = ws.cell(row=row, column=2, value=sec_title); c.font = Font(name=FONT, size=10, bold=True, color=GREEN); c.fill = fill_band
    ws.row_dimensions[row].height = 22
    row += 1
    for key, label, val, unit, fmt, desc in items:
        setc(ws, f'B{row}', label)
        if val == 'AUTO':
            auto_cell(ws, f'C{row}', f'=IF($C${row-1}="",TODAY(),$C${row-1})', fmt=fmt)
        else:
            input_cell(ws, f'C{row}', val, fmt=fmt)
        setc(ws, f'D{row}', unit, align=CENTER)
        setc(ws, f'E{row}', desc, font=f_note, align=WRAPC)
        input_cell(ws, f'F{row}', '2026年度の制度で設定（初期値）' if sec_title.startswith('⑩') else None)
        ws[f'F{row}'].alignment = WRAPC
        ws.row_dimensions[row].height = 30
        S[key] = "'%s'!$C$%d" % (SH['set'], row)
        row += 1
    row += 1
# 保証会社
ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
c = ws.cell(row=row, column=2, value='⑫ 保証会社ごとの事故報告の期限（一般保証型の入居者が滞納で赤になったとき）')
c.font = Font(name=FONT, size=10, bold=True, color=GREEN); c.fill = fill_band
row += 1
ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
ws.cell(row=row, column=2, value='保証会社に滞納を報告する期限は、会社と契約で異なります。契約書で確認して、会社名と日数を入力してください（最大20社）。').font = f_note
row += 1
GU_HEAD = row
for i, t in enumerate(['保証会社名', '報告期限（日）', '', '備考'], start=2):
    c = ws.cell(row=row, column=i, value=t); c.font = f_head; c.fill = fill_head; c.border = border; c.alignment = CENTER
GU0 = row + 1
for i in range(20):
    r = GU0 + i
    input_cell(ws, f'B{r}', 'サンプル保証会社' if i == 0 else None)
    input_cell(ws, f'C{r}', 14 if i == 0 else None, fmt='0')
    setc(ws, f'D{r}', None)
    input_cell(ws, f'E{r}', 'サンプル（架空）。契約書の日数に書き換えてください。' if i == 0 else None)
GU1 = GU0 + 19
GU_NAMES = "'%s'!$B$%d:$B$%d" % (SH['set'], GU0, GU1)
GU_DAYS = "'%s'!$C$%d:$C$%d" % (SH['set'], GU0, GU1)
ws.freeze_panes = 'A6'
TODAY = S['today']

# ============================================================ 台帳
ws = wb.create_sheet(SH['led'])
title(ws, '台帳（1行＝1契約。100件まで）',
      '号室は文字で入力し、重複させません。支払期日は 1〜31 の数字、または「月末」。退去した契約は削除せず、入居状況を「退去済」にします。')
header(ws, 4, ['号室', '入居者名', '登録名義（振込人名。複数可）', '電話', 'メール', '家賃', '共益費', '支払期日', '契約開始日', '契約更新日',
               '保証会社名', '保証の型', '火災保険の期限', '保証会社の契約期限', '入居状況', '改定前の家賃', '家賃の改定日', '物件'],
       [8, 14, 26, 14, 24, 11, 10, 10, 12, 12, 18, 12, 14, 16, 10, 12, 14, 18])
ledger = [
 ('101', '青木 一郎', 'アオキ イチロウ', '090-0000-0001', 'aoki@example.com', 70000, 5000, 27, D(2025,4,1), D(2027,3,31), 'サンプル保証会社', '一般保証型', D(2027,2,10), D(2027,3,31), '入居中', 68000, D(2026,7,1), '本館サンプル'),
 ('102', '伊藤 花子', 'イトウ ハナコ', '090-0000-0002', 'ito@example.com', 65000, 5000, 27, D(2024,9,1), D(2027,1,20), '', '保証なし', D(2027,1,20), None, '入居中', None, None, '本館サンプル'),
 ('103', '上田 健', 'ウエダ ケン', '090-0000-0003', 'ueda@example.com', 70000, 0, '月末', D(2024,12,16), D(2026,12,15), 'サンプル保証会社', '支払委託型', D(2026,12,15), D(2026,12,15), '入居中', None, None, '本館サンプル'),
 ('104', '遠藤 美咲', 'エンドウ ミサキ', '090-0000-0004', 'endo@example.com', 65000, 5000, '月末', D(2025,7,1), D(2027,6,30), 'サンプル保証会社', '支払委託型', D(2027,6,30), D(2027,6,30), '入居中', None, None, '本館サンプル'),
 ('105', '岡田 翔', 'オカダ ショウ', '090-0000-0005', 'okada@example.com', 72000, 3000, 25, D(2025,2,1), D(2027,1,31), 'サンプル保証会社', '一般保証型', D(2027,1,31), D(2027,1,31), '入居中', None, None, '本館サンプル'),
 ('106', '加藤 彩', 'カトウ アヤ', '090-0000-0006', 'kato@example.com', 60000, 0, 27, D(2023,5,1), D(2025,4,30), '', '保証なし', None, None, '退去済', None, None, '本館サンプル'),
 ('107', '森 優', 'モリ ユウ', '090-0000-0007', 'mori@example.com', 65000, 0, 27, D(2024,3,1), D(2026,10,31), '', '保証なし', None, None, '退去済', None, None, '本館サンプル'),
]
for i in range(LED1 - LED0 + 1):
    r = LED0 + i
    rowv = ledger[i] if i < len(ledger) else [None] * 18
    for j, v in enumerate(rowv):
        col = ws.cell(row=r, column=j + 1).column_letter
        fmt = YEN if col in 'FGP' else (DATE if col in 'IJMNQ' else None)
        c = input_cell(ws, f'{col}{r}', v, fmt=fmt)
        if col == 'A': c.number_format = '@'
ws.freeze_panes = 'C5'
dv_list(ws, '"一般保証型,支払委託型,保証なし"', f'L{LED0}:L{LED1}')
dv_list(ws, '"入居中,退去済"', f'O{LED0}:O{LED1}')
dv_list(ws, '"月末,' + ','.join(str(i) for i in range(1, 32)) + '"', f'H{LED0}:H{LED1}')
dv = DataValidation(type='custom', formula1=f'COUNTIF($A${LED0}:$A${LED1},A{LED0})=1', allow_blank=True,
                    errorTitle='号室が重複しています', error='同じ号室は1行だけにしてください。退去済の行を上書きして使います。', showErrorMessage=True)
ws.add_data_validation(dv); dv.add(f'A{LED0}:A{LED1}')
PRO_NAMES = rng('pro', 'A', PRO0, PRO1)
dv_list(ws, '=' + PRO_NAMES, f'R{LED0}:R{LED1}')
ROOMS = rng('led', 'A', LED0, LED1)
from openpyxl.workbook.defined_name import DefinedName
_dn = DefinedName('号室一覧', attr_text="'%s'!$A$%d:$A$%d" % (SH['led'], LED0, LED1))
try:
    wb.defined_names['号室一覧'] = _dn
except TypeError:
    wb.defined_names.append(_dn)
ROOM_LIST = '=号室一覧'
def room_src(ws):
    # 同じシートの右端に、台帳の号室を写した列（非表示）を作り、プルダウンはそこから選ぶ。
    # Excel 形式のまま Google スプレッドシートで開くと、他のシートを参照するプルダウンが動かないため。
    from openpyxl.utils import get_column_letter
    col = get_column_letter(ws.max_column + 2)
    ws[f'{col}4'].value = '号室一覧（自動）'; ws[f'{col}4'].font = f_note
    for i in range(LED1 - LED0 + 1):
        ws[f'{col}{5 + i}'].value = f"=IF('{SH['led']}'!$A${LED0 + i}=\"\",\"\",'{SH['led']}'!$A${LED0 + i})"
    ws.column_dimensions[col].hidden = True
    return f'=${col}$5:${col}${5 + LED1 - LED0}'

# ============================================================ 物件台帳
ws = wb.create_sheet(SH['pro'])
ws['A1'].value = '物件台帳（20物件まで）'; ws['A1'].font = f_title
ws['A2'].value = ('物件ごとの取得価格と税の推定、予定収支を管理します。台帳（契約）の「物件」列に物件名を入れると、戸数と家賃が自動で集計されます。'
                  '税の推定は目安です。納税通知書の評価額を入力すると、推定より正確になります。')
ws['A2'].font = f_note
ws['A3'].value = '="特例の目安：土地に住宅が建っていれば、住宅1戸あたり"&' + S['sqm'] + '&"㎡以下の部分は「小規模」、それを超える部分は「一般」です。割合と税率は設定シートの⑩で変えられます。迷うときは税理士や自治体の通知書で確認してください。"'; ws['A3'].font = f_note
header(ws, 4, ['物件名', '所在地', '取得価格', '取得日', '総戸数', '土地の割合', '土地の評価額（任意）', '建物の評価額（任意）', '住宅用地の特例', '想定空室率', 'その他の年間経費',
               '土地の評価額（使用値）', '建物の評価額（使用値）', '固定資産税（推定）', '都市計画税（推定）', '税の合計（推定）', '入居中の戸数', '空室数', '月額家賃の合計', '年間の見込み家賃', '年間の予定収支', '表面利回り', '税カレンダーの登録額（対象年）', '採用する税', '税の根拠', '推定で補う税'],
       [18, 18, 14, 12, 8, 10, 16, 16, 14, 10, 14, 16, 16, 14, 14, 14, 10, 8, 14, 16, 16, 10, 18, 14, 10, 14])
ws.row_dimensions[4].height = 44
LD_PRO, LD_ST = rng('led', 'R', LED0, LED1), rng('led', 'O', LED0, LED1)
pros = [('本館サンプル', '架空市1-2-3', 30000000, D(2020,4,1), 10, 0.3, None, None, '小規模', None, 100000)]
for i in range(PRO1 - PRO0 + 1):
    r = PRO0 + i
    v = pros[i] if i < len(pros) else [None] * 11
    input_cell(ws, f'A{r}', v[0]); input_cell(ws, f'B{r}', v[1]); input_cell(ws, f'C{r}', v[2], fmt=YEN); input_cell(ws, f'D{r}', v[3], fmt=DATE)
    input_cell(ws, f'E{r}', v[4], fmt='0'); input_cell(ws, f'F{r}', v[5], fmt='0%'); input_cell(ws, f'G{r}', v[6], fmt=YEN); input_cell(ws, f'H{r}', v[7], fmt=YEN)
    input_cell(ws, f'I{r}', v[8]); input_cell(ws, f'J{r}', v[9], fmt=PCT); input_cell(ws, f'K{r}', v[10], fmt=YEN)
    lr = f'IF($F{r}="",{S["lsh"]},$F{r})'
    auto_cell(ws, f'L{r}', f'=IF($A{r}="","",IF($G{r}<>"",$G{r},ROUND(N($C{r})*{lr}*{S["lr"]},0)))', fmt=YEN)
    auto_cell(ws, f'M{r}', f'=IF($A{r}="","",IF($H{r}<>"",$H{r},ROUND(N($C{r})*(1-{lr})*{S["br"]},0)))', fmt=YEN)
    ff = f'IF($I{r}="小規模",{S["smf"]},IF($I{r}="一般",{S["gef"]},1))'; fc = f'IF($I{r}="小規模",{S["smc"]},IF($I{r}="一般",{S["gec"]},1))'
    auto_cell(ws, f'N{r}', f'=IF($A{r}="","",INT(ROUND($L{r}*{ff}*{S["fix"]}+$M{r}*{S["fix"]},4)))', fmt=YEN)
    auto_cell(ws, f'O{r}', f'=IF($A{r}="","",INT(ROUND($L{r}*{fc}*{S["city"]}+$M{r}*{S["city"]},4)))', fmt=YEN)
    auto_cell(ws, f'P{r}', f'=IF($A{r}="","",$N{r}+$O{r})', fmt=YEN)
    auto_cell(ws, f'Q{r}', f'=IF($A{r}="","",COUNTIFS({LD_PRO},$A{r},{LD_ST},"入居中"))', fmt='0')
    auto_cell(ws, f'R{r}', f'=IF($A{r}="","",IF($E{r}="","",MAX(0,$E{r}-$Q{r})))', fmt='0')
    rentsum = (f"SUMIFS({rng('led','F',LED0,LED1)},{LD_PRO},$A{r},{LD_ST},\"入居中\")+SUMIFS({rng('led','G',LED0,LED1)},{LD_PRO},$A{r},{LD_ST},\"入居中\")")
    auto_cell(ws, f'S{r}', f'=IF($A{r}="","",{rentsum})', fmt=YEN)
    auto_cell(ws, f'T{r}', f'=IF($A{r}="","",ROUND($S{r}*12*(1-IF($J{r}="",{S["pv"]},$J{r})),0))', fmt=YEN)
    ty_s = f"DATE('{SH['ann']}'!$B$2,1,1)"; ty_e = f"DATE('{SH['ann']}'!$B$2,12,31)"
    cnt = f"COUNTIFS({rng('tax','C',TAX0,TAX1)},$A{r},{rng('tax','E',TAX0,TAX1)},\">=\"&{ty_s},{rng('tax','E',TAX0,TAX1)},\"<=\"&{ty_e})"
    auto_cell(ws, f'W{r}', f'=IF($A{r}="","",SUMIFS({rng("tax","F",TAX0,TAX1)},{rng("tax","C",TAX0,TAX1)},$A{r},{rng("tax","E",TAX0,TAX1)},">="&{ty_s},{rng("tax","E",TAX0,TAX1)},"<="&{ty_e}))', fmt=YEN)
    auto_cell(ws, f'X{r}', f'=IF($A{r}="","",IF({cnt}>0,$W{r},$P{r}))', fmt=YEN)
    auto_cell(ws, f'Y{r}', f'=IF($A{r}="","",IF({cnt}>0,"入力値","推定"))', align=CENTER)
    auto_cell(ws, f'Z{r}', f'=IF($A{r}="","",IF($Y{r}="推定",$P{r},0))', fmt=YEN0)
    auto_cell(ws, f'U{r}', f'=IF($A{r}="","",$T{r}-$X{r}-N($K{r}))', fmt='#,##0;-#,##0;0')
    auto_cell(ws, f'V{r}', f'=IF($A{r}="","",IF(N($C{r})=0,"",$S{r}*12/$C{r}))', fmt=PCT)
TR = PRO1 + 2
setc(ws, f'A{TR}', '合計', font=f_bold, fill=fill_band)
for col in 'CEKPQSTUWXZ':
    setc(ws, f'{col}{TR}', f'=SUM({col}{PRO0}:{col}{PRO1})', font=f_bold, fmt=YEN if col not in 'QE' else '0', fill=fill_band)
setc(ws, f'N{TR}', f'=SUM(N{PRO0}:N{PRO1})', font=f_bold, fmt=YEN, fill=fill_band)
setc(ws, f'O{TR}', f'=SUM(O{PRO0}:O{PRO1})', font=f_bold, fmt=YEN, fill=fill_band)
dv_list(ws, '"小規模,一般,なし"', f'I{PRO0}:I{PRO1}')
dv = DataValidation(type='custom', formula1=f'COUNTIF($A${PRO0}:$A${PRO1},A{PRO0})=1', allow_blank=True,
                    errorTitle='物件名が重複しています', error='同じ物件名は1行だけにしてください。', showErrorMessage=True)
ws.add_data_validation(dv); dv.add(f'A{PRO0}:A{PRO1}')
cf(ws, f'U{PRO0}:U{PRO1}', f'AND($U{PRO0}<>"",$U{PRO0}<0)', fill_red)
cf(ws, f'Y{PRO0}:Y{PRO1}', f'$Y{PRO0}="入力値"', fill_grn)
ws.freeze_panes = 'B5'
PRO_TOT = lambda col: "'%s'!$%s$%d" % (SH['pro'], col, TR)

# ============================================================ 入金記録
ws = wb.create_sheet(SH['pay'])
title(ws, '入金記録（入金があるたびに1行追加。日付順に入力）',
      '「対象の年・月」は、その入金が何月分の家賃かを入力します。保証会社からの入金は、入居者ごとに1行ずつ割り当て、入金元を「保証会社」にします。')
header(ws, 4, ['入金日', '号室', '対象年', '対象月', '金額', '入金元', '保証会社名', '備考', '請求額（自動）', '計上額（自動）', '預り（自動）', '台帳の行（自動）'],
       [12, 8, 8, 8, 12, 10, 18, 24, 14, 14, 12, 12])
pay = [
 (D(2026,10,26), '101', 2026, 10, 75000, '入居者', '', ''),
 (D(2026,10,27), '102', 2026, 10, 40000, '入居者', '', '一部入金（サンプル）'),
 (D(2026,10,31), '103', 2026, 10, 70000, '保証会社', 'サンプル保証会社', '保証会社の振込を割り当て'),
 (D(2026,10,31), '104', 2026, 10, 70000, '保証会社', 'サンプル保証会社', '保証会社の振込を割り当て'),
]
ADJ_Y, ADJ_M, ADJ_R, ADJ_V = rng('adj', 'A', ADJ0, ADJ1), rng('adj', 'B', ADJ0, ADJ1), rng('adj', 'C', ADJ0, ADJ1), rng('adj', 'D', ADJ0, ADJ1)
for i in range(PAY1 - PAY0 + 1):
    r = PAY0 + i
    rowv = pay[i] if i < len(pay) else [None] * 8
    for j, v in enumerate(rowv):
        col = 'ABCDEFGH'[j]
        fmt = DATE if col == 'A' else (YEN if col == 'E' else None)
        c = input_cell(ws, f'{col}{r}', v, fmt=fmt)
        if col == 'B': c.number_format = '@'
    L = f'$L{r}'
    adj_if = f'COUNTIFS({ADJ_Y},$C{r},{ADJ_M},$D{r},{ADJ_R},$B{r})>0'
    adj_sum = f'SUMIFS({ADJ_V},{ADJ_Y},$C{r},{ADJ_M},$D{r},{ADJ_R},$B{r})'
    rev = f"INDEX({rng('led','Q',LED0,LED1)},{L})"
    base = (f"IF(AND({rev}<>\"\",DATE($C{r},$D{r},1)<{rev}),INDEX({rng('led','P',LED0,LED1)},{L}),INDEX({rng('led','F',LED0,LED1)},{L}))"
            f"+INDEX({rng('led','G',LED0,LED1)},{L})")
    auto_cell(ws, f'I{r}', f'=IF(OR($B{r}="",$C{r}="",$D{r}="",{L}=""),"",IF({adj_if},{adj_sum},{base}))', fmt=YEN)
    auto_cell(ws, f'J{r}', f'=IF(OR($I{r}="",$E{r}=""),"",MIN($E{r},MAX(0,$I{r}-SUMIFS($E$4:$E{r-1},$B$4:$B{r-1},$B{r},$C$4:$C{r-1},$C{r},$D$4:$D{r-1},$D{r}))))', fmt=YEN)
    auto_cell(ws, f'K{r}', f'=IF($J{r}="","",$E{r}-$J{r})', fmt=YEN)
    auto_cell(ws, f'L{r}', f'=IF($B{r}="","",IFERROR(MATCH($B{r},{ROOMS},0),""))', fmt='0')
ws.freeze_panes = 'A5'
dv_list(ws, room_src(ws), f'B{PAY0}:B{PAY1}')
dv_list(ws, '"入居者,保証会社"', f'F{PAY0}:F{PAY1}')
dv_list(ws, '"1,2,3,4,5,6,7,8,9,10,11,12"', f'D{PAY0}:D{PAY1}')

# ============================================================ 請求調整
ws = wb.create_sheet(SH['adj'])
title(ws, '請求調整（その月だけ請求額を変えるとき。任意）',
      '退去月に日割り相当の額にする場合や、家賃を2回以上改定した場合の過去の月に使います。入力した月は、台帳の家賃ではなく、ここの請求額（共益費込み）で判定します。')
header(ws, 4, ['対象年', '対象月', '号室', '請求額（共益費込み）', '理由'], [10, 10, 10, 20, 40])
adj = [(2026, 11, '104', 35000, '退去月（日割り相当）サンプル')]
for i in range(ADJ1 - ADJ0 + 1):
    r = ADJ0 + i
    rowv = adj[i] if i < len(adj) else [None] * 5
    for j, v in enumerate(rowv):
        col = 'ABCDE'[j]
        c = input_cell(ws, f'{col}{r}', v, fmt=(YEN if col == 'D' else None))
        if col == 'C': c.number_format = '@'
dv_list(ws, room_src(ws), f'C{ADJ0}:C{ADJ1}')
dv_list(ws, '"1,2,3,4,5,6,7,8,9,10,11,12"', f'B{ADJ0}:B{ADJ1}')
ws.freeze_panes = 'A5'

# ============================================================ 入金チェック
ws = wb.create_sheet(SH['chk'])
ws['A1'].value = '入金チェック'; ws['A1'].font = f_title
setc(ws, 'A2', '対象年', font=f_bold); input_cell(ws, 'B2', 2026, fmt='0')
setc(ws, 'A3', '対象月', font=f_bold); input_cell(ws, 'B3', 10, fmt='0')
setc(ws, 'A4', '基準日', font=f_bold); auto_cell(ws, 'B4', '=' + TODAY, fmt=DATE)
dv_list(ws, '"1,2,3,4,5,6,7,8,9,10,11,12"', 'B3', allow_blank=False)
G_RNG = f'$G${CHK0}:$G${CHK1}'; K_RNG = f'$K${CHK0}:$K${CHK1}'
setc(ws, 'D2', '入金済', font=f_bold); auto_cell(ws, 'E2', f'=COUNTIF({G_RNG},"入金済")', fmt='0')
setc(ws, 'D3', '黄', font=f_bold); auto_cell(ws, 'E3', f'=COUNTIF({K_RNG},"黄")', fmt='0')
setc(ws, 'D4', '赤', font=f_bold); auto_cell(ws, 'E4', f'=COUNTIF({K_RNG},"赤")', fmt='0')
setc(ws, 'G2', '請求額の合計', font=f_bold); auto_cell(ws, 'H2', f'=SUM($C${CHK0}:$C${CHK1})', fmt=YEN)
setc(ws, 'G3', '入金額の合計', font=f_bold); auto_cell(ws, 'H3', f'=SUM($E${CHK0}:$E${CHK1})', fmt=YEN)
setc(ws, 'G4', '不足額の合計', font=f_bold); auto_cell(ws, 'H4', f'=SUM($H${CHK0}:$H${CHK1})', fmt=YEN)
header(ws, 6, ['号室', '入居者名', '請求額', '支払期日', '入金額', '最終入金日', '状態', '不足額', '超過額', '超過日数', '滞納', '保証の型', '事故報告の期限', '期限までの残り'],
       [8, 14, 12, 12, 12, 12, 11, 12, 12, 10, 8, 12, 14, 16])
Y, M = '$B$2', '$B$3'
PB, PC, PD, PE, PA = rng('pay', 'B', PAY0, PAY1), rng('pay', 'C', PAY0, PAY1), rng('pay', 'D', PAY0, PAY1), rng('pay', 'E', PAY0, PAY1), rng('pay', 'A', PAY0, PAY1)
for r in range(CHK0, CHK1 + 1):
    k = r - CHK0 + LED0
    auto_cell(ws, f'A{r}', f"=IF(AND('台帳'!$A{k}<>\"\",'台帳'!$O{k}=\"入居中\"),'台帳'!$A{k},\"\")", align=CENTER)
    auto_cell(ws, f'B{r}', f"=IF($A{r}=\"\",\"\",'台帳'!B{k})")
    base = f"IF(AND('台帳'!Q{k}<>\"\",DATE({Y},{M},1)<'台帳'!Q{k}),'台帳'!P{k},'台帳'!F{k})+'台帳'!G{k}"
    auto_cell(ws, f'C{r}', (f"=IF($A{r}=\"\",\"\",IF(COUNTIFS({ADJ_Y},{Y},{ADJ_M},{M},{ADJ_R},$A{r})>0,"
                            f"SUMIFS({ADJ_V},{ADJ_Y},{Y},{ADJ_M},{M},{ADJ_R},$A{r}),{base}))"), fmt=YEN)
    auto_cell(ws, f'D{r}', f"=IF($A{r}=\"\",\"\",IF('台帳'!H{k}=\"月末\",DATE({Y},{M}+1,0),DATE({Y},{M},MIN('台帳'!H{k},DAY(DATE({Y},{M}+1,0))))))", fmt=DATE)
    auto_cell(ws, f'E{r}', f"=IF($A{r}=\"\",\"\",SUMIFS({PE},{PB},$A{r},{PC},{Y},{PD},{M}))", fmt=YEN)
    mx = f"SUMPRODUCT(MAX(({PB}=$A{r})*({PC}={Y})*({PD}={M})*{PA}))"
    auto_cell(ws, f'F{r}', f"=IF($A{r}=\"\",\"\",IF({mx}=0,\"\",{mx}))", fmt=DATE)
    auto_cell(ws, f'G{r}', f"=IF($A{r}=\"\",\"\",IF(E{r}=C{r},\"入金済\",IF(E{r}>C{r},\"過入金\",IF(E{r}>0,\"一部入金\",IF({TODAY}<=D{r},\"期日前\",\"未入金\")))))", align=CENTER)
    auto_cell(ws, f'H{r}', f"=IF($A{r}=\"\",\"\",MAX(0,C{r}-E{r}))", fmt=YEN0)
    auto_cell(ws, f'I{r}', f"=IF($A{r}=\"\",\"\",MAX(0,E{r}-C{r}))", fmt=YEN0)
    auto_cell(ws, f'J{r}', f"=IF($A{r}=\"\",\"\",IF(OR(G{r}=\"入金済\",G{r}=\"過入金\"),0,MAX(0,{TODAY}-D{r})))", fmt='0')
    auto_cell(ws, f'K{r}', f"=IF($A{r}=\"\",\"\",IF(J{r}>{S['r']},\"赤\",IF(AND(J{r}>={S['y']},J{r}>0),\"黄\",\"\")))", align=CENTER)
    auto_cell(ws, f'L{r}', f"=IF($A{r}=\"\",\"\",IF('台帳'!L{k}=\"\",\"\",'台帳'!L{k}))")
    days = f"INDEX({GU_DAYS},MATCH('台帳'!K{k},{GU_NAMES},0))"
    auto_cell(ws, f'M{r}', (f"=IF($A{r}=\"\",\"\",IF(AND('台帳'!L{k}=\"一般保証型\",K{r}=\"赤\"),"
                            f"IFERROR(IF({days}=\"\",\"期限日数未設定\",D{r}+{days}),\"期限日数未設定\"),\"\"))"), fmt=DATE)
    auto_cell(ws, f'N{r}', f"=IF(M{r}=\"\",\"\",IF(ISNUMBER(M{r}),IF(M{r}-{TODAY}<0,\"報告期限超過\",\"残り\"&(M{r}-{TODAY})&\"日\"),M{r}))")
ws.freeze_panes = 'C7'
cf(ws, f'A{CHK0}:N{CHK1}', f'$K{CHK0}="赤"', fill_red)
cf(ws, f'A{CHK0}:N{CHK1}', f'$K{CHK0}="黄"', fill_yel)
cf(ws, f'G{CHK0}:G{CHK1}', f'$G{CHK0}="入金済"', fill_grn)

# ============================================================ 保証会社入金
ws = wb.create_sheet(SH['gua'])
title(ws, '保証会社入金（支払委託型）',
      '保証会社からの1回の振込を入力すると、入金記録で入居者へ割り当てた合計との差を確認できます。入金記録に割り当てを入力するとき、入金日と保証会社名を揃えてください。')
header(ws, 4, ['入金日', '保証会社名', '振込額', '割り当て合計（自動）', '差額（自動）', '判定'], [12, 22, 14, 18, 14, 10])
PF, PG = rng('pay', 'F', PAY0, PAY1), rng('pay', 'G', PAY0, PAY1)
for i in range(GUA1 - GUA0 + 1):
    r = GUA0 + i
    vals = (D(2026,10,31), 'サンプル保証会社', 140000) if i == 0 else (None, None, None)
    input_cell(ws, f'A{r}', vals[0], fmt=DATE); input_cell(ws, f'B{r}', vals[1]); input_cell(ws, f'C{r}', vals[2], fmt=YEN)
    auto_cell(ws, f'D{r}', f"=IF(OR($A{r}=\"\",$C{r}=\"\"),\"\",SUMIFS({PE},{PA},$A{r},{PG},$B{r},{PF},\"保証会社\"))", fmt=YEN)
    auto_cell(ws, f'E{r}', f"=IF($D{r}=\"\",\"\",$C{r}-$D{r})", fmt=YEN)
    auto_cell(ws, f'F{r}', f"=IF($D{r}=\"\",\"\",IF($E{r}=0,\"OK\",\"差異\"))", align=CENTER)
cf(ws, f'A{GUA0}:F{GUA1}', f'$F{GUA0}="差異"', fill_red)
ws.freeze_panes = 'A5'

# ============================================================ 期限管理
ws = wb.create_sheet(SH['due'])
ws['A1'].value = '期限管理（入居中の契約）'; ws['A1'].font = f_title
setc(ws, 'A2', '赤・期限超過', font=f_bold); auto_cell(ws, 'B2', None)
setc(ws, 'D2', '黄', font=f_bold); auto_cell(ws, 'E2', None)
ws['A3'].value = '残り日数が小さいほど期限が近いです。黄（期限が近い）と赤（期限が迫っている、または超過）の色で、対応が必要な契約が分かります。'; ws['A3'].font = f_note
header(ws, 4, ['号室', '入居者名', '契約更新日', '残り日数', '判定', '火災保険の期限', '残り日数', '判定', '保証会社の契約期限', '残り日数', '判定', '最短の残り日数'],
       [10, 14, 12, 10, 10, 14, 10, 10, 16, 10, 10, 14])
DUE0, DUE1 = 5, 5 + (LED1 - LED0)
for r in range(DUE0, DUE1 + 1):
    k = r - DUE0 + LED0
    auto_cell(ws, f'A{r}', f"=IF(AND('台帳'!$A{k}<>\"\",'台帳'!$O{k}=\"入居中\"),'台帳'!$A{k},\"\")", align=CENTER)
    auto_cell(ws, f'B{r}', f"=IF($A{r}=\"\",\"\",'台帳'!B{k})")
    for dcol, ncol, jcol, src in (('C', 'D', 'E', 'J'), ('F', 'G', 'H', 'M'), ('I', 'J', 'K', 'N')):
        auto_cell(ws, f'{dcol}{r}', f"=IF($A{r}=\"\",\"\",IF('台帳'!{src}{k}=\"\",\"\",'台帳'!{src}{k}))", fmt=DATE)
        auto_cell(ws, f'{ncol}{r}', f"=IF({dcol}{r}=\"\",\"\",{dcol}{r}-{TODAY})", fmt='0')
        auto_cell(ws, f'{jcol}{r}', f"=IF({ncol}{r}=\"\",\"\",IF({ncol}{r}<0,\"期限超過\",IF({ncol}{r}<={S['dr']},\"赤\",IF({ncol}{r}<={S['dy']},\"黄\",\"\"))))", align=CENTER)
    auto_cell(ws, f'L{r}', f"=IF(COUNT(D{r},G{r},J{r})=0,\"\",MIN(D{r},G{r},J{r}))", fmt='0')
ws['B2'].value = '=' + '+'.join(f'COUNTIF($%s$%d:$%s$%d,"%s")' % (c, DUE0, c, DUE1, v) for c in 'EHK' for v in ('赤', '期限超過'))
ws['E2'].value = '=' + '+'.join(f'COUNTIF($%s$%d:$%s$%d,"黄")' % (c, DUE0, c, DUE1) for c in 'EHK')
for a, b_, c_ in (('C', 'D', 'E'), ('F', 'G', 'H'), ('I', 'J', 'K')):
    cf(ws, f'{a}{DUE0}:{c_}{DUE1}', f'OR(${c_}{DUE0}="赤",${c_}{DUE0}="期限超過")', fill_red)
    cf(ws, f'{a}{DUE0}:{c_}{DUE1}', f'${c_}{DUE0}="黄"', fill_yel)
ws.freeze_panes = 'C5'

# ============================================================ 敷金台帳
ws = wb.create_sheet(SH['dep'])
ws['A1'].value = '敷金台帳（敷金・保証金・預り金）'; ws['A1'].font = f_title
setc(ws, 'A2', '預り中の合計', font=f_bold); auto_cell(ws, 'C2', f'=SUMIFS($D${DEP0}:$D${DEP1},$L${DEP0}:$L${DEP1},"<>精算済")', fmt=YEN)
setc(ws, 'E2', '精算遅れの件数', font=f_bold); auto_cell(ws, 'G2', f'=COUNTIF($N${DEP0}:$N${DEP1},"精算遅れ")', fmt='0')
ws['A3'].value = '退去日と控除額を入力すると、返還額と追加請求額が出ます。状況を「精算済」にすると、遅れの判定から外れます。'; ws['A3'].font = f_note
header(ws, 4, ['号室', '入居者名（自動）', '区分', '預かり額', '預かり日', '退去日', '原状回復費', 'その他の控除', '控除合計（自動）', '返還額（自動）', '追加請求額（自動）', '状況', '精算の目安日（自動）', '判定（自動）'],
       [8, 14, 12, 12, 12, 12, 12, 12, 14, 14, 14, 10, 14, 12])
deps = [
 ('101', '敷金', 70000, D(2025,4,1), None, None, None, '預り中'),
 ('102', '敷金', 65000, D(2024,9,1), None, None, None, '預り中'),
 ('106', '敷金', 60000, D(2023,5,1), D(2025,4,30), 30000, 0, '精算済'),
 ('107', '敷金', 65000, D(2024,3,1), D(2026,10,31), 28000, 0, '預り中'),
]
for i in range(DEP1 - DEP0 + 1):
    r = DEP0 + i
    v = deps[i] if i < len(deps) else [None] * 8
    c = input_cell(ws, f'A{r}', v[0]); c.number_format = '@'
    auto_cell(ws, f'B{r}', f"=IF($A{r}=\"\",\"\",IFERROR(INDEX({rng('led','B',LED0,LED1)},MATCH($A{r},{ROOMS},0)),\"\"))")
    input_cell(ws, f'C{r}', v[1]); input_cell(ws, f'D{r}', v[2], fmt=YEN); input_cell(ws, f'E{r}', v[3], fmt=DATE)
    input_cell(ws, f'F{r}', v[4], fmt=DATE); input_cell(ws, f'G{r}', v[5], fmt=YEN); input_cell(ws, f'H{r}', v[6], fmt=YEN)
    auto_cell(ws, f'I{r}', f'=IF($D{r}="","",N($G{r})+N($H{r}))', fmt=YEN)
    auto_cell(ws, f'J{r}', f'=IF($D{r}="","",MAX(0,$D{r}-$I{r}))', fmt=YEN)
    auto_cell(ws, f'K{r}', f'=IF($D{r}="","",MAX(0,$I{r}-$D{r}))', fmt=YEN0)
    input_cell(ws, f'L{r}', v[7])
    auto_cell(ws, f'M{r}', f'=IF(OR($D{r}="",$F{r}=""),"",$F{r}+{S["dt"]})', fmt=DATE)
    auto_cell(ws, f'N{r}', f'=IF($D{r}="","",IF($L{r}="精算済","精算済",IF($F{r}="","預り中",IF({TODAY}>$M{r},"精算遅れ","精算期間内"))))', align=CENTER)
dv_list(ws, room_src(ws), f'A{DEP0}:A{DEP1}')
dv_list(ws, '"敷金,保証金,その他預り金"', f'C{DEP0}:C{DEP1}')
dv_list(ws, '"預り中,精算済"', f'L{DEP0}:L{DEP1}')
cf(ws, f'A{DEP0}:N{DEP1}', f'$N{DEP0}="精算遅れ"', fill_red)
cf(ws, f'N{DEP0}:N{DEP1}', f'$N{DEP0}="精算期間内"', fill_yel)
cf(ws, f'N{DEP0}:N{DEP1}', f'$N{DEP0}="精算済"', fill_grn)
ws.freeze_panes = 'C5'

# ============================================================ 空室・募集
ws = wb.create_sheet(SH['vac'])
ws['A1'].value = '空室・募集'; ws['A1'].font = f_title
setc(ws, 'A2', '空室数', font=f_bold); auto_cell(ws, 'B2', f'=SUMPRODUCT(--($A${VAC0}:$A${VAC1}<>""))', fmt='0')
setc(ws, 'C2', '空室率', font=f_bold); auto_cell(ws, 'D2', f'=IF({PRO_TOT("E")}=0,"",B2/{PRO_TOT("E")})', fmt=PCT)
setc(ws, 'F2', '空室損失の合計（概算）', font=f_bold); auto_cell(ws, 'G2', f'=SUM($L${VAC0}:$L${VAC1})', fmt=YEN)
ws['A3'].value = '空いている部屋だけを入力します。入居が決まったら、その行を空欄にして、台帳に契約を入力します。空室損失は、募集家賃×空室日数÷30の概算です。'; ws['A3'].font = f_note
header(ws, 4, ['号室', '状況', '空室開始日', '募集開始日', '募集家賃', '募集媒体', '内見の件数', '申込の件数', '備考', '空室日数（自動）', '判定（自動）', '空室損失・概算（自動）'],
       [8, 12, 12, 12, 12, 22, 10, 10, 24, 12, 10, 16])
vacs = [('201', '募集中', D(2026,9,20), D(2026,10,1), 68000, 'ポータルサイトA（架空）', 3, 1, ''),
        ('202', '空室', D(2026,11,1), None, 72000, '', 0, 0, '')]
for i in range(VAC1 - VAC0 + 1):
    r = VAC0 + i
    v = vacs[i] if i < len(vacs) else [None] * 9
    c = input_cell(ws, f'A{r}', v[0]); c.number_format = '@'
    input_cell(ws, f'B{r}', v[1]); input_cell(ws, f'C{r}', v[2], fmt=DATE); input_cell(ws, f'D{r}', v[3], fmt=DATE)
    input_cell(ws, f'E{r}', v[4], fmt=YEN); input_cell(ws, f'F{r}', v[5]); input_cell(ws, f'G{r}', v[6], fmt='0')
    input_cell(ws, f'H{r}', v[7], fmt='0'); input_cell(ws, f'I{r}', v[8])
    auto_cell(ws, f'J{r}', f'=IF(OR($A{r}="",$C{r}=""),"",MAX(0,{TODAY}-$C{r}))', fmt='0')
    auto_cell(ws, f'K{r}', f'=IF($J{r}="","",IF($J{r}>={S["vr"]},"赤",IF($J{r}>={S["vy"]},"黄","")))', align=CENTER)
    auto_cell(ws, f'L{r}', f'=IF(OR($J{r}="",$E{r}=""),"",ROUND($E{r}*$J{r}/30,0))', fmt=YEN)
dv_list(ws, '"空室,募集中,申込あり"', f'B{VAC0}:B{VAC1}')
dv_list(ws, room_src(ws), f'A{VAC0}:A{VAC1}')
cf(ws, f'A{VAC0}:L{VAC1}', f'$K{VAC0}="赤"', fill_red)
cf(ws, f'A{VAC0}:L{VAC1}', f'$K{VAC0}="黄"', fill_yel)
ws.freeze_panes = 'B5'

# ============================================================ 修繕履歴
ws = wb.create_sheet(SH['rep'])
ws['A1'].value = '修繕・設備の履歴'; ws['A1'].font = f_title
setc(ws, 'A2', '目安超過', font=f_bold); auto_cell(ws, 'B2', f'=COUNTIF($L${REP0}:$L${REP1},"目安超過")', fmt='0')
setc(ws, 'C2', '赤', font=f_bold); auto_cell(ws, 'D2', f'=COUNTIF($L${REP0}:$L${REP1},"赤")', fmt='0')
ws['A3'].value = '工事や交換をするたびに1行追加します。次回の目安（年）を入れると、交換時期が近づいたときに色で知らせます。「年間収支に含める」の扱い（修繕費か資本的支出か）は、税理士に確認してください。'; ws['A3'].font = f_note
header(ws, 4, ['実施日', '号室（または共用）', '区分', '設備・箇所', '内容', '業者', '費用', '年間収支に含める', '次回の目安（年）', '次回の目安日（自動）', '残り日数（自動）', '判定（自動）'],
       [12, 14, 10, 16, 24, 18, 12, 14, 12, 14, 12, 12])
reps = [
 (D(2026,3,15), '共用', '点検', '消防設備', '年次点検', 'サンプル消防設備', 30000, '含める', 1),
 (D(2024,8,20), '101', '交換', '給湯器', '機器交換', 'サンプル設備', 180000, '含めない', 10),
 (D(2016,6,10), '102', '交換', 'エアコン', '機器交換', 'サンプル設備', 90000, '含めない', 10),
 (D(2026,5,10), '103', '修繕', '水栓', '水漏れ修理', 'サンプル設備', 20000, '含める', None),
 (D(2026,6,12), '102', '交換', 'エアコン', '機器交換（更新）', 'サンプル設備', 95000, '含めない', 10),
]
RB, RD, RA = rng('rep', 'B', REP0, REP1), rng('rep', 'D', REP0, REP1), rng('rep', 'A', REP0, REP1)
for i in range(REP1 - REP0 + 1):
    r = REP0 + i
    v = reps[i] if i < len(reps) else [None] * 9
    input_cell(ws, f'A{r}', v[0], fmt=DATE); c = input_cell(ws, f'B{r}', v[1]); c.number_format = '@'
    input_cell(ws, f'C{r}', v[2]); input_cell(ws, f'D{r}', v[3]); input_cell(ws, f'E{r}', v[4]); input_cell(ws, f'F{r}', v[5])
    input_cell(ws, f'G{r}', v[6], fmt=YEN); input_cell(ws, f'H{r}', v[7]); input_cell(ws, f'I{r}', v[8], fmt='0')
    auto_cell(ws, f'J{r}', f'=IF(OR($A{r}="",$I{r}=""),"",EDATE($A{r},12*$I{r}))', fmt=DATE)
    auto_cell(ws, f'K{r}', f'=IF($J{r}="","",$J{r}-{TODAY})', fmt='0')
    auto_cell(ws, f'L{r}', (f'=IF($J{r}="","",IF(COUNTIFS({RB},$B{r},{RD},$D{r},{RA},">"&$A{r})>0,"更新済",'
                            f'IF($K{r}<0,"目安超過",IF($K{r}<={S["er"]},"赤",IF($K{r}<={S["ey"]},"黄","")))))'), align=CENTER)
dv_list(ws, '"修繕,交換,点検,その他"', f'C{REP0}:C{REP1}')
_dv = DataValidation(type='list', formula1=room_src(ws), allow_blank=True, errorStyle='warning', errorTitle='台帳にない号室です', error='共用部などの場合は、このまま入力できます。', showErrorMessage=True)
ws.add_data_validation(_dv); _dv.add(f'B{REP0}:B{REP1}')
dv_list(ws, '"含める,含めない"', f'H{REP0}:H{REP1}')
cf(ws, f'A{REP0}:L{REP1}', f'OR($L{REP0}="赤",$L{REP0}="目安超過")', fill_red)
cf(ws, f'A{REP0}:L{REP1}', f'$L{REP0}="黄"', fill_yel)
ws.freeze_panes = 'A5'

# ============================================================ 税カレンダー
ws = wb.create_sheet(SH['tax'])
ws['A1'].value = '税カレンダー（固定資産税・都市計画税など）'; ws['A1'].font = f_title
setc(ws, 'A2', '今後の納付予定額', font=f_bold)
auto_cell(ws, 'C2', f'=SUMIFS($F${TAX0}:$F${TAX1},$H${TAX0}:$H${TAX1},"未納",$E${TAX0}:$E${TAX1},">="&{TODAY},$E${TAX0}:$E${TAX1},"<="&({TODAY}+{S["tw"]}))', fmt=YEN)
setc(ws, 'E2', '納期限を過ぎた未納額', font=f_bold)
auto_cell(ws, 'G2', f'=SUMIFS($F${TAX0}:$F${TAX1},$H${TAX0}:$H${TAX1},"未納",$E${TAX0}:$E${TAX1},"<"&{TODAY})', fmt=YEN)
ws['A3'].value = '納税通知書の納期限と税額を入力し、納付したら納付日を入力します。納期は自治体によって異なります。サンプルは架空の日付・金額です。'; ws['A3'].font = f_note
header(ws, 4, ['年度', '税目', '対象', '期', '納期限', '納付額', '納付日', '状況（自動）', '残り日数（自動）', '判定（自動）'],
       [8, 14, 18, 8, 12, 12, 12, 12, 14, 12])
taxes = [
 (2026, '固定資産税', '本館サンプル', '第1期', D(2026,4,30), 38000, D(2026,4,28)),
 (2026, '固定資産税', '本館サンプル', '第2期', D(2026,7,31), 37000, D(2026,7,30)),
 (2026, '固定資産税', '本館サンプル', '第3期', D(2026,12,28), 37000, None),
 (2026, '固定資産税', '本館サンプル', '第4期', D(2027,2,28), 37000, None),
]
for i in range(TAX1 - TAX0 + 1):
    r = TAX0 + i
    v = taxes[i] if i < len(taxes) else [None] * 7
    input_cell(ws, f'A{r}', v[0], fmt='0'); input_cell(ws, f'B{r}', v[1]); input_cell(ws, f'C{r}', v[2]); input_cell(ws, f'D{r}', v[3])
    input_cell(ws, f'E{r}', v[4], fmt=DATE); input_cell(ws, f'F{r}', v[5], fmt=YEN); input_cell(ws, f'G{r}', v[6], fmt=DATE)
    auto_cell(ws, f'H{r}', f'=IF(OR($E{r}="",$F{r}=""),"",IF($G{r}<>"","納付済","未納"))', align=CENTER)
    auto_cell(ws, f'I{r}', f'=IF($H{r}="未納",$E{r}-{TODAY},"")', fmt='0')
    auto_cell(ws, f'J{r}', f'=IF($H{r}="","",IF($H{r}="納付済","納付済",IF($I{r}<0,"納期限超過",IF($I{r}<={S["tr"]},"赤",IF($I{r}<={S["ty"]},"黄","")))))', align=CENTER)
dv_list(ws, '"固定資産税,都市計画税,その他"', f'B{TAX0}:B{TAX1}')
dv_list(ws, '=' + PRO_NAMES, f'C{TAX0}:C{TAX1}')
cf(ws, f'A{TAX0}:J{TAX1}', f'OR($J{TAX0}="赤",$J{TAX0}="納期限超過")', fill_red)
cf(ws, f'A{TAX0}:J{TAX1}', f'$J{TAX0}="黄"', fill_yel)
cf(ws, f'J{TAX0}:J{TAX1}', f'$J{TAX0}="納付済"', fill_grn)
ws.freeze_panes = 'A5'

# ============================================================ 家賃改定試算
ws = wb.create_sheet(SH['rev'])
ws['A1'].value = '家賃改定試算'; ws['A1'].font = f_title
setc(ws, 'A2', '年間の増減額の合計', font=f_bold); auto_cell(ws, 'C2', f'=SUM($I${REV0}:$I${REV1})', fmt=YEN)
setc(ws, 'E2', '期待される年間の影響額の合計', font=f_bold); auto_cell(ws, 'H2', f'=SUM($J${REV0}:$J${REV1})', fmt=YEN)
ws['A3'].value = '改定を検討する号室を入力します。値上げの期待値は、設定シートの想定退去率と空室想定日数を使った概算です。結果を保証するものではありません。'; ws['A3'].font = f_note
header(ws, 4, ['号室', '入居者名（自動）', '現行の家賃（自動）', '周辺相場', '相場との差（自動）', '改定案の家賃', '増減額（自動）', '改定率（自動）', '年間の増減額（自動）', '期待される年間の影響額（自動）'],
       [8, 14, 14, 12, 14, 14, 12, 12, 16, 20])
revs = [('101', 74000, 72000), ('102', 63000, 65000)]
for i in range(REV1 - REV0 + 1):
    r = REV0 + i
    v = revs[i] if i < len(revs) else [None] * 3
    c = input_cell(ws, f'A{r}', v[0]); c.number_format = '@'
    auto_cell(ws, f'B{r}', f"=IF($A{r}=\"\",\"\",IFERROR(INDEX({rng('led','B',LED0,LED1)},MATCH($A{r},{ROOMS},0)),\"\"))")
    auto_cell(ws, f'C{r}', f"=IF($A{r}=\"\",\"\",IFERROR(INDEX({rng('led','F',LED0,LED1)},MATCH($A{r},{ROOMS},0)),\"\"))", fmt=YEN)
    input_cell(ws, f'D{r}', v[1], fmt=YEN)
    auto_cell(ws, f'E{r}', f'=IF(OR($C{r}="",$D{r}=""),"",($C{r}-$D{r})/$D{r})', fmt=PCT)
    input_cell(ws, f'F{r}', v[2], fmt=YEN)
    auto_cell(ws, f'G{r}', f'=IF(OR($C{r}="",$F{r}=""),"",$F{r}-$C{r})', fmt='#,##0;-#,##0;0')
    auto_cell(ws, f'H{r}', f'=IF($G{r}="","",$G{r}/$C{r})', fmt=PCT)
    auto_cell(ws, f'I{r}', f'=IF($G{r}="","",$G{r}*12)', fmt=YEN)
    auto_cell(ws, f'J{r}', f'=IF($G{r}="","",IF($G{r}>0,$I{r}*(1-{S["churn"]})-{S["churn"]}*$C{r}*{S["vdays"]}/30,$I{r}))', fmt=YEN)
dv_list(ws, room_src(ws), f'A{REV0}:A{REV1}')
ws.freeze_panes = 'A5'

# ============================================================ 年間収支
ws = wb.create_sheet(SH['ann'])
ws['A1'].value = '年間収支（参考値）'; ws['A1'].font = f_title
setc(ws, 'A2', '対象年', font=f_bold); input_cell(ws, 'B2', 2026, fmt='0')
ws['A3'].value = '家賃収入は入金日ベースで、請求額を超えた入金は「預り」として収入に含めません。修繕費は修繕履歴で「含める」にしたもの、税金は税カレンダーで納付済のものです。確定申告の内容は、税務署や税理士に確認してください。'; ws['A3'].font = f_note
header(ws, 5, ['月', '家賃収入', '預り（参考）', '修繕費（自動）', '税金（自動）', 'その他の経費（手入力）', '経費合計', '収支'], [8, 16, 16, 16, 16, 20, 16, 16])
exp_sample = {12: 20000}
JR, KR, AR = rng('pay', 'J', PAY0, PAY1), rng('pay', 'K', PAY0, PAY1), rng('pay', 'A', PAY0, PAY1)
for m in range(1, 13):
    r = 5 + m
    start, end = f'DATE($B$2,$A{r},1)', f'EOMONTH(DATE($B$2,$A{r},1),0)'
    setc(ws, f'A{r}', m, fmt='0"月"', align=CENTER)
    auto_cell(ws, f'B{r}', f'=SUMIFS({JR},{AR},">="&{start},{AR},"<="&{end})', fmt=YEN)
    auto_cell(ws, f'C{r}', f'=SUMIFS({KR},{AR},">="&{start},{AR},"<="&{end})', fmt=YEN)
    auto_cell(ws, f'D{r}', f"=SUMIFS({rng('rep','G',REP0,REP1)},{rng('rep','A',REP0,REP1)},\">=\"&{start},{rng('rep','A',REP0,REP1)},\"<=\"&{end},{rng('rep','H',REP0,REP1)},\"含める\")", fmt=YEN)
    auto_cell(ws, f'E{r}', f"=SUMIFS({rng('tax','F',TAX0,TAX1)},{rng('tax','G',TAX0,TAX1)},\">=\"&{start},{rng('tax','G',TAX0,TAX1)},\"<=\"&{end},{rng('tax','H',TAX0,TAX1)},\"納付済\")", fmt=YEN)
    input_cell(ws, f'F{r}', exp_sample.get(m), fmt=YEN)
    auto_cell(ws, f'G{r}', f'=D{r}+E{r}+N(F{r})', fmt=YEN)
    auto_cell(ws, f'H{r}', f'=B{r}-G{r}', fmt=YEN)
setc(ws, 'A18', '合計', font=f_bold, align=CENTER)
for col in 'BCDEFGH':
    setc(ws, f'{col}18', f'=SUM({col}6:{col}17)', font=f_bold, fmt=YEN, fill=fill_band)
ws['A20'].value = 'その他の経費（12月の20,000円）は架空のサンプルです。'; ws['A20'].font = f_note

# ---- 予定年間収支（保有物件にもとづく目安）
ws['A22'].value = '予定年間収支（保有物件にもとづく目安）'; ws['A22'].font = Font(name=FONT, size=13, bold=True, color=GREEN)
ws['A23'].value = ('物件台帳の見込みから、対象年の予定を出します。税は、税カレンダーに登録（対象＝物件名）があればその額を納期限の月に入れ（入力値を優先）、登録がない物件は推定額を12等分で入れます。'
                   '設備は修繕履歴の次回の目安日の月に、前回の費用を目安として入れます。あくまで目安です。')
ws['A23'].font = f_note
header(ws, 25, ['月', '見込み家賃', '税（納付予定）', '設備の交換見込み', 'その他経費の見込み', '経費合計', '収支'])
TXE, TXF = rng('tax', 'E', TAX0, TAX1), rng('tax', 'F', TAX0, TAX1)
RPJ, RPG, RPL = rng('rep', 'J', REP0, REP1), rng('rep', 'G', REP0, REP1), rng('rep', 'L', REP0, REP1)
for m in range(1, 13):
    r = 25 + m
    start, end = f'DATE($B$2,$A{r},1)', f'EOMONTH(DATE($B$2,$A{r},1),0)'
    setc(ws, f'A{r}', m, fmt='0"月"', align=CENTER)
    auto_cell(ws, f'B{r}', f"=ROUND({PRO_TOT('T')}/12,0)", fmt=YEN)
    auto_cell(ws, f'C{r}', f'=SUMIFS({TXF},{TXE},">="&{start},{TXE},"<="&{end})+ROUND({PRO_TOT("Z")}/12,0)', fmt=YEN)
    auto_cell(ws, f'D{r}', f'=SUMIFS({RPG},{RPJ},">="&{start},{RPJ},"<="&{end},{RPL},"<>更新済")', fmt=YEN)
    auto_cell(ws, f'E{r}', f"=ROUND({PRO_TOT('K')}/12,0)", fmt=YEN)
    auto_cell(ws, f'F{r}', f'=C{r}+D{r}+E{r}', fmt=YEN)
    auto_cell(ws, f'G{r}', f'=B{r}-F{r}', fmt='#,##0;-#,##0;0')
setc(ws, 'A38', '合計', font=f_bold, align=CENTER)
for col in 'BCDEFG':
    setc(ws, f'{col}38', f'=SUM({col}26:{col}37)', font=f_bold, fmt='#,##0;-#,##0;0', fill=fill_band)
cf(ws, 'G26:G38', 'G26<0', fill_red)
ws['A40'].value = '物件別の予定年間収支'; ws['A40'].font = Font(name=FONT, size=11, bold=True, color=GREEN)
header(ws, 41, ['物件', '年間の見込み家賃', '採用する税', '税の根拠', 'その他の年間経費', '予定の年間収支', '表面利回り', '取得価格'])
for i in range(PRO1 - PRO0 + 1):
    r = 42 + i; k = PRO0 + i
    P = lambda c: f"'{SH['pro']}'!{c}{k}"
    auto_cell(ws, f'A{r}', f'=IF({P("A")}="","",{P("A")})')
    auto_cell(ws, f'B{r}', f'=IF($A{r}="","",{P("T")})', fmt=YEN)
    auto_cell(ws, f'C{r}', f'=IF($A{r}="","",{P("X")})', fmt=YEN)
    auto_cell(ws, f'D{r}', f'=IF($A{r}="","",{P("Y")})', align=CENTER)
    auto_cell(ws, f'E{r}', f'=IF($A{r}="","",N({P("K")}))', fmt=YEN)
    auto_cell(ws, f'F{r}', f'=IF($A{r}="","",{P("U")})', fmt='#,##0;-#,##0;0')
    auto_cell(ws, f'G{r}', f'=IF($A{r}="","",{P("V")})', fmt=PCT)
    auto_cell(ws, f'H{r}', f'=IF($A{r}="","",{P("C")})', fmt=YEN)
setc(ws, 'A62', '合計', font=f_bold, align=CENTER)
for col in 'BCEFH':
    setc(ws, f'{col}62', f'=SUM({col}42:{col}61)', font=f_bold, fmt='#,##0;-#,##0;0', fill=fill_band)
cf(ws, 'F42:F61', 'AND($F42<>"",$F42<0)', fill_red)
cf(ws, 'D42:D61', '$D42="入力値"', fill_grn)

# ============================================================ ダッシュボード
ws = ws_dash
ws.sheet_view.showGridLines = False
for col, w in zip('ABCDEF', (2, 30, 18, 3, 30, 18)):
    ws.column_dimensions[col].width = w
ws.merge_cells('B1:F1')
ws['B1'].value = 'ダッシュボード'; ws['B1'].font = Font(name=FONT, size=18, bold=True, color='FFFFFF'); ws['B1'].fill = fill_head
ws['B1'].alignment = Alignment(vertical='center', indent=1); ws.row_dimensions[1].height = 38
ws.merge_cells('B2:F2')
ws['B2'].value = '各シートの集計を1画面にまとめています。数値はすべて自動で出ます（入力は不要）。'; ws['B2'].font = f_note
def block(ws, top, left_col, heading, items):
    c0 = left_col; c1 = chr(ord(left_col) + 1)
    ws.merge_cells(f'{c0}{top}:{c1}{top}')
    h = ws[f'{c0}{top}']; h.value = heading; h.font = Font(name=FONT, size=11, bold=True, color=GREEN); h.fill = fill_band
    ws.row_dimensions[top].height = 24
    for i, (label, formula, fmt, strong) in enumerate(items, start=1):
        r = top + i
        setc(ws, f'{c0}{r}', label, font=f_base)
        setc(ws, f'{c1}{r}', formula, font=Font(name=FONT, size=12 if strong else 10, bold=strong), fmt=fmt, align=Alignment(horizontal='right'))
chk = SH['chk']
block(ws, 4, 'B', '今月の入金', [
    ('対象', f"='{chk}'!B2&\"年\"&'{chk}'!B3&\"月分\"", None, False),
    ('入金済（件）', f"='{chk}'!E2", '0', True),
    ('黄（件）', f"='{chk}'!E3", '0', True),
    ('赤（件）', f"='{chk}'!E4", '0', True),
    ('請求額の合計', f"='{chk}'!H2", YEN, False),
    ('入金額の合計', f"='{chk}'!H3", YEN, False),
    ('不足額の合計', f"='{chk}'!H4", YEN, True)])
block(ws, 4, 'E', '期限（契約・保険・保証会社）', [
    ('赤・期限超過（件）', f"='{SH['due']}'!B2", '0', True),
    ('黄（件）', f"='{SH['due']}'!E2", '0', True)])
block(ws, 8, 'E', '空室・募集', [
    ('空室数（室）', f"='{SH['vac']}'!B2", '0', True),
    ('空室率', f"='{SH['vac']}'!D2", PCT, True),
    ('空室損失の合計（概算）', f"='{SH['vac']}'!G2", YEN, False)])
block(ws, 13, 'B', '敷金・保証金', [
    ('預り中の合計', f"='{SH['dep']}'!C2", YEN, True),
    ('精算遅れ（件）', f"='{SH['dep']}'!G2", '0', True)])
block(ws, 13, 'E', '固定資産税などの納付', [
    ('今後の納付予定額', f"='{SH['tax']}'!C2", YEN, True),
    ('納期限を過ぎた未納額', f"='{SH['tax']}'!G2", YEN, True)])
block(ws, 17, 'B', '設備の交換目安', [
    ('目安超過（件）', f"='{SH['rep']}'!B2", '0', True),
    ('赤（件）', f"='{SH['rep']}'!D2", '0', True)])
block(ws, 17, 'E', '家賃改定の試算', [
    ('年間の増減額の合計', f"='{SH['rev']}'!C2", YEN, False),
    ('期待される年間の影響額', f"='{SH['rev']}'!H2", YEN, True)])
block(ws, 21, 'B', '年間収支（対象年）', [
    ('対象年', f"='{SH['ann']}'!B2&\"年\"", None, False),
    ('家賃収入', f"='{SH['ann']}'!B18", YEN, False),
    ('経費合計', f"='{SH['ann']}'!G18", YEN, False),
    ('収支', f"='{SH['ann']}'!H18", YEN, True)])
ws['E21'].value = None
block(ws, 27, 'B', '物件・予定収支', [
    ('物件の取得価格の合計', f"={PRO_TOT('C')}", YEN, False),
    ('推定の年間税額', f"={PRO_TOT('P')}", YEN, False),
    ('年間の予定収支（物件台帳の合計）', f"={PRO_TOT('U')}", YEN, True)])
block(ws, 27, 'E', '予定年間収支（対象年）', [
    ('対象年', f"='{SH['ann']}'!B2&\"年\"", None, False),
    ('経費の見込み合計', f"='{SH['ann']}'!F38", YEN, False),
    ('収支の見込み', f"='{SH['ann']}'!G38", '#,##0;-#,##0;0', True)])

# ============================================================ 使い方
ws = ws_guide
ws.sheet_view.showGridLines = False
for col, w in zip('ABCD', (2, 16, 24, 84)):
    ws.column_dimensions[col].width = w
for r in range(1, 60):
    for col in 'ABCD':
        ws[f'{col}{r}'].fill = fill_cream
ws.merge_cells('B1:D1')
ws['B1'].value = '家賃入金チェック表'; ws['B1'].font = Font(name=FONT, size=22, bold=True, color='FFFFFF'); ws['B1'].fill = fill_head
ws['B1'].alignment = Alignment(vertical='center', indent=1); ws.row_dimensions[1].height = 46
ws.merge_cells('B2:D2')
ws['B2'].value = '小規模の個人大家向け（手入力版・契約100件まで）。入金、滞納、期限、敷金、空室、修繕、税、家賃改定、収支をまとめて管理します。'
ws['B2'].font = Font(name=FONT, size=10, color='555555'); ws['B2'].alignment = Alignment(vertical='center', indent=1); ws.row_dimensions[2].height = 26

def heading(r, text):
    ws.merge_cells(f'B{r}:D{r}')
    c = ws[f'B{r}']; c.value = text; c.font = Font(name=FONT, size=13, bold=True, color=GREEN)
    c.alignment = Alignment(vertical='center'); ws.row_dimensions[r].height = 28
    for col in 'BCD':
        ws[f'{col}{r}'].border = Border(bottom=Side(style='medium', color=GREEN))

# 注意
ws.merge_cells('B4:D4')
ws['B4'].value = 'ご利用の前に'; ws['B4'].font = Font(name=FONT, size=11, bold=True, color='9A4B00'); ws['B4'].fill = fill_warn
ws['B4'].alignment = Alignment(vertical='center', indent=1); ws.row_dimensions[4].height = 24
ws.merge_cells('B5:D5')
ws['B5'].value = ('このシートの判定は、入力した記録にもとづく目安です。導入後の最初の数か月は、通帳や銀行の明細と必ず突き合わせて確認してください。'
                  '年間収支、空室損失、家賃改定の試算は参考値です。税金や会計の扱いは、税務署や税理士に確認してください。サンプルのデータはすべて架空です。')
ws['B5'].font = f_base; ws['B5'].fill = fill_warn; ws['B5'].alignment = Alignment(wrap_text=True, vertical='top', indent=1)
ws.row_dimensions[5].height = 62

heading(7, 'はじめの5ステップ')
steps = [
 ('設定を確認する', '「設定」シートで、滞納を黄色・赤にする日数などを確認します。最初は初期値のままで使えます。'),
 ('契約を登録する', '「台帳」シートに、契約を1行ずつ入力します（100件まで）。号室は重複させません。サンプルの行は、書き換えて使います。'),
 ('入金を記録する', '入金があったら「入金記録」シートに1行追加します。何月分の家賃かを、対象の年・月で入力します。'),
 ('入金チェックを見る', '「入金チェック」シートで、対象の年・月を選びます。入金済、一部入金、未入金、滞納の黄・赤が出ます。'),
 ('月に1回、突き合わせる', '通帳や銀行の明細と、入金記録を突き合わせます。漏れや誤りがあれば、入金記録を直します。')]
for i, (t, d_) in enumerate(steps, start=1):
    r = 7 + i
    c = ws[f'B{r}']; c.value = i; c.font = Font(name=FONT, size=14, bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor=ORANGE)
    c.alignment = CENTER; c.border = Border(left=Side(style='thick', color=CREAM), right=Side(style='thick', color=CREAM), top=Side(style='thick', color=CREAM), bottom=Side(style='thick', color=CREAM))
    ws[f'C{r}'].value = t; ws[f'C{r}'].font = Font(name=FONT, size=11, bold=True, color=GREEN); ws[f'C{r}'].alignment = WRAPC
    ws[f'D{r}'].value = d_; ws[f'D{r}'].font = f_base; ws[f'D{r}'].alignment = WRAPC
    ws.row_dimensions[r].height = 36

heading(14, 'シートの一覧')
for col, t in zip('BCD', ('使いかた', 'シート', '役割')):
    c = ws[f'{col}15']; c.value = t; c.font = f_head; c.fill = fill_head; c.alignment = CENTER
ws.row_dimensions[15].height = 22
guide_rows = [
 ('入力', '設定', '日数や期限の基準を決めます。'),
 ('入力', '台帳', '契約ごとに、入居者、家賃、支払期日、契約・保険の期限、家賃の改定履歴を入力します。'),
 ('入力', '入金記録', '入金のたびに1行追加します。分割や保証会社からの入金も、1回ごとに1行です。'),
 ('入力', '請求調整', '退去月など、その月だけ請求額を変えるときに入力します（任意）。'),
 ('確認', '入金チェック', '対象の年・月の入金状態と、滞納の黄・赤、一般保証型の事故報告の期限を表示します。'),
 ('入力＋確認', '保証会社入金', '保証会社からの1回の振込と、入居者への割り当て合計の差を確認します。'),
 ('確認', '期限管理', '契約更新、火災保険、保証会社の契約の残り日数を、色で表示します。'),
 ('入力＋確認', '敷金台帳', '預り金、控除、返還額、追加請求額、精算の遅れを管理します。'),
 ('入力＋確認', '空室・募集', '空室の日数、募集家賃、内見・申込の件数、空室損失（概算）を管理します。'),
 ('入力＋確認', '修繕履歴', '工事・交換の履歴と費用を記録し、設備の交換時期が近づくと色で知らせます。'),
 ('入力＋確認', '税カレンダー', '固定資産税などの納期限、納付額、納付日を管理し、納付予定額を出します。'),
 ('入力＋確認', '物件台帳', '物件ごとの取得価格を入れると、戸数・家賃の集計、固定資産税の推定、年間の予定収支、表面利回りが出ます。'),
 ('入力＋確認', '家賃改定試算', '改定案を入力すると、増減額、年間の影響額（概算）が出ます。'),
 ('入力＋確認', '年間収支', '月別の家賃収入と経費、収支を出します。修繕費と税金は自動で入ります。'),
 ('確認', 'ダッシュボード', '各シートの集計を1画面で見られます。'),
]
tagfill = {'入力': ORANGE, '確認': '2F7D5B', '入力＋確認': '8A6D3B'}
for i, (tag, name, desc) in enumerate(guide_rows):
    r = 16 + i
    c = ws[f'B{r}']; c.value = tag; c.font = Font(name=FONT, size=9, bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor=tagfill[tag])
    c.alignment = CENTER; c.border = Border(left=Side(style='thick', color=CREAM), right=Side(style='thick', color=CREAM), top=Side(style='medium', color=CREAM), bottom=Side(style='medium', color=CREAM))
    ws[f'C{r}'].value = name; ws[f'C{r}'].font = f_bold
    ws[f'C{r}'].hyperlink = f"#'{name}'!A1"
    ws[f'C{r}'].font = Font(name=FONT, size=10, bold=True, color=GREEN, underline='single')
    ws[f'C{r}'].alignment = WRAPC
    ws[f'D{r}'].value = desc; ws[f'D{r}'].font = f_base; ws[f'D{r}'].alignment = WRAPC
    ws.row_dimensions[r].height = 26
    for col in 'BCD':
        pass

r0 = 16 + len(guide_rows) + 1
heading(r0, '色の見かた')
samples = [
 ('入力するセル', '黄色の背景・青い文字です。ここだけを書き換えます。', fill_input, f_input),
 ('自動で出るセル', '灰色の背景・黒い文字です。計算式が入っているので、書き換えません。', fill_gray, f_base),
 ('黄', '注意の段階です（期日を過ぎた、期限が近い、など）。', PatternFill('solid', fgColor='FFF0A8'), f_bold),
 ('赤', '対応が必要な段階です（滞納が続いている、期限が迫っている、など）。', PatternFill('solid', fgColor='F8CBCB'), f_bold),
 ('緑', '入金済、納付済など、問題のない状態です。', PatternFill('solid', fgColor='D5EBD9'), f_bold),
]
for i, (name, desc, fl, fn) in enumerate(samples, start=1):
    r = r0 + i
    c = ws[f'B{r}']; c.value = name; c.fill = fl; c.font = fn; c.alignment = CENTER
    c.border = Border(left=Side(style='thick', color=CREAM), right=Side(style='thick', color=CREAM), top=Side(style='medium', color=CREAM), bottom=Side(style='medium', color=CREAM))
    ws.merge_cells(f'C{r}:D{r}')
    ws[f'C{r}'].value = desc; ws[f'C{r}'].font = f_base; ws[f'C{r}'].alignment = WRAPC
    ws.row_dimensions[r].height = 24

r1 = r0 + len(samples) + 2
heading(r1, '知っておくと便利なこと')
tips = [
 ('号室は重複させない', '台帳の号室は、1行ずつ別にします。同じ号室に新しい入居者が入るときは、退去済の行を上書きします。'),
 ('家賃を改定したとき', '台帳の「改定前の家賃」と「家賃の改定日」を入力します。過去の月は、改定前の家賃で計算されます。2回以上の改定は、請求調整で対応します。'),
 ('退去月の請求額', '日割りにする場合は、請求調整にその月の請求額を入力します。'),
 ('台帳の行は並べ替えない', '台帳の行を並べ替えたり、削除したりすると、他のシートの表示がずれることがあります。退去した契約は「退去済」にします。'),
 ('「含める」の扱い', '修繕履歴の「年間収支に含める」は、修繕費として経費にするかどうかの選択です。資本的支出との区別は、税理士に確認してください。'),
]
for i, (t, d_) in enumerate(tips, start=1):
    r = r1 + i
    ws[f'B{r}'].value = '■'; ws[f'B{r}'].font = Font(name=FONT, size=10, color=ORANGE); ws[f'B{r}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'C{r}'].value = t; ws[f'C{r}'].font = Font(name=FONT, size=10, bold=True, color=GREEN); ws[f'C{r}'].alignment = WRAPC
    ws[f'D{r}'].value = d_; ws[f'D{r}'].font = f_base; ws[f'D{r}'].alignment = WRAPC
    ws.row_dimensions[r].height = 40

r2 = r1 + len(tips) + 2
heading(r2, '税制・法制が変わったとき（ご自身で対応する手順）')
law_steps = [
 ('変更の内容を確かめる', '総務省・国税庁・お住まいの自治体の発表や、届いた納税通知書で、税率や特例の割合が変わったかを確かめます。'),
 ('設定の⑩を書き換える', '「設定」シートの⑩「税制・法制の値」の該当セルを書き換えます。税率は％で（例：1.4%）、割合は「=1/6」のようにイコールから入力します。'),
 ('確認日と出典をメモする', '同じ行の右端「確認日・出典（メモ）」に、書き換えた日と根拠（自治体名、通知書の年度など）を入力します。次に見直すときの手がかりになります。'),
 ('結果を確かめる', '「物件台帳」の税の推定と、「年間収支」の予定年間収支が変わったことを確かめます。納税通知書が届いたら税カレンダーに入力すると、推定より入力値が優先されます。'),
 ('設定で表せない変更のとき', '新しい特例ができたなど、仕組みそのものが変わった場合は、値の書き換えでは対応できません。物件台帳の「土地の評価額（使用値）」〜「税の合計（推定）」の列の計算式を直す必要があります。'),
]
for i, (t, d_) in enumerate(law_steps, start=1):
    r = r2 + i
    c = ws[f'B{r}']; c.value = i; c.font = Font(name=FONT, size=14, bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor=GREEN)
    c.alignment = CENTER; c.border = Border(left=Side(style='thick', color=CREAM), right=Side(style='thick', color=CREAM), top=Side(style='thick', color=CREAM), bottom=Side(style='thick', color=CREAM))
    ws[f'C{r}'].value = t; ws[f'C{r}'].font = Font(name=FONT, size=11, bold=True, color=GREEN); ws[f'C{r}'].alignment = WRAPC
    ws[f'D{r}'].value = d_; ws[f'D{r}'].font = f_base; ws[f'D{r}'].alignment = WRAPC
    ws.row_dimensions[r].height = 44
r3 = r2 + len(law_steps) + 1
ws.merge_cells(f'B{r3}:D{r3}')
ws[f'B{r3}'].value = ('制度に関わるその他の値：保証会社の事故報告の期限（設定の⑫）は契約書で、敷金の精算の目安（設定の⑥）はご自身の運用で決める値です。'
                      '契約や運用が変わったときも、同じように設定シートで書き換えます。')
ws[f'B{r3}'].font = f_base; ws[f'B{r3}'].fill = fill_warn; ws[f'B{r3}'].alignment = Alignment(wrap_text=True, vertical='top', indent=1)
ws.row_dimensions[r3].height = 40
for rr in range(60, r3 + 3):
    for col in 'ABCD':
        if ws[f'{col}{rr}'].fill.fgColor.rgb in (None, '00000000'):
            ws[f'{col}{rr}'].fill = fill_cream

# ---- 全体 ----
for w in wb.worksheets:
    if w.title not in (SH['guide'], SH['dash']):
        w.sheet_view.showGridLines = False
    w.page_setup.orientation = 'landscape'
    w.page_setup.fitToWidth = 1
    w.page_setup.fitToHeight = 0
    w.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
for name, color in ((SH['guide'], ORANGE), (SH['dash'], GREEN), (SH['set'], 'A0A0A0')):
    wb[name].sheet_properties.tabColor = color
for name in (SH['pro'], SH['led'], SH['pay'], SH['adj'], SH['gua'], SH['dep'], SH['vac'], SH['rep'], SH['tax'], SH['rev']):
    wb[name].sheet_properties.tabColor = 'E8C25A'
for name in (SH['chk'], SH['due'], SH['ann']):
    wb[name].sheet_properties.tabColor = '2F7D5B'
wb.save(OUT)
print('saved', OUT)

import json as _json
_json.dump({'S': {k: v.split('!')[1].replace('$', '') for k, v in S.items()}, 'GU0': GU0}, open(OUT + '.addr.json', 'w'))
