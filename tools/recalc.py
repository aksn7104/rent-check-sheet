# -*- coding: utf-8 -*-
"""LibreOffice（無料の表計算ソフト）で .xlsx の計算式をすべて計算し直し、結果を同じファイルに保存する。
build_template.py が作るファイルには計算式だけが入っていて、計算結果はまだ入っていないため、
突き合わせ（verify_template.py）の前にこれで計算させる。
注意：LibreOffice で保存し直すため、書式の一部が変わることがある。検証用のコピーに使い、配布用のファイルには使わない
（配布用のファイルは、Google スプレッドシートや Excel で開いたときに自動で計算される）。
実行: python3 tools/recalc.py <ファイル.xlsx> [待ち時間（秒）]
結果は JSON で出力する: {"status": "success"|"errors_found", "total_formulas": n, "total_errors": n, "errors": [...]}"""
import sys, os, json, shutil, subprocess, tempfile
from pathlib import Path
from openpyxl import load_workbook

MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
Sub RecalculateAndSave()
  ThisComponent.calculateAll()
  ThisComponent.store()
  ThisComponent.close(True)
End Sub
</script:module>"""
ERRORS = ('#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', '#NUM!', '#N/A')

def soffice():
    for name in ('soffice', 'libreoffice'):
        if shutil.which(name):
            return name
    sys.exit(json.dumps({'error': 'LibreOffice（soffice）が見つかりません。インストールしてください。'}))

def recalc(path, timeout=180):
    path = Path(path).resolve()
    exe = soffice()
    with tempfile.TemporaryDirectory() as tmp:
        profile = Path(tmp) / 'profile'
        url = profile.as_uri()
        env = dict(os.environ, HOME=tmp)
        # 1回目：設定フォルダを作るだけ
        subprocess.run([exe, f'-env:UserInstallation={url}', '--headless', '--terminate_after_init'],
                       env=env, capture_output=True, timeout=timeout)
        basic = profile / 'user' / 'basic' / 'Standard'
        basic.mkdir(parents=True, exist_ok=True)
        (basic / 'Module1.xba').write_text(MACRO, encoding='utf-8')
        # 2回目：ファイルを開いて、全部計算し直して保存する
        subprocess.run([exe, f'-env:UserInstallation={url}', '--headless', '--norestore',
                        'vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application',
                        str(path)], env=env, capture_output=True, timeout=timeout)
    formulas = load_workbook(path)
    values = load_workbook(path, data_only=True)
    total, errors = 0, []
    for ws in formulas.worksheets:
        wv = values[ws.title]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    total += 1
                    v = wv[c.coordinate].value
                    if isinstance(v, str) and v in ERRORS:
                        errors.append(f'{ws.title}!{c.coordinate} {v}')
    return {'status': 'errors_found' if errors else 'success', 'total_formulas': total,
            'total_errors': len(errors), 'errors': errors[:100]}

if __name__ == '__main__':
    t = int(sys.argv[2]) if len(sys.argv) > 2 else 180
    print(json.dumps(recalc(sys.argv[1], t), ensure_ascii=False, indent=2))
