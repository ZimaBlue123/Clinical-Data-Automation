# -*- coding: utf-8 -*-
"""
[Legacy / 示例] 针对特定项目版式（Test-1/Test-2、固定工作表与单元格区域）的一次性脚本。

生产环境请使用 word_to_excel_to_figure.py（通用 chart 引用 + Excel COM 保结构）。
"""
import re, shutil, win32com.client, openpyxl
from pathlib import Path

def strip(t):
    return str(t).replace('\x07','').replace('\r','').replace('\x0b',' ').strip() if t else ""

def get_pct(t):
    m = re.search(r'\(([0-9.]+)\)', str(t))
    return float(m.group(1)) if m else None

def match(w, e):
    w = re.sub(r'[^A-Z0-9\u4e00-\u9fff]', '', strip(w)).upper()
    e = re.sub(r'[^A-Z0-9\u4e00-\u9fff]', '', strip(e)).upper()
    return e in w or w in e

base = Path(__file__).parent
word = win32com.client.Dispatch('Word.Application')
word.Visible = False

doc1 = word.Documents.Open(str(base/'input'/'Test-1.docx'))
tbl1 = doc1.Content.Tables.Item(1)
data1 = {}
for j in range(2, 7):
    try:
        g = strip(tbl1.Cell(1, j).Range.Text)
        if not g: continue
        col = 2*j-1
        for i in range(3, tbl1.Rows.Count+1):
            try:
                row = strip(tbl1.Cell(i, 1).Range.Text)
                if not row: continue
                v = get_pct(tbl1.Cell(i, col).Range.Text)
                if v is not None:
                    if row not in data1: data1[row] = {}
                    data1[row][g] = v
            except: pass
    except: pass
doc1.Close(False)

doc2 = word.Documents.Open(str(base/'input'/'Test-2.docx'))
tbl2 = doc2.Content.Tables.Item(1)
data2 = {}
for j in range(2, 7):
    try:
        g = strip(tbl2.Cell(1, j).Range.Text)
        if not g: continue
        col = 2*j-1
        for i in range(3, tbl2.Rows.Count+1):
            try:
                row = strip(tbl2.Cell(i, 1).Range.Text)
                if not row: continue
                v = get_pct(tbl2.Cell(i, col).Range.Text)
                if v is not None:
                    if row not in data2: data2[row] = {}
                    data2[row][g] = v
            except: pass
    except: pass
doc2.Close(False)
word.Quit()

tmpl = list((base/'Template').glob('*.xlsx'))[0]
out = base/'output'/'final_output.xlsx'
out.parent.mkdir(exist_ok=True)
shutil.copy2(tmpl, out)

wb = openpyxl.load_workbook(str(out))
ws = wb['0-30天ADR PT']

for r in range(3, 13):
    rn = ws.cell(r,1).value
    if not rn: continue
    wr = next((k for k in data1 if match(k, rn)), None)
    if not wr: continue
    for c in range(2, 7):
        h = ws.cell(2,c).value
        if not h: continue
        wg = next((k for k in data1[wr] if match(k, h)), None)
        if wg: ws.cell(r,c).value = data1[wr][wg]

for r in range(15, 25):
    rn = ws.cell(r,1).value
    if not rn: continue
    wr = next((k for k in data2 if match(k, rn)), None)
    if not wr: continue
    for c in [2,3]:
        h = ws.cell(14,c).value
        if not h: continue
        wg = next((k for k in data2[wr] if match(k, h)), None)
        if wg: ws.cell(r,c).value = data2[wr][wg]

wb.save(str(out))
wb.close()
print(f"完成：{out}")
