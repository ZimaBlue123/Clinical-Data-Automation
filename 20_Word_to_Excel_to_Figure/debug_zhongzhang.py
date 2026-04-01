# -*- coding: utf-8 -*-
import sys, re, win32com.client
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

base = Path(__file__).parent
word = win32com.client.Dispatch('Word.Application')
word.Visible = False

doc = word.Documents.Open(str(base/'input'/'Test-1.docx'))
tbl = doc.Content.Tables.Item(1)

print("Word中包含'肿胀'的行：")
for i in range(3, min(20, tbl.Rows.Count+1)):
    try:
        row = tbl.Cell(i, 1).Range.Text.strip()
        if '肿胀' in row:
            print(f"  行{i}: '{row}'")
            for j in [3,5,7,9,11]:
                try:
                    val = tbl.Cell(i, j).Range.Text.strip()
                    print(f"    列{j}: '{val}'")
                except: pass
    except: pass

doc.Close(False)
word.Quit()

print("\nExcel中的条目名：")
import openpyxl
tmpl = list((base/'Template').glob('*.xlsx'))[0]
wb = openpyxl.load_workbook(str(tmpl), data_only=True)
ws = wb['0-30天ADR PT']
for r in range(3, 13):
    rn = ws.cell(r,1).value
    if rn and '肿胀' in str(rn):
        print(f"  第{r}行: '{rn}'")
wb.close()
