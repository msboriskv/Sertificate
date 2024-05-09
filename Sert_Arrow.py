# -*- coding: cp1251 -*-
# -*- coding: utf8 -*-

from asyncio.windows_events import NULL
import json
from venv import create
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import rows_from_range
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, Fill, PatternFill, NamedStyle
from fpdf import FPDF

workBook = load_workbook(filename = 'Anckett.xlsx')
sheets = workBook.active

full_name = ''
code = ''
for collum in range(0, sheets.max_collum):
    if sheets[0].value == NULL:
        break
    for row in range(1, sheets.max_row):
        fist_name = sheets[row][0].value
        name = sheets[row][1].value
        second_name = sheets[row][2].value
        code = sheets[row][3].value
        mail = sheets[row][4].value
    
        full_name = fist_name+' '+name+second_name
    
        print(row, full_name)

pdf = FPDF('P', 'mm', 'A4')

pdf.add_font("Sans", style="", fname="Noto Sans/NotoSans-Regular.ttf", uni=True)
pdf.add_font("Sans", style="B", fname="Noto Sans/NotoSans-Bold.ttf", uni=True)
pdf.add_font("Sans", style="I", fname="Noto Sans/NotoSans-Italic.ttf", uni=True)
pdf.add_font("Sans", style="BI", fname="Noto Sans/NotoSans-BoldItalic.ttf", uni=True)

pdf.add_page()

pdf.image("555.png", x=5, y=5, w=200)

pdf.set_font("Sans", "B", 20)
pdf.cell(0, 340, '������� ' + full_name, new_x="LMARGIN", align='C')

pdf.set_font("Sans", "", 16)
pdf.cell(0, 380, '�� ������� � XII ������������� ������', new_x="LMARGIN", align='C')
pdf.cell(0, 395, '"�����������: ������ � �����������"', new_x="LMARGIN", align='C')

pdf.set_font("Sans", "", 14)
pdf.cell(0, 500, code, new_x="LMARGIN", new_y="NEXT", align='C')

pdf.output('Forum.pdf')