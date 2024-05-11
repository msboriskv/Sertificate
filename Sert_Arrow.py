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
import smtplib
import os

# workBook = load_workbook(filename = 'Anckett.xlsx')
# sheets = workBook.active

# baze_point = 340

# full_name = ''
# code = ''

# for row in range(1, sheets.max_row + 1):
    
#     fist_name = sheets[row][0].value
#     if fist_name == None:
#         fist_name = ''
        
#     name = sheets[row][1].value
#     if name == None:
#         name = ''
        
#     second_name = sheets[row][2].value
#     if second_name == None:
#         second_name = ''
        
#     code = sheets[row][3].value
#     if code == None:
#         code = ''
        
#     mail = sheets[row][4].value
#     if mail == None:
#         mail = ''
    
#     full_name = fist_name + ' ' + name + ' ' + second_name
    
#     print(row, full_name, code, mail)

#     pdf = FPDF('P', 'mm', 'A4')

#     pdf.add_font("Sans", style="", fname="Noto Sans/NotoSans-Regular.ttf", uni=True)
#     pdf.add_font("Sans", style="B", fname="Noto Sans/NotoSans-Bold.ttf", uni=True)
#     pdf.add_font("Sans", style="I", fname="Noto Sans/NotoSans-Italic.ttf", uni=True)
#     pdf.add_font("Sans", style="BI", fname="Noto Sans/NotoSans-BoldItalic.ttf", uni=True)

#     pdf.add_page()

#     pdf.image("555.png", x=5, y=5, w=200)

#     pdf.set_font("Sans", "B", 20)
#     pdf.cell(0, baze_point, 'Получил', new_x="LMARGIN", align='C')
#     pdf.cell(0, baze_point + 20, full_name, new_x="LMARGIN", align='C')

#     pdf.set_font("Sans", "", 16)
#     pdf.cell(0, baze_point + 60, 'за участие в XII международном форуме', new_x="LMARGIN", align='C')
#     pdf.cell(0, baze_point + 75, '"ОБРАЗОВАНИЕ: РЕАЛИИ И ПЕРСПЕКТИВЫ"', new_x="LMARGIN", align='C')

#     pdf.set_font("Sans", "", 14)
#     pdf.cell(0, 500, code, new_x="LMARGIN", new_y="NEXT", align='C')
    
#     pdf.set_display_mode(zoom='fullpage', layout='continuous')

#     pdf.output(full_name + '_' + code + '.pdf')

send = "msboriskv@gmail.com"
pas = "aixd svcr amnq fsvj"
to = "boris@karnaval.su"

server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()

server.login(send, pas)

server.sendmail(send, to, "Hello fghdfgjhfsgjfhsdgjhf")