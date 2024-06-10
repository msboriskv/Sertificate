# -*- coding: cp1251 -*-
# -*- coding: utf8 -*-

from asyncio.windows_events import NULL
import encodings
import json
from queue import Full
from venv import create
from fpdf import FPDF
import smtplib
import os
import mimetypes
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from dotenv import load_dotenv

load_dotenv()

sender = "msboriskv@gmail.com"
password = os.getenv("point")
mail = 'boris@karnaval.su'
            
msg = MIMEMultipart()
msg["From"] = sender
msg["To"] = mail
msg["Subject"] = 'Сертификат форума "Образование: Реалии и перспективы"'

msg.attach(MIMEText("Здравствуйте, отправляем Вам сертификат форума - Образование: Реалии и перспективы"))

# for file in os.listdir("To_Send"):
#     filename = os.path.basename(file)
#     ftype, encoding = mimetypes.guess_type(file)
#     file_type, subtype = ftype.split("/")
    
#     if file_type == "text":
#         with open(f"To_Send/{file}") as f:
#             file = MIMEText(f.read())
#     elif file_type == "image":
#         with open(f"To_Send/{file}", "rb") as f:
#             file = MIMEImage(f.read(), subtype)
#     elif file_type == "audio":
#         with open(f"To_Send/{file}", "rb") as f:
#             file = MIMEAudio(f.read(), subtype)
#     elif file_type == "application":
#         with open(f"To_Send/{file}", "rb") as f:
#             file = MIMEApplication(f.read(), subtype)
#     else:
#         with open(f"To_Send"/{file}, "rb") as f:
#             file = MIMEBase(file_type, subtype)
#             file.set_payload(f.read())
#             encoders.encode_base64(file)
    
    # file.add_header('Content-Disposition', 'attachment', filename=filename)
    # msg.attach(file)

server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()

server.login(sender, password)

server.sendmail(sender, mail, msg.as_string())