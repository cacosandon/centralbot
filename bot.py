import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
from credentials import mail, password
import smtplib, ssl
from email.mime.text import MIMEText

sched = BlockingScheduler(timezone="America/Santiago")

@sched.scheduled_job('cron', hour=14)
def task():
    port = 465  # For SSL
    smtp_server = "smtp.gmail.com"
    sender = mail

    # API INFO
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('estacioncentral.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open('PATENTES CENTRAL EJEMPLO').worksheet("Prueba")

    for x in sheet.get_all_records():

        nombre, receiver, estado, comentario = x["Nombre"], x["Correo"], x["Estado patente"], x["Comentario"]

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            message = f"""Hola, \n\nNombre empresa: {nombre}\n\nLa obtención de patente en la Municipalidad de Estación Central para tu empresa está en el estado:  {estado}. {comentario}\n\nPor favor no responder este correo,\nMunicipalidad de Estación Central"""

            msg = MIMEText(message)
            msg['Subject'] = 'Estado de patente - Municipalidad Estación Central'
            msg['From'] = sender
            msg['To'] = receiver

            server.login(sender, password)
            server.sendmail(sender, receiver, msg.as_string())

sched.start()
