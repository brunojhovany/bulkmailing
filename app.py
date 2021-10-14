import sys
import os
import time
import ssl
import email
import smtplib
from loguru import logger
from yaml import safe_load
from datetime import datetime
from yaml.error import YAMLError
from pandas import read_excel, DataFrame, read_csv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

config = None
date_now = datetime.now()
date_str = date_now.strftime('%Y-%m-%d')
logger.add(f'logs/file_{date_str}.log', rotation="12:00", level='ERROR')

def throw(_, error, message, *args, **kwargs):
    message = message.format(*args, **kwargs)
    logger.opt(depth=1).error(message)
    sys.exit(1)

logger.__class__.throw = throw

try:
    with open('config.yaml', 'r') as stream:
        config = safe_load(stream)
        config = config['config']
except Exception as err:
    err_inf = f'Error al cargar el archivo de configuracion ☠️ | {err}'
    logger.throw(err, err_inf)


def sendMail():
    email_config = config['email']
    email_sender = email_config['user_name']
    email_password = email_config['password']

    smtp_ssl_config = email_config['SMTP_SSL']
    context = ssl.create_default_context()
    server = smtplib.SMTP_SSL(smtp_ssl_config['host'], int(
        smtp_ssl_config['port']), context=context)
    server.ehlo()
    try:
        server.login(user=email_sender, password=email_password)
        logger.success('Inicio de sesion con la cuenta de correo exitoso.')
    except smtplib.SMTPAuthenticationError as err:
        err_inf = f'Error al iniciar sesion en la cuenta de correo | {err}'
        logger.throw(err, err_inf)

    try:
        data_config = email_config['data']
        data: DataFrame = read_excel(
            '%s/%s' % (data_config['path'], data_config['file_name']))

        creative_config: dict = email_config['creative']
        with open('%s/%s' % (creative_config['path'], creative_config['file_name']), 'r', encoding="utf8") as file:
            creative = file.read().replace('\n', '')
    except Exception as err:
        err_inf = f'Error al cargar los archivos necesarios | {err}'
        logger.throw(err, err_inf)

    html = creative
    email_data = email_config['email_data']
    count = 0

    for index, element in data.iterrows():
        name_reader: str = element['First Name']
        message = MIMEMultipart("alternative")
        message["From"] = email_data['sender']
        message.attach(MIMEText(email_data['bodytext'], "plain"))
        message.attach(MIMEText(html, "html"))
        message["Subject"] = f'{name_reader} ' + email_data['subject']
        message["To"] = element['Email']
        message['Date'] = email.utils.formatdate()
        message['Message-ID'] = email.utils.make_msgid(domain="reedem.site")
        count += 1
        try:
            server.sendmail(
                email_sender, element['Email'], message.as_string()
            )
            logger.info(str(count) + ". Sent to " + element['Email'])
        except Exception as error:
            logger.error(
                "No se pudo enviar el correo al destinatario: %s ☠️" % element['Email'])

        if(count % 80 == 0):
            server.quit()
            print("Server cooldown for 100 seconds")
            time.sleep(100)
            server.ehlo()
            server.login(email_sender, email_password)
    server.quit()


if __name__ == '__main__':
    logger.info('Iniciando aplicacion... 📬🔥 | version:1.0.0')
    sendMail()
