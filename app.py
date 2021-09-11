import time
import ssl
import smtplib
from loguru import logger
from yaml import safe_load
from datetime import datetime
from yaml.error import YAMLError
from pandas import read_excel, DataFrame
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


config = None
date_now = datetime.now()
date_str = date_now.strftime('%Y-%m-%d')
logger.add(f'logs/file_{date_str}.log', rotation="12:00", level='ERROR')


with open('config.yaml', 'r') as stream:
    try:
        config = safe_load(stream)
        config = config['config']
    except YAMLError as err:
        err_inf = f'Error al cargar el archivo de configuracion | {err}'
        logger.error(err_inf)
        raise Exception(err_inf)



def sendMail():
    email_config = config['email']
    email_sender = email_config['user_name']
    email_password = email_config['password']

    smtp_ssl_config = email_config['SMTP_SSL']
    context = ssl.create_default_context()
    server = smtplib.SMTP_SSL(smtp_ssl_config['host'], int(smtp_ssl_config['port']), context=context )
    server.ehlo()
    try:
        server.login(user=email_sender,password= email_password)
        logger.info('Inicio de sesion con la cuenta de correo exitoso.')
    except smtplib.SMTPAuthenticationError as err:
        err_inf = f'Error al iniciar sesion en la cuenta de correo | {err}'
        logger.error(err_inf)
        raise Exception(err_inf)

    data :DataFrame = read_excel('data/example-data.xlsx')
    
    with open('creative/preview.html', 'r', encoding="utf8") as file:
        creative = file.read().replace('\n', '')


    html = creative
    email_data = email_config['email_data']
    message = MIMEMultipart("alternative")
    message["From"] = email_sender
    message.attach(MIMEText(html, "html"))
    count = 0

    for index, element in data.iterrows():
        print(index, element)
        name_reader: str = element['First Name']
        message["Subject"] =  f'{name_reader} ' + email_data['subject']
        count += 1
        logger.info(str(count) + ". Sent to " + element['Email'])

        server.sendmail(
            email_sender, data['Email'], message.as_string()
        )

        if(count%80 == 0):
            server.quit()
            print("Server cooldown for 100 seconds")
            time.sleep(100)
            server.ehlo()
            server.login(email_sender, email_password)
    server.quit()


if __name__ == '__main__':
    logger.info('Iniciando aplicacion..')
    sendMail()