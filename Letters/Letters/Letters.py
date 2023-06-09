import smtplib
import os
import csv
import fitz
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders

def send_email(from_address, password, to_address, subject, body, a):
    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = subject
    body_text = MIMEText(body)
    msg.attach(body_text)
    with open(a, "rb") as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=a)
        msg.attach(part)
    server = smtplib.SMTP('smtp.mail.ru', 587)
    server.starttls()
    server.login(from_address, password)
    server.sendmail(from_address, to_address, msg.as_string())
    server.quit()

with open('data.csv', encoding='utf-8') as f:
    reader = csv.reader(f)
    mail = input('Введите свой почтовый адрес: ')
    p = input('Введите пароль: ')
    for row in reader:
        email_address = row[0]
        surname = row[1]
        name = row[2]
        attachment_path = f'{surname}_{name}_приглашение.pdf'
        with fitz.open(attachment_path) as doc:
            page = doc.load_page(0)
            attachment_text = page.get_text()
        file = f'{surname}_{name}_сертификат.pdf'
        send_email(mail, p, email_address, 'Приглашение', attachment_text, file)
        print('Приглашение на адрес ' + email_address + ' отправлено!')
