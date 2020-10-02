import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
import win32com.client as win32
import re


def email_to(file_path, address):
    smtp_server = 'smtp.qq.com'
    from_address = '2587354021@qq.com'
    passwd = 'rxuowjfzqqindhhj'
    to_address = address

    server = smtplib.SMTP_SSL(smtp_server)
    server.login(from_address, passwd)

    msg = MIMEMultipart()
    filename = file_path
    subject = filename.split('/')[-1][:-4]

    msg['Subject'] = Header(subject)
    msg['From'] = Header(from_address)
    msg['To'] = Header(to_address)

    part = MIMEApplication(open(filename, 'rb').read())
    part.add_header('Content-Disposition', 'attachment', filename=filename.split('/')[-1])
    msg.attach(part)
    try:
        server.sendmail(from_address, to_address, msg.as_string())
    except smtplib.SMTPException:
        server.quit()
        return False
    else:
        server.quit()
        return True


def sendMailByOutLook(file_path, address_list):
    if not file_path:
        return
    pat = re.compile(r'^[A-Za-z\d]+([-_.][A-Za-z\d]+)*@([A-Za-z\d]+[-.])+[A-Za-z\d]{2,4}$')
    suc = 0
    fail = 0
    for address in address_list:
        # print(address)
        # print(type(address))
        if pat.match(str(address)):
            sendSingleMail(file_path, address)
            suc += 1
        else:
            fail += 1
    return suc, fail


def sendSingleMail(file_path, address):
    outlook = win32.gencache.EnsureDispatch('Outlook.Application')

    mail_item = outlook.CreateItem(0)

    mail_item.Recipients.Add(address)
    mail_item.Subject = 'Share a book'

    mail_item.Attachments.Add(file_path)
    mail_item.Send()
