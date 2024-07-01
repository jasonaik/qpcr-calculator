import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate
from tkinter import messagebox


def show_error(text):
    messagebox.showerror("Error", text)


def email_excel(files: list, password, send_from, send_to: str, cc="", server="smtp.gmail.com", port=587):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Cc'] = cc
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = 'qPCR Results'
    for file in files:
        fp = open(file, 'rb')
        part = MIMEBase('application', 'vnd.ms-excel')
        part.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=file)
        msg.attach(part)
        msg.attach(MIMEText(""))
    smtp = smtplib.SMTP(server, port)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(send_from, password)
    try:
        smtp.send_message(msg)
    except smtplib.SMTPDataError:
        show_error("Invalid Email Entered")
        raise Exception("Invalid Email Entered")
    smtp.quit()


