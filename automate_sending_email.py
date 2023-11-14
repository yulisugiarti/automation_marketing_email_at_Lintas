import datetime as dt
import pandas as pd
import smtplib
import ssl

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

now = dt.datetime.now()
day_now = now.date().day
month_now = now.date().month

data = pd.read_csv("Daftar Nama Leads.csv")
list_name_email = {row_content.Nama: row_content.email for (index, row_content) in data.iterrows()}
list_name_field = {row_content.Nama: row_content.Bidang_Usaha for (index_data, row_content) in data.iterrows()}

#we can also change the day_now and month_now to be the same date and month inside the CSV file that we target
name_email_list = {row.Nama: row.email for (index, row) in data.iterrows() if row.Tanggal == day_now and row.Bulan == month_now}

with open("Offering_letter.html") as content:
    fix_content = content.read()
    for names in list_name_field:
        subject = f"Intro Lintas.app for {names}"
        mail_content = fix_content.replace("[COMPANY]", names)
        complete_mail = mail_content.replace("[FIELD]", list_name_field[names])
        body = complete_mail
        sender_email = "PUT YOUR EMAIL ADRESS HERE"
        password = "PUT YOUR APP PASSWORD HERE (it can be from google account or email that you use"
        receiver_email = name_email_list[names]
        # Create a multipart message and set headers
        message = MIMEMultipart("alternative")
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject
        # Add body to email
        message.attach(MIMEText(body, "html"))
        filename = "LINTAS Company Profile.pdf"  # In same directory as script
        # Open PDF file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)
        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={filename}",
        )
        # Add attachment to message and convert message to string
        message.attach(part)
        text = message.as_string()
        # Log in to server using secure context and send email
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, text)
            server.close()
