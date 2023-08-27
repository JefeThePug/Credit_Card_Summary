import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

smtp = smtplib.SMTP('smtp.gmail.com', 587) #465 587
sender = "email@email.com"

if smtp.ehlo()[0] != 250:
    sys.exit("Could not connect to email server")

if smtp.starttls()[0] != 220:
    sys.exit("Could not start TLS encryption)")

smtp.login(sender, sys.argv[3])

msg = MIMEMultipart()
msg['From'] = sender
msg['To'] = sys.argv[2]
msg['Date'] = formatdate(localtime=True)
file = sys.argv[1][-11:]
msg['Subject'] = f"Finances: {file}"
msg.attach(MIMEText(f"Attached: PDF of Financial summary {file[:-4]}"))

pdf = MIMEBase('application', "octet-stream")
with open(sys.argv[1], 'rb') as f:
    pdf.set_payload(f.read())
encoders.encode_base64(pdf)
pdf.add_header('Content-Disposition', f'attachment; filename={file}')
msg.attach(pdf)

result = smtp.sendmail(sender, sys.argv[2], msg.as_string())

if result:
    smtp.quit()
    sys.exit("error sending: {result}")

smtp.quit()
sys.exit("File sent to email!")