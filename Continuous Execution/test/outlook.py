import smtplib
from email.message import EmailMessage

sender = "vft-test@outlook.com"
receiver = "xuche1998@gmail.com"
password = "vflowtech123"
msg_body = "Hello, this is a test email from Python."

msg = EmailMessage()
msg.set_content(msg_body)
msg['Subject'] = 'Test Email'
msg['From'] = sender
msg['To'] = receiver

server = smtplib.SMTP('smtp-mail.outlook.com', 587)
server.starttls()
server.login(sender, password)
server.send_message(msg)
server.quit()
