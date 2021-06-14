import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path

def Email(Get,Email):
        sender = r'wandee.payslip@gmail.com'
        password = r'TandB101'
        send_to_email = Email
        subject = 'Payslip'
        message = 'This is your Payslip for this month'
        filename = (r'PAYROLL' + Get + '.pdf')
#wandee@tandbmediaglobal.com
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = send_to_email
        # Setup the attachment
        msg.attach(MIMEText(message, 'plain'))
        attachment = open(filename, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

                # Attach the attachment to the MIMEMultipart object
        msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com',587)
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, send_to_email, msg.as_string())
        print(server)
        server.quit()




