import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
class JiehongEmail:
    """
    the variables receivers and cc are list []
    """
    def __init__(self, subject, sender, receivers, cc, mail_content, mail_content_style, sender_password):
        self.subject = subject
        self.sender = sender
        self.receivers = receivers
        self.cc = cc
        self.mail_content = mail_content
        self.sender_password = sender_password
        self.mail_content_style = mail_content_style
    def sendMail(self):
        smtp_host = 'smtp.office365.com'
        msg = MIMEText(self.mail_content)
        msg['From'] = self.sender
        msg['To'] = ",".join(self.receivers)
        msg['Cc'] = ",".join(self.cc)
        msg['Subject'] = self.subject
        smtp_server = smtplib.SMTP(smtp_host,587)
        smtp_server.starttls()
        smtp_server.login(self.sender,self.sender_password)
        smtp_server.sendmail(self.sender,self.receivers+self.cc,msg.as_string())
        smtp_server.quit() 
        return
    def sendHTMLMail(self):
        smtp_host = 'smtp.office365.com'
        msg = MIMEMultipart()
        msgText = MIMEText(self.mail_content, self.mail_content_style, 'UTF-8')
        msg.attach(msgText)
        msg['From'] = self.sender
        msg['To'] = ",".join(self.receivers)
        msg['Cc'] = ",".join(self.cc)
        msg['Subject'] = self.subject
        smtp_server = smtplib.SMTP(smtp_host,587)
        smtp_server.starttls()
        smtp_server.login(self.sender,self.sender_password)
        smtp_server.sendmail(self.sender,self.receivers+self.cc,msg.as_string())
        smtp_server.quit() 
        return


#template = open("email.j2", "r").read()
#templated = Environment().from_string(template).render(noobAccount='nora.yang.a00814',Password='Today@0214',Email='nora.yang.a00814@oppo-aed.tw',printer='345621')
#open("nora.yang.a00814.html", "w").write(templated)

#test_email = JiehongEmail(subject = 'test-class-sender', sender = 'jiehongnoreply@gmail.com', receivers=['sock40038@gmail.com', 'sock81912@gmail.com'], cc=['jack.hong@oppo-aed.tw'], mail_content = templated, mail_content_style='html', sender_password='xxxxxxxx')
#test_email.sendHTMLMail()