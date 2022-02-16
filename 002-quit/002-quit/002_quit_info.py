import datetime
import subprocess
import os
from jinja2 import Environment
import time
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


execute_date = datetime.datetime.today().strftime("%Y%m%d")
info_date = [int(i.split("_")[1].replace(".txt","")) for i in os.listdir("D:/103_MIS/jh_scripts/ad.v2/002-quit/002-quit/input") if int(i.split("_")[1].replace(".txt","")) >= int(execute_date)]
quit_imformation = [open(f"D:/103_MIS/jh_scripts/ad.v2/002-quit/002-quit/input/quitList_{i}.txt", "r", encoding="UTF-8").readlines() for i in info_date ]
employee_list = open("D:/103_MIS/jh_scripts/ad.v2/employeeList.txt", "r", encoding="UTF-8").readlines()
employee_names = [i.split(",")[0] for i in employee_list]
cred = open("D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/999_credentials.txt", "r", encoding="UTF-8").readlines()
sender = cred[0].split(":")[1].strip()
sender_password = cred[1].split(":")[1].strip()
ccList = open("D:/103_MIS/jh_scripts/ad.v2/cc.txt","r").readlines()
for i in quit_imformation:
    for the_row in i:
        splited = the_row.split("\t")
        templated_output = f"D:/103_MIS/jh_scripts/ad.v2/002-quit/002-quit/output/{datetime.datetime.strptime(splited[3],'%Y/%m/%d').strftime('%Y%m%d')}_{splited[0]}_{splited[1]}.html"
        if not os.path.exists(templated_output):
            quit_supervisor = splited[-1].strip()
            quit_supervisor_index = employee_names.index(quit_supervisor)
            quit_supervisor_Email = employee_list[quit_supervisor_index].split(",")[3]
            template = open("D:/103_MIS/jh_scripts/ad.v2/002-quit/002-quit/email.j2", "r", encoding="UTF-8").read()
            templated = Environment().from_string(template).render(employee_ID=splited[0], employee_name=splited[1], employee_E_name=splited[2], last_work_date=splited[3], quit_date=splited[4], depart_team_title=splited[5], supervisor=splited[6])
            open(templated_output,"w",encoding="UTF-8").write(templated)
            JiehongEmail(subject='同仁離職通知',sender=sender,sender_password=sender_password,mail_content=templated,mail_content_style='html',receivers=[quit_supervisor_Email],cc=ccList).sendHTMLMail()
