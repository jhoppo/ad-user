import jhMail
import datetime
import subprocess
import os
from jinja2 import Environment
import time
info_date = max([int(i.split("_")[1].replace(".txt","")) for i in os.listdir("D:/103_MIS/jh_scripts/ad.v2/002-quit/002-quit/input")])
quit_imformation = open(f"D:/103_MIS/jh_scripts/ad.v2/002-quit/002-quit/input/quitList_{info_date}.txt", "r", encoding="UTF-8").readlines()
employee_list = open("D:/103_MIS/jh_scripts/ad.v2/employeeList.txt", "r", encoding="UTF-8").readlines()
employee_names = [i.split(",")[0] for i in employee_list]
cred = open("D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/999_credentials.txt", "r", encoding="UTF-8").readlines()
sender = cred[0].split(":")[1].strip()
sender_password = cred[1].split(":")[1].strip()
for i in quit_imformation:
    i = quit_imformation[0]
    splited = i.split("\t")
    templated_output = f"D:/103_MIS/jh_scripts/ad.v2/002-quit/002-quit/output/{datetime.datetime.strptime(splited[3],'%Y/%m/%d').strftime('%Y%m%d')}_{splited[0]}_{splited[1]}.html"
    if not os.path.exists(templated_output):
        quit_supervisor = splited[-1]
        quit_supervisor_index = employee_names.index(quit_supervisor)
        quit_supervisor_Email = employee_list[quit_supervisor_index].split(",")[3]
        template = open("D:/103_MIS/jh_scripts/ad.v2/002-quit/002-quit/email.j2", "r", encoding="UTF-8").read()
        templated = Environment().from_string(template).render(employee_ID=splited[0], employee_name=splited[1], employee_E_name=splited[2], last_work_date=splited[3], quit_date=splited[4], depart_team_title=splited[5], supervisor=splited[6])
        open(templated_output,"w",encoding="UTF-8").write(templated)
        jhMail.JiehongEmail(subject='同仁離職通知',sender=sender,sender_password=sender_password,mail_content=templated,mail_content_style='html',receivers=[quit_supervisor_Email],cc=['leo.hsu@oppo-aed.tw', 'jack.hong@oppo-aed.tw']).sendHTMLMail()