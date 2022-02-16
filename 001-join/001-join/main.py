from jinja2 import Environment
import jhMail
import _001_join
import datetime
import os
import subprocess
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
execDate = datetime.datetime.today().strftime("%Y%m%d")
test1 = _001_join.NewEmployee(noobFile = f"D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/input/noobsList_{execDate}.txt")
test1.check_duplicated()
test1.create_PS_input()
for _theCompany in set(test1.companys):
    if not os.path.isfile(f"D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/output/{_theCompany}/list_{test1._execDate}.csv"):
        cred = open("D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/999_credentials.txt", "r", encoding="UTF-8").readlines()
        sender = cred[0].split(":")[1].strip()
        sender_password = cred[1].split(":")[1].strip()
        jhMail.JiehongEmail(subject='Create Account Failed',
                            sender=sender,
                            receivers=['leo.hsu@oppo-aed.tw'],
                            cc=['jack.hong@oppo-aed.tw','tsunghui.li@oppo-aed.tw'],
                            mail_content = f'{test1._execDate} Create {",".join(test1.Emails)} Failed.',
                            mail_content_style='txt',
                            sender_password=sender_password).sendMail()
    else:
        test1.create_PS_cmd()
        test1.write_new_employee_data()
        for i in ['oppo','realme']:
            if os.path.isfile(f"{test1.casePath}azure_cmd/AzureAD_{i}_{test1._execDate}.ps1"):
                try:
                    subprocess.call(["C:/WINDOWS/system32/WindowsPowerShell/v1.0/powershell.exe", f"{test1.casePath}azure_cmd/AzureAD_{i}_{test1._execDate}.ps1" ])
                    print("Azre AD done!")
                except Exception as e:
                    f = open(f"D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/logs/powershell_failed_{execDate}.log")
                    print("Azure AD Failed!")
                    time.sleep(600)

        test1.send_noob_mail()
