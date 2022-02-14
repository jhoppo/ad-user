import datetime
import jhMail
class NewEmployee:
    def __init__(self, noobFile):
        self._execDate = datetime.datetime.today().strftime("%Y%m%d")
        self.casePath = "D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/"
        _readNoobData = open(noobFile,"r",encoding="UTF-8")
        self._noobData = _readNoobData.readlines()
        _readNoobData.close()
        self._noobLength = len(self._noobData)
        # input: _noobData = OPPO   A99999  洪測測(Test Hong)  女   2022/02/14  執行長室    拓展管理組   專員  
        # EmployeeID input[1]
        self.eIDs = [i.split("\t")[1] for i in self._noobData]
        # Company input[0]
        self.companys = [i.split("\t")[0] for i in self._noobData]
        # office: 總部/xx分部
        self.offices = [i.split("\t")[5] if '分部' in i.split("\t")[5] else '總部' for i in self._noobData ]
        # Name 洪測測
        self.names = [i.split("\t")[2].split("(")[0] for i in self._noobData]
        # surname & givenname
        # surname 洪
        self.surnames = [i[0] for i in self.names]
        # givenname 測測
        self.givennames = [i[1:] for i in self.names]
        # DisName Test Hong
        self.disNames = [i.split("\t")[2].split("(")[1].split(")")[0] for i in self._noobData]
        # Licenses 沒特殊要求都是標準版
        self.licenses = ['標準版'] * self._noobLength
        # sAmAccountNames test.hong
        self.sAmAccountNames = [i.split("\t")[2].split("(")[1].replace(" ",".").lower().replace(")","") for i in self._noobData]
        # Displayname test.hong(洪測測)
        self.Displaynames = [ "%s(%s)" % (self.sAmAccountNames[i], self.names[i]) for i in range(len(self._noobData))]
        # Password: Today@ + _jionDate = Today@0214
        self._joinDate = [datetime.datetime.strptime(i.split("\t")[4], "%Y/%m/%d").strftime("%m%d") for i in self._noobData]
        self.Passwords = [f'Today@{i}' for i in self._joinDate]
        # Email: test.hong@oppo-aed.tw OR test.hong@realme.com.tw # it's up to {Company}
        self.Emails = [f"{self.sAmAccountNames[i]}@oppo-aed.tw" if self.companys[i] == 'OPPO' else f"{self.sAmAccountNames[i]}@realme.com.tw" for i in range(self._noobLength)]
        # UserPrincipalName: same as Email
        self.UserPrincipalNames = self.Emails
        # Department: 執行長室
        self.Departments = [i.split("\t")[5] for i in self._noobData]
        # Team: 拓展管理組
        self.Teams = [i.split("\t")[6] for i in self._noobData]
        # title: 專員
        self.titles = [i.split("\t")[7] for i in self._noobData]
        # manager: no data
        self.managers = [''] * self._noobLength
        # Description: OPPO.執行長室.拓展管理組
        self.Descriptions = [f'{self.companys[i]}.{self.Departments[i]}.{self.Teams[i]}' for i in range(self._noobLength)]
    def check_duplicated(self):
        employeeList = open("D:/103_MIS/jh_scripts/ad.v2/employeeList.txt", "r", encoding="UTF-8").readlines()
        employeeEmailList = [i.split(",")[3] for i in employeeList]
        res = [i in employeeEmailList for i in self.Emails]
        [i in employeeEmailList for i in self.Emails]
        True_indices = [i for i in range(self._noobLength) if res[i] is True ]
        for i in True_indices:
            self.sAmAccountNames[i] = self.sAmAccountNames[i]+"."+self.eIDs[i].lower()
            self.Emails[i] = self.Emails[i].replace("@", f".{self.eIDs[i].lower()}@")
            self.UserPrincipalNames[i] = self.Emails[i]
        return
    def create_PS_input(self):
        for _theCompany in set(self.companys):
            # on local
            f = open(f"D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/output/{_theCompany}/list_{self._execDate}.csv","w")
            f.write("No,EmployeeID,Company,Office,Name,surname,givenname,Displayname,DisName,License,sAmAccountName,Password,Email,UserPrincipalName,Department,Team,title,manager,Description\n")
            [f.write(f"{i+1},{self.eIDs[i]},{self.companys[i]},{self.offices[i]},{self.names[i]},{self.surnames[i]},{self.givennames[i]},{self.Displaynames[i]},{self.disNames[i]},{self.licenses[i]},{self.sAmAccountNames[i]},{self.Passwords[i]},{self.Emails[i]},{self.UserPrincipalNames[i]},{self.Departments[i]},{self.Teams[i]},{self.titles[i]},{self.managers[i]},{self.Descriptions[i]}\n") for i in range(self._noobLength)]
            f.close()
        # on op01
        opF = open(f"//192.168.0.1/ModifyAD/list_{self._execDate}.csv", "w")
        opF.write("No,EmployeeID,Company,Office,Name,surname,givenname,Displayname,DisName,License,sAmAccountName,Password,Email,UserPrincipalName,Department,Team,title,manager,Description\n")
        [opF.write(f"{i+1},{self.eIDs[i]},{self.companys[i]},{self.offices[i]},{self.names[i]},{self.surnames[i]},{self.givennames[i]},{self.Displaynames[i]},{self.disNames[i]},{self.licenses[i]},{self.sAmAccountNames[i]},{self.Passwords[i]},{self.Emails[i]},{self.UserPrincipalNames[i]},{self.Departments[i]},{self.Teams[i]},{self.titles[i]},{self.managers[i]},{self.Descriptions[i]}\n") for i in range(self._noobLength)]
        opF.close()
        return
    def write_new_employee_data(self):
        _fn = "D:/103_MIS/jh_scripts/ad.v2/employeeList.txt"
        f = open(_fn, "a",encoding="UTF-8")
        [f.write(f"{self.names[i]},{self.surnames[i]},{self.givennames[i]},{self.Emails[i]},{self.Departments[i]},{self.eIDs[i]}\n") for i in range(self._noobLength)]
        f.close()
        return
    def create_PS_cmd(self):
        if 'OPPO' in self.companys or 'oppo' in self.companys:
            fn_O = f"AzureAD_oppo_{self._execDate}.ps1"
            full_fn_O_path = self.casePath +"azure_cmd/"+ fn_O
            f = open(full_fn_O_path, "w")
            f.write("Import-Module MSOnline\n")
            f.write("$passwd=ConvertTo-SecureString 'Xin1mao4huan2qiu2' -AsPlainText -Force\n")
            f.write("$liveCred = New-Object System.Management.Automation.PSCredential('admin@oppo-aed.tw', $passwd)\n")
            f.write("$listname=%s\n" % ('"list_' + self._execDate + '"'))
            f2 = open(self.casePath+"AzureAD_oppo_execute.txt", "r", encoding='utf-8')
            for f2_lines in f2.readlines():
                f.write(f2_lines)
            f.close()
            f2.close()
        if 'realme' in self.companys or 'REALME' in self.companys:
            fn_R = f"AzureAD_realme_{self._execDate}.ps1"
            full_fn_R_path = self.casePath +"azure_cmd/"+ fn_R
            f = open(full_fn_R_path, "w")
            f.write("Import-Module MSOnline\n")
            f.write("$passwd=ConvertTo-SecureString 'Xin1mao4huan2qiu2' -AsPlainText -Force\n")
            f.write("$liveCred = New-Object System.Management.Automation.PSCredential('admin@realme.com.tw', $passwd)\n")
            f.write("$listname=%s\n" % ('"list_' + self._execDate + '"'))
            f2 = open(self.casePath+"AzureAD_realme_execute.txt", "r", encoding='utf-8')
            for f2_lines in f2.readlines():
                f.write(f2_lines)
            f.close()
            f2.close()
        return
    def printerCode(self):
        res_code=[]
        printers = open("D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/printer_code.txt", "r", encoding="UTF-8").readlines()
        for i in self.Departments:
            check_res = []
            for j in printers:
                if i == j.split("\t")[0]:
                    check_res.append(True)
                else:
                    check_res.append(False)
            if sum(check_res) == 0:
                res_code.append("NA")
            else:
                res_code.append(printers[check_res.index(True)].split("\t")[1].strip())
        return res_code
    def update_print_server():
        in_file = open("D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/output/list_" + td + ".csv", "r")
        user_info = [[i.split(",")[4], i.split(",")[10], i.split(",")[12]] for i in in_file.readlines()[1:]]
        in_file.close()
        printer_server = "http://192.168.0.232/?MAIN=TOPACCESS"
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.get(printer_server)
        time.sleep(10)
        # click login button
        driver.switch_to.frame("TopLevelFrame")
        driver.switch_to.frame("topframe")
        driver.find_element_by_xpath("/html/body/div/table/tbody/tr[@class='clsBackYellow']/td[@class='clsDeviceStatusLink']/a[@class='clsLogin']").click()
        time.sleep(5)
        # key in login
        driver.switch_to.frame("TopLevelFrame")
        driver.switch_to.frame("Loginframe")
        driver.find_element_by_name("USERNAME").send_keys("admin")
        driver.find_element_by_name("PASS").send_keys("123456")
        driver.find_element_by_name("Login").click()
        time.sleep(10)
        driver.switch_to.frame("TopLevelFrame")
        driver.switch_to.frame("topframe")
        driver.find_element_by_id("REGISTRATION-anchor").click()
        time.sleep(3)
        # click AddressBook
        driver.switch_to.default_content()
        driver.switch_to.frame("TopLevelFrame")
        driver.switch_to.frame("SubMenu")
        driver.find_element_by_id("AddressBook").click()
        time.sleep(3)
        for the_user in user_info:
            # click add new contacts
            driver.switch_to.default_content()
            driver.switch_to.frame("TopLevelFrame")
            driver.switch_to.frame("contents")
            driver.switch_to.frame("fraTitle")
            driver.find_element_by_name("btnLdapAdd").click()
            time.sleep(3)
            # key in new contact's data
            driver.switch_to.default_content()
            driver.switch_to.frame("TopLevelFrame")
            driver.switch_to.frame("contents")
            driver.switch_to.frame("fraList")
            driver.find_element_by_id("NAME").send_keys(the_user[0])
            driver.find_element_by_id("FURIGANA").send_keys(the_user[1][0])
            driver.find_element_by_name("EMAIL").send_keys(the_user[2])
            driver.find_element_by_id("SEL_DEST").send_keys("分享")
            # click save new contact
            driver.switch_to.default_content()
            driver.switch_to.frame("TopLevelFrame")
            driver.switch_to.frame("contents")
            driver.switch_to.frame("fraTitle")
            driver.find_element_by_name("btnSave").click()
            time.sleep(3)
        # click share settings
        driver.switch_to.default_content()
        driver.switch_to.frame("TopLevelFrame")
        driver.switch_to.frame("contents")
        driver.switch_to.frame("fraTitle")
        driver.find_element_by_id("SHAREDSETT").click()
        time.sleep(3)
        # click sync
        driver.switch_to.default_content()
        driver.switch_to.frame("TopLevelFrame")
        driver.switch_to.frame("contents")
        driver.switch_to.frame("fraList")
        driver.switch_to.frame("fraTitle")
        driver.find_element_by_id("btnSynAll").click()
        # click I am sure
        driver.switch_to.alert.accept()
        time.sleep(10)
        driver.close()
        return
    def send_noob_mail(self):
        cred = open("D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/999_credentials.txt", "r", encoding="UTF-8").readlines()
        sender = cred[0].split(":")[1].strip()
        sender_password = cred[1].split(":")[1].strip()
        ccList = open("D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/cc.txt","r").readlines()
        template = open("email.j2", "r", encoding='UTF-8').read()
        templated = [Environment().from_string(template).render(noobAccount=self.sAmAccountNames[i],Password=self.Passwords[i],Email=self.Emails[i],printer=self.printerCode()[i]) for i in range(len(self.sAmAccountNames)) ]
        templated_path = [f"D:/103_MIS/jh_scripts/ad.v2/001-join/001-join/output/999_{self._execDate}_mail_{self.sAmAccountNames[i]}.html" for i in range(len(templated))]
        [open(templated_path[i],"w",encoding='UTF-8').write(templated[i]) for i in range(len(templated))]
        [ jhMail.JiehongEmail(subject = '【新進人員】系統帳號密碼通知', sender=sender, sender_password=sender_password, receivers=[self.Emails[i]], cc=['jack.hong@oppo-aed.tw','tsunghui.li@oppo-aed.tw'], mail_content = templated[i], mail_content_style='html').sendHTMLMail() for i in range(len(self.Emails)) ]
        return templated_path

def forti_ssl_vpn(device, username, password, add_user_info, add_group):
    #### add_user_info: [ 'username', 'user_email' ]
    #### user_group: "SSL_ERP" or "SSL_IT"
    remote_device = {
        "device_type": "fortinet",
        "host": device,
        "username": username,
        "password": password
    }
    net_connect = ConnectHandler(**remote_device)
    print("connected.....start to config")
    for the_user in add_user_info:
        add_user_name, add_user_mail = the_user
        add_cmds = "config user local\nedit %s\nset type ldap\nset ldap-server op01\nset status enable\nset email-to %s\nset two-factor email\nend\nconfig user group\nedit %s\nappend member %s\nend\n" % (add_user_name, add_user_mail, add_group, add_user_name)
        add_cmds_list = add_cmds.split("\n")
        print(add_cmds_list)
        net_connect.send_config_set(add_cmds_list)
        print("config %s end" % (add_user_name))
    net_connect.disconnect()
    print("close connection")
    return
def forti_backup(device, username, password):
    remote_device = {
        "device_type": "fortinet",
        "host": device,
        "username": username,
        "password": password
    }
    net_connect = ConnectHandler(**remote_device)
    backup_list = open("D:/103_MIS/jh_scripts/for_ad/cmds/forti_backup_cmds.txt", "r").readlines()
    backup_fn = "D:/103_MIS/jh_scripts/for_ad/outputs/backup/forti_%s_%s.config" % (device, datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
    with open(backup_fn, "w", encoding="UTF-8") as write_backup:
        [write_backup.write(net_connect.send_command(the_cmd)) for the_cmd in backup_list]
    net_connect.disconnect()
    return

