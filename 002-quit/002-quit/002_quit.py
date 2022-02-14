import jhMail
import datetime
import subprocess
import os
def quit_process(org, adpswd, username):
    of = "D:/103_MIS/jh_scripts/ad.v2/002-quit/002-quit/azure_cmd/quit_%s_%s_%s.ps1" % (org, td, username)
    if org.upper() == "OPPO":
        out_cmd = """
            Import-Module MSOnline
            $passwd=ConvertTo-SecureString '%s' -AsPlainText -Force
            $liveCred = New-Object System.Management.Automation.PSCredential('admin@oppo-aed.tw', $passwd)
            Connect-MsolService -Credential $liveCred
            $o365Uri = 'https://outlook.office365.com/powershell-liveid/'
            $uid ="admin@oppo-aed.tw"
            $pwd = (Get-ItemProperty -Path hkcu:software\Hsg).ohs | ConvertTo-SecureString -AsPlainText -Force
            $pwd.MakeReadOnly()
            $o365Cred = New-Object System.Management.Automation.PSCredential($uid,$pwd)
            $exoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $o365Uri -Credential $o365Cred -Authentication Basic -AllowRedirection
            Import-PSSession $exoSession -AllowClobber
            $SamAcc="%s"
            $oppoaddress=$SamAcc + "@oppomobile.tw"
            Set-Mailbox -Identity $SamAcc -type "Shared"
            Get-Mailbox -Identity $SamAcc | fl *display*,*share*
            Set-Mailbox $SamAcc -HiddenFromAddressListsEnabled $true
            Get-Mailbox -Identity $SamAcc | fl *display*,*HiddenFromAddressListsEnabled*
            Remove-DistributionGroupMember -Identity opporealme_all@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity opporealme_tp@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity oppo_all@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity oppo_tp@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity realme_all@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity realme_tp@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity ceo@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity gm@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity ad@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity fn@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity sr@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity pd@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity se@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity br@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity te@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity sa@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity account.notice@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity dt@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity digital@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity ec@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity EDTI_Daily.Report@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity ehr.ceo@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity gp.o365.test@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity it_alarm@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity mas@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity cmt@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity promoter-ka@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity pd.test@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity pd.gtm@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity sc@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity gp.op.1f1700@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity sn@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity ProjectDept@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity nb@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity gp.op.1f1800@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity st@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity oppotwhq@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity tp_sa@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity sk@oppo-aed.tw -Member $SamAcc -Confirm:$false
            Remove-DistributionGroupMember -Identity pd-1@oppo-aed.tw -Member $SamAcc -Confirm:$false
            $o365acc="%s"
            $o365domain="@oppo-aed.tw"
            $o365upn="$o365acc"+"$o365domain"
            Get-MsolUser -UserPrincipalName $o365upn | fl DisplayName,Licenses,BlockCredential
            $user=@(Get-MsolUser -UserPrincipalName $o365upn)
            $user.licenses.AccountSkuId
            Get-MsolUser -UserPrincipalName $o365upn | Set-MsolUserLicense -RemoveLicenses $user.licenses.AccountSkuId
            Set-MsolUser -UserPrincipalName $o365upn -BlockCredential $true
            Set-MsolUserPassword -UserPrincipalName $o365upn -NewPassword "1qaz@WSX"
            Get-MsolUser -UserPrincipalName $o365upn | fl DisplayName,Licenses,BlockCredential
            """ % (adpswd, username, username)
    elif org.upper() == "REALME":
        out_cmd = """
            Import-Module MSOnline
            $passwd=ConvertTo-SecureString '%s' -AsPlainText -Force
            $liveCred = New-Object System.Management.Automation.PSCredential('admin@realme.com.tw', $passwd)

            Connect-MsolService -Credential $liveCred
            $o365Uri = 'https://outlook.office365.com/powershell-liveid/'
            $uid ="admin@realme.com.tw"
            $pwd = (Get-ItemProperty -Path hkcu:software\Hsg).ohs | ConvertTo-SecureString -AsPlainText -Force
            $pwd.MakeReadOnly()
            $o365Cred = New-Object System.Management.Automation.PSCredential($uid,$pwd)
            $exoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $o365Uri -Credential $o365Cred -Authentication Basic -AllowRedirection
            Import-PSSession $exoSession -AllowClobber
            $SamAcc="%s"
            Set-Mailbox -Identity $SamAcc -type "Shared"
            Get-Mailbox -Identity $SamAcc | fl *display*,*share*
            Set-Mailbox $SamAcc -HiddenFromAddressListsEnabled $true
            Get-Mailbox -Identity $SamAcc | fl *display*,*HiddenFromAddressListsEnabled*
            Remove-DistributionGroupMember -Identity realmetwhq@realme.com.tw -Member $SamAcc -Confirm:$false
            $o365acc="%s"
            $o365domain="@realme.com.tw"
            $o365upn="$o365acc"+"$o365domain"
            Get-MsolUser -UserPrincipalName $o365upn | fl DisplayName,Licenses,BlockCredential
            $user=@(Get-MsolUser -UserPrincipalName $o365upn)
            $user.licenses.AccountSkuId
            Get-MsolUser -UserPrincipalName $o365upn | Set-MsolUserLicense -RemoveLicenses $user.licenses.AccountSkuId
            Set-MsolUser -UserPrincipalName $o365upn -BlockCredential $true
            Set-MsolUserPassword -UserPrincipalName $o365upn -NewPassword "1qaz@WSX"
            Get-MsolUser -UserPrincipalName $o365upn | fl DisplayName,Licenses,BlockCredential
            """ % (adpswd, username, username)
    else:
        return "org error, must be OPPO or Realme"
    ff = open(of, "w")
    ff.write(out_cmd)
    ff.close()
    return of

def disable_ad_user(quitName):
    out_cmd=f"Disable-ADAccount -Identity {quitName}\n"
    return out_cmd

td = datetime.datetime.today().strftime("%Y%m%d")
inFilePath = f"D:/103_MIS/jh_scripts/ad.v2/002-quit/002-quit/input/quitList_{td}.txt"
cronLogPath = f"D:/103_MIS/jh_scripts/ad.v2/002-quit/002-quit/cron_log/000_cron_quit.log"
inFileExisted = os.path.isfile(inFilePath)
f = open(cronLogPath, "a")
f.write(f"{inFilePath} is {inFileExisted} existed.\n")
f.close()
if inFileExisted:
    f = open(cronLogPath, "a")
    f.write(f"There is quit co-worker in {td}.\n")
    f.close()
    infile = open(inFilePath, "r", encoding="UTF-8")
    du_output = open(f"//192.168.0.1/ModifyAD/disable_AD_user.ps1", "w")
    du_output.write("Import-Module ActiveDirectory\n")
    du_output.close()
    du_output = open(f"//192.168.0.1/ModifyAD/disable_AD_user.ps1", "a")
    employeeList = open("D:/103_MIS/jh_scripts/ad.v2/employeeList.txt", "r", encoding="UTF-8").readlines()
    for quitList in infile.readlines():
        quitList_sep = quitList.split(sep="\t")
        quitName = quitList_sep[1]
        quitInformation = employeeList[[quitName in i for i in employeeList].index(True)].strip().split(",")
        quitEmail = quitInformation[3]
        quitIdentity = quitEmail.split("@")[0]
        if 'oppo-aed.tw' in quitEmail:
            quitOrg = 'OPPO'
        elif 'realme.com.tw' in quitEmail:
            quitOrg = 'realme'
        else:
            quitOrg = 'ERROR'
        f = open(cronLogPath, "a")
        f.write(f"{quitOrg}-{quitName} quit in {td}.\n")
        f.close()
        cmdPath = quit_process(org=quitOrg, adpswd="Xin1mao4huan2qiu2", username=quitIdentity)
        subprocess.call(["C:/WINDOWS/system32/WindowsPowerShell/v1.0/powershell.exe", cmdPath])
        disableUser = disable_ad_user(quitName=quitIdentity)
        du_output.write(disableUser)
    du_output.close()
else:
    cronLogPath = f"D:/103_MIS/jh_scripts/ad.v2/002-quit/002-quit/cron_log/000_cron_quit.log"
    f = open(cronLogPath, "a")
    f.write(f"There is no quit co-worker in {td}.\n")
    f.close()