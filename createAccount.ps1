$source = Import-Csv -Path .\createUser.csv -Encoding Default

$password = ConvertTo-SecureString 'NbVcX2345!@' -AsPlainText -Force
$Usercredential= New-Object System.Management.Automation.PSCredential ('delta\fc.nea', $password )
$session2 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://twtpeexhybrid03.delta.corp/PowerShell/ -Authentication Kerberos -Credential $Usercredential

Import-PSSession $session2
$ErrorLog = ".\error_log.txt"

foreach ($user in $source) {
    try {
        if ($user.SamAccountName -ne "") {
            $account = $user.SamAccountName.ToUpper()
            $pwdString = "Dej" + $user.startDate
            $defaultPwd = ConvertTo-SecureString $pwdString -AsPlainText -Force
            $firstName = $account.split('.')[0]
            $lastName = $account.split('.')[1]
            $ChineseName = $user.chineseName
            $OU = $user.OU
            $Name = $account
            $DisplayName = $account

            if ($OU -eq "JPDEJ") {
                $firstName = $account
                $lastName = $ChineseName
                $Name = $account + " " + $ChineseName
                $DisplayName = $account + " " + $ChineseName
            }

            if ($account -and $user.startDate -and $user.OU) {
                New-Mailbox -UserPrincipalName "$account@deltaww.com" `
                -Alias $account `
                -Database "O365_TEMP-NEW" `
                -Name $Name `
                -OrganizationalUnit "delta.corp/Delta/JP/$OU/Users" `
                -Password $defaultPwd `
                -FirstName $firstName `
                -LastName $lastName `
                -DisplayName $DisplayName
            }
        }
    } catch {
        $ErrMsg = $_.Exception.Message
        Add-Content $ErrorLog $ErrMsg
    }
}
