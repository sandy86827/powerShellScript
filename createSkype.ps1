$source = Import-Csv -Path .\createUser.csv -Encoding Default
$password = ConvertTo-SecureString 'Admin.Full1688' -AsPlainText -Force
$Usercredential= New-Object System.Management.Automation.PSCredential ('delta\fc.nea', $password )
$session1 = New-PSSession -ConnectionURI https://twsfbpool01.deltaww.com/OcsPowershell -Credential $UserCredential

Import-PSSession $session1


foreach ($user in $source) {
    $account = $user.SamAccountName.ToUpper()

    if ($account) {
        Write-Host "User: $account"
        Write-Host "Creating TEL ......"

        $UPN = "$account@deltaww.com"
        $vposaka = 'JP-Osaka-International'
        $dposaka = 'JP-Osaka'
        $vptokyo = 'JP-Tokyo-International'
        $dptokyo = 'JP-Tokyo'
        $vpseoul = 'KR-Seoul-International'
        $dpseoul = 'KR-Seoul'
        $phone = $user.telephoneNumber.split('x')[0] -replace "\s+"
        $sitecode = $user.otherTelephone.split('-')[0]
        $ext = $user.otherTelephone.split('-')[1]
        $line = 'TEL:' + $phone + ';ext=' + $ext
        $privateline = 'TEL:+' + $sitecode + $ext
        $line_137 = $privateline + ';ext=' + $ext
        $line_138 = $privateline + ';ext=' + $ext

        try {
            Enable-CsUser -Identity $account -RegistrarPool 'twsfbpool01.deltaww.com' -SipAddressType SamAccountName -SipDomain deltaww.com -DomainController TWTPEDCS02

            if ($sitecode -eq '137') {
                Set-CsUser -Identity $account -EnterpriseVoiceEnabled $true -LineURI $line_137 -DomainController TWTPEDCS02
                Grant-CsDialPlan -Identity $account -PolicyName $dposaka -DomainController TWTPEDCS02
                Grant-CsVoicePolicy -Identity $account -PolicyName $vposaka -DomainController TWTPEDCS02
                Start-Sleep -Seconds 60
                Set-CsClientPin -Identity $account -pin '123456'
            } elseif ($sitecode -eq '138') {
                Set-CsUser -Identity $account -EnterpriseVoiceEnabled $true -LineURI $line_138 -DomainController TWTPEDCS02
                Grant-CsDialPlan -Identity $account -PolicyName $dpseoul -DomainController TWTPEDCS02
                Grant-CsVoicePolicy -Identity $account -PolicyName $vpseoul -DomainController TWTPEDCS02
                Start-Sleep -Seconds 60
                Set-CsClientPin -Identity $account -pin '123456'
            } else {
                Set-CsUser -Identity $account -EnterpriseVoiceEnabled $true -LineURI $line -PrivateLine $privateline -DomainController TWTPEDCS02
                Grant-CsDialPlan -Identity $account -PolicyName $dptokyo -DomainController TWTPEDCS02
                Grant-CsVoicePolicy -Identity $account -PolicyName $vptokyo -DomainController TWTPEDCS02
                Start-Sleep -Seconds 60
                Set-CsClientPin -Identity $account -pin '123456'
            }

            Add-DistributionGroupMember -Identity "jpdej@deltaww.com" -Member "$account@deltaww.com"
            Add-ADGroupMember -Identity "JPSG_ALL" -Members $account
            Add-ADGroupMember -Identity "L-JP-SSLVPN" -Members $account
            $check = $account[0] + $account[1]
            if ($check -ne 'V-') {
                Get-Mailbox -Identity $account | Set-Mailbox -CustomAttribute10 "o365new"
            }

        } catch {
            $ErrMsg = $_.Exception.Message
            Add-Content $ErrorLog $ErrMsg
        }
    } 
}

# 清理 PowerShell 会话
Get-PSSession | Remove-PSSession