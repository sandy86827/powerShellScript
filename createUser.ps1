
$source = Import-Csv -Path .\user_service_1_DEJ.csv -Encoding Default

$password = ConvertTo-SecureString 'Admin.Full1688' -AsPlainText -Force
$Usercredential= New-Object System.Management.Automation.PSCredential ('delta\fc.nea', $password )
$session1 = New-PSSession -ConnectionURI https://twsfbpool01.deltaww.com/OcsPowershell -Credential $UserCredential
$Session2 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://twtpeexhybrid03.delta.corp/PowerShell/ -Authentication Kerberos -Credential $UserCredential

Import-PSSession $session1
Import-PSSession $session2
$ErrorLog = ".\error_log.txt"

ForEach ($user In $source) {
    try {
        
        $account = $user.SamAccountName.ToUpper()
        $pwdString = "Dej"+ $user.startDate
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
            $Name = $account +" "+ $ChineseName
            $DisplayName = $account +" "+ $ChineseName

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
    catch { 
        $ErrMsg = $_.Exception.Message
        Add-Content $ErrorLog $ErrMsg
    }
    
}
Write-Host "Waiting ....... "
sleep 120

$retryCount = 30
$retryIntervalSeconds = 60

$user = $null


while ($user -eq $null -and $retryCount -gt 0) {
    # search new user
    $user = Get-ADUser -Identity $account
    
    if ($user -eq $null) {
        $retryCount--
      
        Start-Sleep -Seconds $retryIntervalSeconds
    }
}

if ($user -ne $null) {
    Write-Host "User found: $account"

    Write-Host "Creating TEL ......"
    ForEach ($user In $source) { 
        if($account) { 
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
            $pin = $sitecode + $ext
            $line_137 = $privateline + ';ext=' + $ext
            $line_138 = $privateline + ';ext=' + $ext
        
    
            try
            {  
                Enable-CsUser -Identity $account -RegistrarPool 'twsfbpool01.deltaww.com' -SipAddressType SamAccountName -SipDomain deltaww.com -DomainController TWTPEDCS02
                
    
                if ($sitecode -eq '137') {
                    Set-CsUser -Identity $account -EnterpriseVoiceEnabled $true -LineURI $line_137 -DomainController TWTPEDCS02
                    Grant-CsDialPlan -Identity $account -PolicyName $dposaka -DomainController TWTPEDCS02
                    Grant-CsVoicePolicy -Identity $account -PolicyName $vposaka -DomainController TWTPEDCS02
                    sleep 60
                    Set-CsClientPin -Identity $account -pin '123456'
                   
                }
                elseif  ($sitecode -eq '138') {
                    Set-CsUser -Identity $account -EnterpriseVoiceEnabled $true -LineURI $line_138 -DomainController TWTPEDCS02
                    Grant-CsDialPlan -Identity $account -PolicyName $dpseoul -DomainController TWTPEDCS02
                    Grant-CsVoicePolicy -Identity $account -PolicyName $vpseoul -DomainController TWTPEDCS02
                    sleep 60
                    Set-CsClientPin -Identity $account -pin '123456'
                    
                }
            
                else {
                    Set-CsUser -Identity $account -EnterpriseVoiceEnabled $true -LineURI $line -PrivateLine $privateline -DomainController TWTPEDCS02
                    Grant-CsDialPlan -Identity $account -PolicyName $dptokyo -DomainController TWTPEDCS02
                    Grant-CsVoicePolicy -Identity $account -PolicyName $vptokyo -DomainController TWTPEDCS02
                    sleep 60
                    Set-CsClientPin -Identity $account -pin '123456'
                   
                    }
                #       Set-CsBusyOptions -Identity $sam -ActionType BusyOnBusy -Confirm:$false
    
                Add-DistributionGroupMember -Identity "jpdej@deltaww.com" -Member "$account@deltaww.com"
                Add-ADGroupMember -Identity "JPSG_ALL" -Members $account 
                Add-ADGroupMember -Identity "L-JP-SSLVPN" -Members $account 
    
                if ($account[0] -ne 'V') 
                {
                    Get-Mailbox -Identity $account| Set-Mailbox -CustomAttribute10 "o365new"
                }
                
                Write-Host "Account created successfully ! "
            }
            
        
            
        catch
        {
          $ErrMsg = $_.Exception.Message
          Add-Content $ErrorLog $ErrMsg
        } 
        }
     }
      


} else {
    Write-Host "User not found."
}







#get-pssession | remove-pssession