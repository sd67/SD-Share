Remove-PSSession $exchangeSession
Remove-PSSession $365exchangeSession
Start-Transcript -Path C:\Office365Logs\fixStudentMail$(Get-Date -Format FileDateTime).txt

import-module activedirectory


#Prep Powershell Sessions
<#
if ($UserCred -eq $null){
    
    $UserCred = new-object Management.Automation.PSCredential "scripting user",$(ConvertTo-SecureString "password" -AsPlainText -Force)
    }

$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://onpremserver/PowerShell/ -Credential $UserCred -Authentication Kerberos 
Import-PSSession $ExchangeSession -Prefix OnPrem -AllowClobber

if ($CloudCred -eq $null){
    
    $CloudCred = new-object Management.Automation.PSCredential "scripting user",$(ConvertTo-SecureString "at4WE89y65T08w46" -AsPlainText -Force)
    }

$365exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $CloudCred -Authentication Basic -AllowRedirection
Import-PSSession $365exchangeSession -Prefix O365 -AllowClobber
#>

#Loop on this many accounts to fix timeouts
$loopstart = 1000

$noremote = 0
$noremotelocal = 0
$remotenolocal = 0
$nomail = 0
$mailok = 0
$loop = $loopstart
$count = 0

(Get-adgroup -identity 'O365_License_StudentFull' -properties members).members | Get-ADUser -properties userPrincipalName | ForEach-Object {
    if ($loop -ge $loopstart) {
        Remove-PSSession $ExchangeSession
        Remove-PSSession $365exchangeSession
        $loop = 0
        if ($UserCred -eq $null){
    
            $UserCred = new-object Management.Automation.PSCredential "script.exchange@mrpm.sd42.ca",$(ConvertTo-SecureString "OSQHAhRicwKA5vbO" -AsPlainText -Force)
        }

        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://deo-ex1301.mrpm.sd42.ca/PowerShell/ -Credential $UserCred -Authentication Kerberos 
        Import-PSSession $ExchangeSession -Prefix OnPrem -AllowClobber

        if ($CloudCred -eq $null){
    
            $CloudCred = new-object Management.Automation.PSCredential "script@schooldistrict42.onmicrosoft.com",$(ConvertTo-SecureString "at4WE89y65T08w46" -AsPlainText -Force)
        }

        $365exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $CloudCred -Authentication Basic -AllowRedirection
        Import-PSSession $365exchangeSession -Prefix O365 -AllowClobber
    }

    $onpremmail = $null
    $onpremremote = $null
    $o365mail = $null

    $onpremmail = $(get-onpremmailbox -identity $_.UserPrincipalName)
    $onpremremote = $(get-onpremremotemailbox -identity $_.UserPrincipalName)
    $o365mail = $(get-o365mailbox -identity $_.UserPrincipalName)

    #If Student Mailbox is On Prem and not in cloud
    if ($onpremremote -eq $null -and $onpremmail -ne $null -and $o365mail -eq $null) {
        $noremote += 1
        Write-Output "Migrating $_.UserPrincipalName"
        New-O365MoveRequest -Identity $_.UserPrincipalName -Remote -RemoteHostName "outlook.sd42.ca" -TargetDeliveryDomain "schooldistrict42.mail.onmicrosoft.com" -RemoteCredential $UserCred
    }
    Else {
    #If Student Mailbox is not in cloud but onprem thinks it is
    if ($onpremremote -ne $null -and $onpremmail -eq $null -and $o365mail -eq $null) {
        $noremotelocal += 1
        Write-Output "Removing Remote Link and creating new mailbox for $_.UserPrincipalName"
        #remove on prem remote link
        Disable-onpremRemoteMailbox -Identity $_.UserPrincipalName -Confirm:$false

        #enable remote mailbox
        Enable-onpremRemoteMailbox -Identity $_.UserPrincipalName -RemoteRoutingAddress $($_.SamAccountName+'@me.sd42.ca')
    }
    Else {
    #If student mailbox is not on prem but is in cloud
    if ($onpremremote -eq $null -and $onpremmail -eq $null -and $o365mail -ne $null) {
        $remotenolocal += 1
        Write-Output "Removing Remote Link and creating new mailbox for $_.UserPrincipalName"
        #enable remote mailbox
        Enable-onpremRemoteMailbox -Identity $_.UserPrincipalName -RemoteRoutingAddress $($_.SamAccountName+'@me.sd42.ca')
    }
    Else {
    #If student mailbox doesn't exist anywhere
    if ($onpremremote -eq $null -and $onpremmail -eq $null -and $o365mail -eq $null) {
        $nomail += 1
        Write-Output "Creating new mailbox for $_.UserPrincipalName"
        #enable remote mailbox
        Enable-onpremRemoteMailbox -Identity $_.UserPrincipalName -RemoteRoutingAddress $($_.SamAccountName+'@me.sd42.ca')
    }
    Else {
        Write-Output "Mailbox OK for $_.UserPrincipalName"
        $mailok += 1
    }
    $loop += 1
    $count += 1
    Write-Output "Processed Count: $count"
    }
    }
    }
}



Write-Output 'Migrated Mailbox: $noremote'
Write-Output 'Missing Remote Mailbox: $remotenolocal'
Write-Output 'No Mailbox: $nomail'
Write-Output 'Mailbox OK: $mailok'

Stop-Transcript