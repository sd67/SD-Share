Remove-PSSession $ExOPSession
Remove-PSSession $ExchangeSession
Remove-Module ActiveDirectory

#to use this script
# If you leave this script in Test mode it will output all the accounts that need to be fixed. 
# You can then use the fixonemailbox script to process the broken accounts one at a time.
# If you put the script into Active mode it will fix all the broken accounts, this takes
# a very long time and may time out.
#
# 1. in line 26 enter the OU
# 2. in lines 32 enter credentials for Exchange On-Prem $usercredential and Exchange Online $livecred
# 2. in line 48+83 enter your exchange on-prem servername
# 3. in line 100 enter a local folder to store your pst backups
# 4. in line 125 enter your O365 tenant name

# !!! READ AND UNDERSTAND THIS SCRIPT BEFORE YOU PUT IT INTO ACTIVE MODE !!! 
# !!! THIS SCRIPT WILL DELETE DATA !!!

#Go 
# 0 = Test
# 1 = Active
# 2 = Process $testnum accounts

$Go = 0
$user = ''
$SearchBase = 'OU=<target OU for Users>'
#number of accounts to test
$testnum = 1
$reset = 0

if ($UserCredential -eq $null){ 
        $UserCredential = new-object Management.Automation.PSCredential "<admin account>",$(ConvertTo-SecureString "password" -AsPlainText -Force)
    }#End IF no credential

#exchange Online
$ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

Import-PSSession $ExOPSession

echo 'Get Exchange Online Users'
$OnlineUsers = get-mailbox -Filter {WindowsEmailAddress -like '*_*'} #-properties windowsemailaddress
#do some stuff

Remove-PSSession $ExOPSession


#Exchange On-prem
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<exchangeserver>/PowerShell/ -Authentication Kerberos 
Import-PSSession $ExchangeSession -AllowClobber

echo 'Get Exchange On-Prem Users'
$OnPremUsers = get-mailbox -Filter {WindowsEmailAddress -like '*_*'} -ResultSize Unlimited

Remove-PSSession $ExchangeSession


$OnpremEmails = New-Object System.Collections.Generic.List[System.Object]
$OnlineEmails = New-Object System.Collections.Generic.List[System.Object]

foreach ($OnPremUser in $OnPremUsers){

$OnPremEmails.add($OnPremUser.WindowsEmailAddress)

#$OnPremUser.WindowsEmailAddress
  
}#End foreach ($OnPremUser in $OnPremUsers)

foreach ($OnlineUser in $OnlineUsers){

$OnlineEmails.add($OnlineUser.WindowsEmailAddress)

#$OnPremUser.WindowsEmailAddress
  
} #End foreach ($OnlineUser in $OnlineUsers)

echo 'Start Compare'

foreach ($OnPremEmail in $OnPremEmails){
    if ($OnPremEmail -in $OnlineEmails){

     $OnpremEmail
    
     $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<exchangeserver>/PowerShell/ -Authentication Kerberos 
     Import-PSSession $ExchangeSession -AllowClobber
     
     echo $OnPremEmail ' Broken Account'
     $User = Get-Mailbox $OnPremEmail
     
     $SearchBase
     $User.DistinguishedName 
    
     if ($User.DistinguishedName -notlike $('*'+$SearchBase+'*')){    
       $go = 0
       echo 'set $go = 0'
       $reset = 1
       echo 'not in correct OU'
       #break
       }
     if ($go -ge 1){#IF Script is Active
     $ExportMailbox = New-MailboxExportRequest -Mailbox $OnPremEmail -FilePath \\<filepath>\$User.pst

     $testmbexport = get-mailboxexportrequest $ExportMailbox.Identity

     if ($testmbexport.Status -notlike 'Completed'){
     do {
     echo "$($User.Name) Export $($testmbexport.Status)"
    
     echo 'wait 30 seconds'
     Start-Sleep -s 30
    
     
     $testmbexport = get-mailboxexportrequest $ExportMailbox.Identity
     }while ($testmbexport.status -ne 'Completed') #End while loop

     }#END IF Status is InProgress

     if ($testmbexport.Status -eq 'Completed'){
        Remove-MailboxExportRequest $testmbexport.Identity -Confirm:$false
        
        if ($go -ge 1){#IF Script is Active
            
            echo 'Disable On-Prem Mailbox'
            echo 'Enable Remote Mailbox'
            Disable-Mailbox -Identity $User.Identity -confirm:$false
            Enable-RemoteMailbox -Identity $User.Identity -RemoteRoutingAddress $($user.SamAccountName+'@tenantname.mail.onmicrosoft.com')
           #break      
        }ELSE{
            echo 'Disable On-Prem Mailbox'
            echo 'Enable Remote Mailbox'
            
        }
       
        
     }#EndIf $testmbexport is completed

     }ELSE{#END IF Script IS Active
     echo 'testing'
     }
   ##  } #End if ($OnPremEmail -in $OnlineEmails) 
     
  Remove-PSSession $ExchangeSession   
  $ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

  Import-PSSession $ExOPSession
  
  $testnum = $testnum - 1
  $CloudMB = get-mailbox $User.SamAccountName
  $ExchangeGUID = $CloudMB.ExchangeGuid
  $ArchiveGUID = $CloudMB.ArchiveGuid
  #Break
Remove-PSSession $ExOPSession 
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<exchangeserver>/PowerShell/ -Authentication Kerberos 
Import-PSSession $ExchangeSession -AllowClobber

echo 'Set GUIDS'

if ($go -ge 1){#IF Script is Active
    Set-RemoteMailbox -Identity $User.Identity -ExchangeGuid $ExchangeGUID.Guid -ArchiveGuid $ArchiveGUID.Guid
    #Break
}


Remove-PSSession $ExchangeSession

#Connect to Active Directory
Import-Module ActiveDirectory

$currentuser = get-ADUser $user.SamAccountName -Properties msExchRemoteRecipientType
#Break
if ($currentuser.msExchRemoteRecipientType -ne 4){
    echo 'Set msExchRemoteRecipientType to 4'
    if ($go -ge 1){#IF Script is Active
        Set-ADUser $User.samaccountname -replace @{msExchRemoteRecipientType = 4}
        #Break
    }

}ELSE{ #End ($currentuser.msExchRemoteRecipientType -ne 4)
 echo 'msExchRemoteRecipientType Already Set'
}#End ($currentuser.msExchRemoteRecipientType -ne 4) ELSE


Remove-Module ActiveDirectory
#Email PST Backup


#Remove PST Backup
 if ($reset -eq 1){    
       $go = 1
       $reset = 0
       echo 'Reset $go to 1'
      echo 'set $reset to 0'
       }
    }Else{ # End if ($OnPremEmail in $OnlineEmails)
   #echo $OnPremEmail ' No Issue'
    }#End if ($OnPremEmail in $OnlineEmails) ELSE

}


