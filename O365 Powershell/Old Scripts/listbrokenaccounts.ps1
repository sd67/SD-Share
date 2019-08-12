Remove-PSSession $ExOPSession
Remove-PSSession $ExchangeSession
Remove-Module ActiveDirectory

#Go 
# 0 = Test
# 1 = Active
# 2 = Process $testnum accounts

# !!! READ AND UNDERSTAND THIS SCRIPT BEFORE YOU PUT IT INTO ACTIVE MODE !!! 
# !!! THIS SCRIPT COULD DELETE DATA !!!

$Go = 1
$user = ''
$SearchBase = 'OU=Elementary Schools,DC=mrpm,DC=sd42,DC=ca'
#number of accounts to test
$testnum = 1
$reset = 0

if ($UserCredential -eq $null){ 
        $UserCredential = new-object Management.Automation.PSCredential "<username>",$(ConvertTo-SecureString "<password>" -AsPlainText -Force)
    }#End IF no credential

#exchange Online
$ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

Import-PSSession $ExOPSession

echo 'Get Exchange Online Users'
$OnlineUsers = get-mailbox -resultsize Unlimited -Filter {WindowsEmailAddress -like '*_*'} #-properties windowsemailaddress
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
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<exchangeserver>/PowerShell/ -Authentication Kerberos 
Import-PSSession $ExchangeSession -AllowClobber

foreach ($OnPremEmail in $OnPremEmails){
    if ($OnPremEmail -in $OnlineEmails){

     
        $User = Get-Mailbox $OnPremEmail
  #  $SearchBase
  #   $User.DistinguishedName 
    
     if ($User.DistinguishedName -like $('*'+$SearchBase+'*')){    
       $OnpremEmail
      
       #break
       }

}}
Remove-PSSession $ExchangeSession

