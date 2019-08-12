Remove-PSSession $ExOPSession
Remove-PSSession $ExchangeSession
Remove-Module ActiveDirectory
Start-Transcript -Path C:\Office365Logs\usersToCloud$(Get-Date -Format FileDateTime).txt


#STEPS
#Set GO
#Set Dates
#Set OU
#Set batchname
#Run as admin


# !!! READ AND UNDERSTAND THIS SCRIPT BEFORE YOU PUT IT INTO ACTIVE MODE !!! 
# !!! THIS SCRIPT COULD DELETE DATA !!!

##configure variables
#Set Operation
# 0 = Test
# 1 = Production
$Go = 0

#Set Start Date
$startdate = "2018-11-20"
#Set Complete Date 
#Ex. UTC "2016-05-06 14:30:00z"
#Updated to use local time
$CompleteDate = Get-Date("2018-11-27 01:00:00")


#Set OU
#Ex. 'OU='
$OU ="OU= put OU here"
#Set batchname
$batchname = '<batch name>'
$groupname = $batchname

if ($CloudCred -eq $null){
    
    $CloudCred = new-object Management.Automation.PSCredential "cloud user",$(ConvertTo-SecureString "pwd" -AsPlainText -Force)
    }

#Get User Credentials
if ($UserCred -eq $null){
    
    $UserCred = new-object Management.Automation.PSCredential "exchange user",$(ConvertTo-SecureString "pwd" -AsPlainText -Force)
    }
#Path for CSV
$filepath =  "C:\migrationCSVs"

#Connect to ExchangeOnline
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $CloudCred -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber

#connect to O365
Install-Module MSOnline -force
Connect-MsolService -Credential $CloudCred

#Connect to Active Directory
Import-Module ActiveDirectory




$Schools = Get-ADOrganizationalUnit -SearchBase $OU -SearchScope OneLevel -Filter * | 
     Select-Object DistinguishedName, Name

foreach ($school in $schools){
$searchbase = $school.DistinguishedName


#Get Users
echo 'get users into $users'
$Users = get-aduser -SearchBase $SearchBase -Filter {Title -ne 'Student'} -Properties *
$count = $users.count
$i = 1

#Turn on EmailAddress Policy
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<exchangeserver>/PowerShell/ -Authentication Kerberos -Credential $UserCred
Import-PSSession $ExchangeSession -AllowClobber

foreach ($User in $Users){
$user.UserPrincipalName
echo "$i of $count"
$i++
$usermailbox = get-user $user.UserPrincipalName
if ($usermailbox.RecipientType -eq 'UserMailbox'){
    If (get-aduser -Identity $User -Properties ProxyAddresses | Where {$_.proxyAddresses} | Where {-not($_.proxyAddresses -like "*@schooldistrict42.mail.onmicrosoft.com")}){
        
        if ($go -eq 1){
        echo 'enable email address policy'
            Set-Mailbox -Identity  $User.Name -EmailAddressPolicyEnabled $true}
        Else{
         echo 'test enable email address policy'   
            #get-mailbox -Identity $user.Name
        }
        
    }

#Send Email
$Who = get-mailbox $user.Name 
echo 'configure mail message'
if ($go -EQ 0){
    $toEMail = 'emailme@schooldistrict.ca'
}
Else {
    $toEMail = '"' +$who.DisplayName +'" <' +$who.WindowsEmailaddress +'>'
}

$from = 'from the helpdesk'
$BCC = 'who to bcc'
$SMTPServer = 'smtp server'
$subject = 'Email migration notice' 
$body = '

Hi there'$Who.DisplayName'
We are moving your email to Office 365 on '$CompleteDate'
Here are some instructions http://linktoinstructions.com

Your friendly Technology Team.
'



if ($go -eq 1){#SEND ALL MESSAGES
echo 'send real mail'
Send-MailMessage -From $from -to $toEMail -Bcc $bcc -Subject $Subject -Body $Body -SmtpServer $SMTPServer -BodyAsHtml     
}
}
}


if ($go -eq 0){ #Send One Message
echo 'test message'
Send-MailMessage -From $from -to $toEMail -Bcc $bcc -Subject $Subject -Body $Body -SmtpServer $SMTPServer -BodyAsHtml  
}

Remove-PSSession $ExchangeSession
Import-PSSession $Session -AllowClobber



$i = 1
#License all Users
foreach($user in $users)
{
    echo "$i of $count"
    $i++
    $user.UserPrincipalName
    $user.title

    #Create CSV for Migration
    echo 'Add to CSV'
    get-user -Identity $User.UserPrincipalName | Where-Object {$_.RecipientType -eq "MailUser"} | select-object @{Expression={$_.windowsemailaddress};Label="EmailAddress"},@{Expression={"PrimaryAndArchive"};Label="MailboxType"},@{Expression={"Unlimited"};Label="BadItemLimit"},@{Expression={"Unlimited"};Label="LargeItemLimit"} | export-csv $filepath\$batchname"_Batch.csv" -NoTypeInformation -append
    $user2 = get-user -Identity $User.UserPrincipalName
    }

}
  
#Migrate Email

$filename = $filepath+"\"+$batchname+"_Batch.csv"

#migration endpoint
echo 'create migration endpoint'
$MigrationEndpointOnPrem = Get-MigrationEndpoint -Identity OnpremEndpoint

if ($go -eq 0){
Write-Output "Start Migration (FAKE TEST RUN)"

}
ELSE{
echo 'start migration'
New-MigrationBatch -Name $GroupName -SourceEndpoint $MigrationEndpointOnprem.Identity -StartAfter $startdate -CompleteAfter $CompleteDate.ToUniversalTime() -NotificationEmails 'emailme@schooldistrict.ca' -TargetDeliveryDomain tenantname.mail.onmicrosoft.com -CSVData ([System.IO.File]::ReadAllBytes($filename)) -AllowUnknownColumnsInCsv $true
#Start-MigrationBatch -Identity $OnboardingBatch.Identity 
get-migrationbatch -identity $groupname
}

Remove-PSSession $Session
Stop-Transcript