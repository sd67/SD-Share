##------------------------------------------------------
## move-disabled-users.ps1
## moves disabled users to the specified OU.
##------------------------------------------------------
## original script matt williams 2019-08-12
##------------------------------------------------------

$DisabledUsers = Search-ADAccount -AccountDisabled -SearchBase "OU=Staff,OU=67-Users,DC=sd67,DC=edu"

$count = 0
foreach ($DisabledUser in $DisabledUsers){
    #$DisabledUser.SamAccountName
    $count++
    Move-ADObject -Identity $DisabledUser -TargetPath "OU=Disabled,OU=67-Users,DC=sd67,DC=edu"
    #Set-ADUser -Identity $DisabledUser -Remove @{extensionAttibute15="Staff"} 
}
#Write-Output "Moved $count Users to Disabled OU"
$DisabledUsers = $null

$DisabledUser