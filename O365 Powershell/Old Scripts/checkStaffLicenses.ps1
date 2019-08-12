#Remove-PSSession $exchangeSession
#Remove-PSSession $365exchangeSession

# to use this script
# 1. In Line 11 enter your OU
# 2. in lines 17 enter credentials for Exchange Online $livecred
# 2. in lines 28+29 enter your tenant name

import-module activedirectory

$ou = "OU= put OU here"


#Prep Powershell Sessions

if ($CloudCred -eq $null){
    
    $CloudCred = new-object Management.Automation.PSCredential "useraccount",$(ConvertTo-SecureString "pwd" -AsPlainText -Force)
    }

Connect-MSOLService -Credential $CloudCred

$users = get-aduser -filter 'employeeType -like "Employee"' -searchbase $ou

$licensedusers = 0
$missingsublicense = 0
$unlicensedusers = 0
$count = 0
$A3licensesku = "<TENANTNAME>:ENTERPRISEPACKPLUS_FACULTY"
$sololicensesku = "<TENANTNAME>:EXCHANGEENTERPRISE_FACULTY"
$sublicensesku = "EXCHANGE_S_ENTERPRISE"


$users = get-aduser -filter 'employeeType -like "Employee"' -searchbase $ou


ForEach ($user in $users) {
    $count += 1
    $userisparsed = $false
    $licenselist = Get-MSOLUser -UserPrincipalName $user.userprincipalname| select-object -expandproperty licenses

    #parse licenses to see if they contain what we are looking for
    ForEach($licenses in $licenselist) {
        $licenses | Where-Object {($_.AccountSkuId -contains $A3licensesku) -or ($_.AccountSkuId -contains $sololicensesku)} | ForEach-Object {
            $license = $_
            $plan = Select-Object -InputObject $license -ExpandProperty servicestatus | where-object provisioningstatus -contains "Success" | select-object -expandproperty serviceplan | where-object servicename -contains $sublicensesku

            #report on license status
            if($plan -ne $null){
                $licensedusers += 1
                $userisparsed = $true
                Write-Output "$($user.userprincipalname) is licensed for $($sublicensesku) in $($license.AccountSkuId)."
                }
            else{
                $missingsublicense += 1
                $userisparsed = $true
                Write-Warning "$($user.userprincipalname) has $($license.AccountSkuId) license but is missing $($sublicensesku)."
                }
            }
        }
    if($userisparsed -ne $true){
        $unlicensedusers += 1
        Write-Warning "$($user.userprincipalname) is not licensed."
        }
    }


Write-Output $ou
Write-Output "Total Users $($count)"
Write-Output "Licensed Users $($licensedusers)"
Write-Output "Missing Sublicense $($missingsublicense)"
Write-Output "Unlicensed Users $($unlicensedusers)"