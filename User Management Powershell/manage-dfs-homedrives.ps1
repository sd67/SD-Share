##------------------------------------------------------
## manage-dfs-homedrives.ps1
## sets the homedrives and creates the DFS path and folder 
## for users who do not have homedrives.
##------------------------------------------------------
## original script matt williams 2019-08-12
##------------------------------------------------------

$Users = Get-ADUser -Filter {(Enabled -eq "True") -and ((EmployeeType -eq "Staff") -or (EmployeeType -eq "Student"))} -Properties * | Where-Object {$_.homedirectory -eq $null}

$Domain = '<domain>'
ForEach ($User in $Users){
    if ($User.employeeType -eq 'Staff'){
        #if it is staff set the path
        if ($User.idautoPersonPriLocCode -eq 6715003){
            #if it is Naramata set the path to this:
            #$DFSPath = "\\$Domain\Users\Staff\$User.name"
            #$FSPath = "\\$Domain\Users\Staff\$User.name"
        }Else{
            #if it is not Naramata set the path to this:
            $DFSPath = "\\$Domain\Users\Staff\$User.name"
            $FSPath = "\\$Domain\Users\Staff\$User.name"
        }
    }
    if ($User.employeeType -eq 'Student'){
        #if it is student set the path
        if ($User.idautoPersonPriLocCode -eq 6715003){
            #if it is Naramata set the path to this:
            #$DFSPath = "\\$Domain\Users\Students\$User.name"
            #$FSPath = "\\$Domain\Users\Students\$User.name"
        }Else{
            #if it is not Naramata set the path to this:
            $DFSPath = "\\$Domain\Users\Students\$User.name"
            $FSPath = "\\$Domain\Users\Students\$User.name"
        }
    }
    try {
        Get-DfsnFolderTarget -Path $DFSPath -ErrorAction Stop
    } 
    catch {
        Write-Host "Path not found. Clear to proceed" -ForegroundColor Green
    
        $NewDFSFolder = @{
            Path = $DFSPath
            State = 'Online'
            TargetPath = $FSPath
            TargetState = 'Online'
            ReferralPriorityClass = 'globalhigh'
            }

        New-DfsnFolder @NewDFSFolder

        # Check that folder now exists:
        Get-DfsnFolderTarget -Path $DFSPath

        # Check that the new DFS Link works using Windows Explorer
        Invoke-Expression "explorer $DFSPath"
    }

}






