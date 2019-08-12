
#Set Complete Date
$CompleteDate = "9/04/2018"
#Set Start Date
$startdate = "08/28/2018"

$batchname = 'AR_AllUsers'
$groupname = $batchname
$filepath =  "C:\users\user1\desktop\migrationCSVs"
$filename = $filepath+"\"+$batchname+"_Batch.csv"

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCred -Authentication Basic -AllowRedirection
Import-PSSession $Session

$MigrationEndpointOnPrem = Get-MigrationEndpoint -Identity OnpremEndpoint
New-MigrationBatch -Name $GroupName -SourceEndpoint $MigrationEndpointOnprem.Identity -StartAfter $startdate -CompleteAfter $CompleteDate -NotificationEmails 'notify@sd42.ca' -TargetDeliveryDomain schooldistrict42.mail.onmicrosoft.com -CSVData ([System.IO.File]::ReadAllBytes($filename)) -AllowUnknownColumnsInCsv $true

Remove-PSSession $Session
