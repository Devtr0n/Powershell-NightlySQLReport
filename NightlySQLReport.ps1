############################# 1 - SNAP-IN SECTION ###################################################
#add SQL "snap-in" tools into PowerShell session
$snapinName = 'SqlServerCmdletSnapin100' # here lives Invoke-SqlCmd
$snapinAdded = Get-PSSnapin | Select-String $snapinName
if (!$snapinAdded)
{
    Add-PSSnapin $snapinName
}
 
$snapinName = 'SqlServerProviderSnapin100'
$snapinAdded = Get-PSSnapin | Select-String $snapinName
if (!$snapinAdded)
{
    Add-PSSnapin $snapinName
}
############################# 2 - SQL LOGGING SECTION ###################################################
#1.) declare variables for database configuration & logging
$username='xxx'
$password='xxx'
$serverinstance='xxx\xxx2008'
$database=’xxx’
$logfile='C:\Tools\NightlySQLScript-log.log'
$htmlFile='C:\Tools\NightlySQLScript-log.html'
 
#2.) create fresh log file
if (Test-Path $logfile)
{
    Clear-Content $logfile
} else {
    New-Item $logfile -ItemType file
}
"$(Get-Date -format F) - Start logging..." | Out-File -filePath $logfile -Append
 
#3.) try/execute SQL statements
"$(Get-Date -format F) - Executing SQL statement #1 - pr_egr_RPT_UpdateEgrantsHeaderData..." | Out-File -filePath $logfile -Append
try {
    Invoke-Sqlcmd "EXEC pr_egr_RPT_UpdateEgrantsHeaderData;" -ServerInstance $serverinstance -Username $username -Password $password -Database $database -Verbose 4>&1 | Out-File -filePath $logfile -Append
} catch {
    "SQL error occured - Executing SQL statement #1 - pr_egr_RPT_UpdateEgrantsHeaderData`n" + $_ | Out-File -filePath $logfile -Append
}
 
"$(Get-Date -format F) - Executing SQL statement #2 - pr_egr_RPT_UpdateEgrantsNetData_Grants_ConsRpt..." | Out-File -filePath $logfile -Append
try {
                Invoke-Sqlcmd "EXEC pr_egr_RPT_UpdateEgrantsNetData_Grants_ConsRpt;" -ServerInstance $serverinstance -Username $username -Password $password -Database $database -Verbose 4>&1 | Out-File -filePath $logfile -Append
} catch {
    "SQL error occured - Executing SQL statement #2 - pr_egr_RPT_UpdateEgrantsNetData_Grants_ConsRpt`n" + $_ | Out-File -filePath $logfile -Append
}
 
#4.) Close the log file
"$(Get-Date -format F) - End logging!!" | Out-File -filePath $logfile -Append
 
############################# 3 - HTML REPORT SECTION ###################################################
#5.) Build an HTML report based on the log file
$header = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">
<head>
<title>Nightly SQL Log Report</title>
<style type="text/css">
                body {
                background-color: #E0E0E0;
                font-family: sans-serif
                }
                table, th, td {
                background-color: white;
                border-collapse:collapse;
                border: 1px solid black;
                padding: 5px
                }
</style>
"@
 
$body = @"
<h1>Nightly SQL Report</h1>
<p>The following report was run on $(get-date -format F).</p>
"@
 
$htmlReport = @()
foreach ($line in (Get-Content $logfile)) {
 
                if ($line | Select-String 'Error','Exception','Parser',' at' -AllMatches -SimpleMatch)
                {
                                $status = "FAIL"
                }
                 else
                {
                                $status = "PASS"
                }
 
                $reportRow = New-Object -TypeName PSObject
                Add-Member -InputObject $reportRow -Type NoteProperty -Name Output -Value $line
                Add-Member -InputObject $reportRow -Type NoteProperty -Name Status -Value $status
                $htmlReport += $reportRow
}
$htmlReport | ConvertTo-Html -Property Output, Status -head $header -body $body | foreach {
    $PSItem -replace "<td>FAIL</td>", "<td style='background-color:#FF8080'>FAIL</td>" -replace "<td>PASS</td>", "<td style='background-color:green'>PASS</td>"
} | Out-File $htmlFile
 
Copy-Item $htmlFile 'C:\TEA\Tools\logs' #copy to IIS directory
 
############################# 4 - EMAIL SECTION #######################################################
#1.) Declare & initialize email variables
$msgto=’recipient@test.com'
$msgfrom='admin@test.com’
$smtpserver='smpt.server.com'
 
if ($htmlReport | Select-String 'Error', 'Insert')
{
                $msgbody='Hi eGrants Support, There were some errors during the nightly process. Please see the attached log file.' #log errors present
}
else
{
                $msgbody='Hi eGrants Support, The nightly SQL process ran successfully without any errors.'                                                                                                     #log is error free
}
 
#2.) Send email to support team with the logfile as an attachment
Send-MailMessage -SMTPServer $smtpserver -To $msgto -From $msgfrom -Subject "Nightly Scripts - SQL Report" -Body $msgbody -attachment $htmlFile
############################# END OF SCRIPT #########################################################
