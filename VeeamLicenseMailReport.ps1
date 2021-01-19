write-host -ForegroundColor Green "Starting Script."


$smtpRecipients =  "email-user 1 <email.user.1@domain.local>", "email-user 2 <email.user.2@domain.local>"
$smtpFrom = 'Veeam Usage report <mail.from@domain.local>'
$smtpServer = "smtp.domain.local"
$base64Logo = ''


# Defined Veeam Servers
$hosts = @(
    #uncomment to use
    #("vbr.fqdn.domain.local", "Customer", "VBR: Enterprise"),
    #("vbo.domain.local", "Customer", "O365")
)

# clearing out variables
$results = @()
$results3 = @()

# Fetches data from Veeam Servers ($hosts array) through WinRM
write-host -ForegroundColor Green "Fetching license data from Veeam servers."
foreach ($server in $hosts) {


    If ($server[2] -eq "O365") {
        $o365Count = 0
        $o365 = New-Object PSObject |	
        Add-Member -type NoteProperty -Name 'PSComputerName' -Value $server[0] -PassThru |
        Add-Member -type NoteProperty -Name 'Type' -Value 'O365' -PassThru |
        Add-Member -type NoteProperty -Name 'Multiplier' -Value 'N/A' -PassThru |
        Add-Member -type NoteProperty -Name 'Source' -Value '' -PassThru			

        $o365Count = Invoke-Command -ScriptBlock { Import-Module Veeam.Archiver.PowerShell ; (Get-VBOLicensedUser | Where-Object { $_.LicenseState -eq 'Licensed' } | Measure-Object).Count } -ComputerName $server[0]
        $o365 | Add-Member -type NoteProperty -Name 'Count' -Value $o365Count 
        $o365 | Add-Member -type NoteProperty -Name 'UsedInstancesNumber' -Value $o365Count

        $results += $o365
    }
    else {

        write-host -ForegroundColor Green "Server: " -nonewline 
        write-host -ForegroundColor Yellow $server
        $results += Invoke-Command -ScriptBlock { Add-PSSnapin VeeamPSSnapin ; Connect-VBRserver -server localhost ; (Get-VBRInstalledLicense).InstanceLicenseSummary.Object } -ComputerName $server[0]

        <# write-host -ForegroundColor Red -BackgroundColor Yellow "Something failed" #>
    }
}



# Add Customer and License Model to array.
for ($i = 0 ; $i -lt $hosts.Count; $i++) { 
    $results | ForEach-Object {
        $_ | Where-Object -Property PSComputername -eq $hosts[$i][0] | Add-Member -MemberType NoteProperty -Name "Customer" -Value $hosts[$i][1]
        $_ | Where-Object -Property PSComputername -eq $hosts[$i][0] | Add-Member -MemberType NoteProperty -Name "License" -Value $hosts[$i][2]
    }
}

$results3 = $results | Select-Object PSComputerName, Customer, License, Type, Count, Multiplier, UsedInstancesNumber, Source

# Store data per license type and version
# VBR Standard
$StandardVM = $results3 | Where-Object { $_.License -eq 'VBR: Standard' -and $_.Type -eq 'VM' }
$StandardAgentServer = $results3 | Where-Objecthere-Object { $_.License -eq 'VBR: Standard' -and $_.Type -eq 'Server' }
$StandardAgentWorkstation = $results3 | Where-Object { $_.License -eq 'VBR: Standard' -and $_.Type -eq 'Workstation' }

# VBR Enterprise
$EnterpriseVM = $results3 | Where-Object { $_.License -eq 'VBR: Enterprise' -and $_.Type -eq 'VM' }
$EnterpriseAgentServer = $results3 | Where-Object { $_.License -eq 'VBR: Enterprise' -and $_.Type -eq 'Server' }
$EnterpriseAgentWorkstation = $results3 | Where-Object { $_.License -eq 'VBR: Enterprise' -and $_.Type -eq 'Workstation' }


# VBR Enterprise
$EnterprisePlusVM = $results3 | Where-Object { $_.License -eq 'VBR: Enterprise Plus' -and $_.Type -eq 'VM' }
$EnterprisePlusAgentServer = $results3 | Where-Object { $_.License -eq 'VBR: Enterprise Plus' -and $_.Type -eq 'Server' }
$EnterprisePlusAgentWorkstation = $results3 | Where-Object { $_.License -eq 'VBR: Enterprise Plus' -and $_.Type -eq 'Workstation' }

# VBR Cloud Connect
$EnterprisePlusCCVM = $results3 | Where-Object { $_.License -eq 'VBR: Enterprise Plus' -and $_.Type -eq 'CloudConnectBackupVM' }


# Store data per license type and version
# VAS Standard
$VASStandardVM = $results3 | Where-Object { $_.License -eq 'VAS: Standard' -and $_.Type -eq 'VM' }
$VASStandardAgentServer = $results3 | Where-Object { $_.License -eq 'VAS: Standard' -and $_.Type -eq 'Server' }
$VASStandardAgentWorkstation = $results3 | Where-Object { $_.License -eq 'VAS: Standard' -and $_.Type -eq 'Workstation' }

# VAS Enterprise
$VASEnterpriseVM = $results3 | Where-Object { $_.License -eq 'VAS: Enterprise' -and $_.Type -eq 'VM' }
$VASEnterpriseAgentServer = $results3 | Where-Object { $_.License -eq 'VAS: Enterprise' -and $_.Type -eq 'Server' }
$VASEnterpriseAgentWorkstation = $results3 | Where-Object { $_.License -eq 'VAS: Enterprise' -and $_.Type -eq 'Workstation' }


# VAS Enterprise
$VASEnterprisePlusVM = $results3 | Where-Object { $_.License -eq 'VAS: Enterprise Plus' -and $_.Type -eq 'VM' }
$VASEnterprisePlusAgentServer = $results3 | Where-Object { $_.License -eq 'VAS: Enterprise Plus' -and $_.Type -eq 'Server' }
$VASEnterprisePlusAgentWorkstation = $results3 | Where-Object { $_.License -eq 'VAS: Enterprise Plus' -and $_.Type -eq 'Workstation' }

# VAS Cloud Connect
$VASEnterprisePlusCCVM = $results3 | Where-Object { $_.License -eq 'VAS: Enterprise Plus' -and $_.Type -eq 'CloudConnectBackupVM' }

#O365 
$VeeamO365 = $results3 | Where-Object { $_.License -eq 'O365' }


write-host -ForegroundColor Green "Fetched data successfully!"
write-host -ForegroundColor Green "Writing output to console:..."
$results3 |   Format-Table -AutoSize
write-host -ForegroundColor Green "Writing output to console done."



write-host -ForegroundColor Green "Generating Email"


$year = (Get-Date -UFormat %Y)
$month = (Get-Date -UFormat %m)
$day = (Get-Date -UFormat %d)
$hour = (Get-date -UFormat %H)
$minute = (Get-date -UFormat %M)
$second = (Get-date -UFormat %S)
#$yearmonth = "$year-$month"
$yearmonthday = "$year-$month-$day"
$dateReportMonth = (get-date).AddMonths(-1).ToString('yyyy-MM')
$script = $MyInvocation.MyCommand.Name
$domain = [System.Environment]::UserDomainName
$username = [System.Environment]::UserName
$machine = [System.Environment]::MachineName
$path = (get-location).toString()



$spacer = @"
<br /><hr />
"@
$head = @"
<head>
<style type='text/css'>
body {
font-family: Verdana,Calibri;
font-size: x-small;
}
th {
background-color: #c02741;
color: white;
padding: 10px;
}

td {
border-bottom: 1px solid #ddd;
padding: 10px;
}

h3 {
font-family: Verdana,Calibri;
font-size: 36px;
}

h4 {
font-family: Verdana,Calibri;
font-size: 16px;
}
</style>
</head>
"@
#<img src='cid:logo.png'>
$header = @"

<img src="data:image/png;base64,$base64Logo">
<h3>Veeam License usage report</h3>
"@

$VBRData = @{
    "VBR: Standard - VM"                         = ($StandardVM | Measure-Object -Sum -Property Count).sum;
    "VBR: Standard - Agent - Server"             = ($StandardAgentServer | Measure-Object -Sum -Property Count).sum;
    "VBR: Standard - Agent - Workstation"        = ($StandardAgentWorkstation | Measure-Object -Sum -Property Count).sum;

    "VBR: Enterprise - VM"                       = ($EnterpriseVM | Measure-Object -Sum -Property Count).sum;
    "VBR: Enterprise - Agent - Server"           = ($EnterpriseAgentServer | Measure-Object -Sum -Property Count).sum;
    "VBR: Enterprise - Agent - Workstation"      = ($EnterpriseAgentWorkstation | Measure-Object -Sum -Property Count).sum;

    "VBR: Enterprise Plus - VM"                  = ($EnterprisePlusVM | Measure-Object -Sum -Property Count).sum;
    "VBR: Enterprise Plus - Agent - Server"      = ($EnterprisePlusAgentServer | Measure-Object -Sum -Property Count).sum;
    "VBR: Enterprise Plus - Agent - Workstation" = ($EnterprisePlusAgentWorkstation | Measure-Object -Sum -Property Count).sum;
    "VBR: Enterprise Plus - VM - Cloud Connect"  = ($EnterprisePlusCCVM | Measure-Object -Sum -Property Count).sum;
}


$VASData = @{
    "VAS: Standard - VM"                         = ($VASStandardVM | Measure-Object -Sum -Property Count).sum;
    "VAS: Standard - Agent - Server"             = ($VASStandardAgentServer | Measure-Object -Sum -Property Count).sum;
    "VAS: Standard - Agent - Workstation"        = ($VASStandardAgentWorkstation | Measure-Object -Sum -Property Count).sum;

    "VAS: Enterprise - VM"                       = ($VASEnterpriseVM | Measure-Object -Sum -Property Count).sum;
    "VAS: Enterprise - Agent - Server"           = ($VASEnterpriseAgentServer | Measure-Object -Sum -Property Count).sum;
    "VAS: Enterprise - Agent - Workstation"      = ($VASEnterpriseAgentWorkstation | Measure-Object -Sum -Property Count).sum;

    "VAS: Enterprise Plus - VM"                  = ($VASEnterprisePlusVM | Measure-Object -Sum -Property Count).sum;
    "VAS: Enterprise Plus - Agent - Server"      = ($VASEnterprisePlusAgentServer | Measure-Object -Sum -Property Count).sum;
    "VAS: Enterprise Plus - Agent - Workstation" = ($VASEnterprisePlusAgentWorkstation | Measure-Object -Sum -Property Count).sum;
    "VAS: Enterprise Plus - VM - Cloud Connect"  = ($VASEnterprisePlusCCVM | Measure-Object -Sum -Property Count).sum;
}

$O365Data = @{
    "O365 - Users" = ($VeeamO365 | Measure-Object -Sum -Property Count).sum;
}

# HTML Mail data
# Results to html
$bodyResults = $results3 |  ConvertTo-Html
$VBRsummary = ($VBRData.GetEnumerator() | Sort-Object -Property Name | ConvertTo-Html -Fragment -Property Name, Value)
$VASsummary = ($VASData.GetEnumerator() | Sort-Object -Property Name | ConvertTo-Html -Fragment -Property Name, Value)
$O365summary = ($O365Data.GetEnumerator() | Sort-Object -Property Name | ConvertTo-Html -Fragment -Property Name, Value)

# CSV Mail data
# Results to CSV (for attachment)
$results3 | Export-Csv  "$dateReportMonth-rawdata.csv" -Delimiter ";" -NoTypeInformation -force
# VBR Data 
($VBRData.GetEnumerator()) | Select-Object Name, Value | Export-Csv -NoTypeInformation -Force -Delimiter ";" $dateReportMonth-VBR-lic-summary.csv

# VAS Data 
($VASData.GetEnumerator()) | Select-Object Name, Value | Export-Csv -NoTypeInformation -Force -Delimiter ";" $dateReportMonth-VAS-lic-summary.csv

#O365 Data
($O365Data.GetEnumerator()) | Select-Object Name, Value | Export-Csv -NoTypeInformation -Force -Delimiter ";" $dateReportMonth-O365-lic-summary.csv

$footer = @"

Script $script executed at $year-$month-$day $hour`:$minute`:$second by $domain\$username from $machine @ path $path

"@

write-host -ForegroundColor Green "Sending Email..."
Send-MailMessage -From $smtpFrom  -To $smtpRecipients -Subject "Veeam Usage Report: $dateReportMonth ($yearmonthday)"  -Body "$head $header $bodyResults $spacer <h4>Veeam Backup & Replication (no Veeam ONE)</h4> $VBRsummary $spacer  <h4>Veeam Availability Suite (VBR+Veeam ONE)</h4> $VASsummary $spacer <h4>Veeam O365</h4> $O365summary <i>$footer</i>	" -SmtpServer $smtpServer  -BodyAsHtml -Attachments "$dateReportMonth-rawdata.csv", "$dateReportMonth-VBR-lic-summary.csv", "$dateReportMonth-VAS-lic-summary.csv", "$dateReportMonth-O365-lic-summary.csv", "logo.png"
write-host -ForegroundColor Green "Email sent."

#clean up
Remove-Item -Path ".\$dateReportMonth-rawdata.csv"
Remove-Item -Path ".\$dateReportMonth-VBR-lic-summary.csv"
Remove-Item -Path ".\$dateReportMonth-VAS-lic-summary.csv"
Remove-Item -Path ".\$dateReportMonth-O365-lic-summary.csv"

write-host -ForegroundColor Green "Script ended, you can now close this window."



