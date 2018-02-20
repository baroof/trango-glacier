<#
.SYNOPSIS
    reports last rebooted time from all machines under the \domain\org\systems OU

.DESCRIPTION
    . test connection to machine, report "OFFLINE" with darkgray if not pingable
    . if pingable, gather last reboot time
    . report "OK" with darkgreen bg if rebooted < 24 hours ago, otherwise "NOTOK" with darkred bg
    . report ERROR with darkyellow bg if pingable but no data (likely means non-windows)
    . group by OU
    . email results / print to html file

.PARAMETER days
    how many days back is "OK"?

.PARAMETER ou
    specify an OU under \domain\org\systems on which to focus
    defaults to ALL if not specified
    can get at sub-OUs like this: -ou "Terminal Servers,OU=ProductionServers"

.PARAMETER mailto
    where should results go?

.PARAMETER htmlfile
    filename to which results shall be printed, and to which the current date is added ("file.html" -> "file.20171008.html")

.OUTPUTS
    report results to html file, and email that file out if specified

.NOTES
  Author:         Will Mooreston
  Creation Date:  2017/08/25

.EXAMPLE
    .\Get-ServerRebootTimes.ps1 -days 2 -ou "Terminal Servers,OU=ProductionServers"

    this writes a report to the screen for all servers under the 
    "\domain\org\systems\ProductionServers\Terminal Servers" OU, 
    where a reboot in the past 2 days is OK   

.EXAMPLE
    .\Get-ServerRebootTimes.ps1 -days 4 -htmlfile C:\Files\ServerRebootTimesFull.html -mailto me@domain.org,me2@domain.org

    this writes a report to C:\Files\ServerRebootTimesFull[datetime].html 
    for all servers found under the \domain\org\systems OU and all sub-OUs, 
    where anything rebooted in the past 4 days is "OK", then emails the results
     to the specified email addresses

.EXAMPLE
    .\Get-ServerRebootTimes.ps1 -days 4 -htmlfile C:\Files\ServerRebootTimesProductionServers.html -ou "ProductionServers" -mailto me@domain.org,me2@gmail.com

    same as above, but restricted to the "\domain\org\systems\ProductionServers" OU and all sub-OUs
#>

param(
    [string]
    $days = 1
,
    [string]
    $ou = "all"
,
    [string]
    $mailto
,
    [string]
    $mailfrom
,
    [string]
    $htmlfile
)

##### ##### ##### ##### ##### ##### ##### ##### ##### ##### ##### #####
# command line verification
##### ##### ##### ##### ##### ##### ##### ##### ##### ##### ##### #####
if ($mailto -and -not $htmlfile) {
    Write-Error "ERROR: -mailto requires -htmlfile."
    exit
}

if ($htmlfile -and -not (Split-path $htmlfile | Test-Path)) {
    Write-Error "ERROR: parent directory of -htmlfile does not exist: $htmlfile"
    exit
}

"... starting run ..."

##### ##### ##### ##### ##### ##### ##### ##### ##### ##### ##### #####
# set up OUs, filtering out "Systems" top level folder (which we expect to be empty)
##### ##### ##### ##### ##### ##### ##### ##### ##### ##### ##### #####
if ($ou -eq "all") {
    $searchBase = "OU=Systems,DC=domain,DC=org"
} else {
    $searchBase = "OU="+$ou+",OU=Systems,DC=domain,DC=org"
}
$system_ous = Get-ADOrganizationalUnit -SearchBase $searchBase -Filter "*" | Where-Object {$_.Name -notlike "Systems"}  | Sort-Object

##### ##### ##### ##### ##### ##### ##### ##### ##### ##### ##### #####
# Gather server info and combine with OUs into custom objects
##### ##### ##### ##### ##### ##### ##### ##### ##### ##### ##### #####
$ServerInfo = @()
foreach ($ou_obj in $system_ous) {
    Write-Host "`n..."  #feedback to know that something is still happening...
    $ou_name = $ou_obj.Name
    $ou_dn = $ou_obj.DistinguishedName

    $ou_computers = Get-ADComputer -SearchBase $ou_dn -SearchScope OneLevel -Filter '*' | Where-Object {$_.DistinguishedName -like "*$ou_name*"} | Sort-Object

    foreach ($computer in $ou_computers) {
        Write-Host "." -NoNewline #feedback to know that something is still happening...

        # make a custom object for each machine
        $system = New-Object -TypeName PSObject

        $ComputerName = $computer.Name
                
        if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {

            try{
                $computerOS = get-wmiobject Win32_OperatingSystem -Computer $ComputerName -ErrorAction SilentlyContinue
            } catch{
            }

            if ($computerOS) {
                $rebooted =  $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
            } else {
                $rebooted = "ERROR"
            }            

            if ($rebooted -eq "ERROR") {
                $bgcolor = "darkyellow"
                $status = "ERROR"
            } elseif ($rebooted -le (Get-Date).AddDays(-$days)) {
                $bgcolor = "darkred"
                $status = "NOTOK"
            } elseif ($rebooted -ge (Get-Date).AddDays(-$days)) {
                $bgcolor = "darkgreen"
                $status = "OK"
            }

        } else {
            $rebooted = "OFFLINE"
            $bgcolor = "darkgray"
            $status = "OFFLINE"
        }

        $system | Add-Member -Type NoteProperty -Name OU -Value $ou_name
        $system | Add-Member -Type NoteProperty -Name Computer -Value $ComputerName
        $system | Add-Member -Type NoteProperty -Name Rebooted -Value $rebooted
        $system | Add-Member -Type NoteProperty -Name BGColor -Value $bgcolor
        $system | Add-Member -Type NoteProperty -Name Status -Value $status
        $ServerInfo += $system
    }
}

##### ##### ##### ##### ##### ##### ##### ##### ##### ##### ##### #####
# Generate HTML report
##### ##### ##### ##### ##### ##### ##### ##### ##### ##### ##### #####
if ($htmlfile) {
    # add datestamp to file name
    $datestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $htmlfile = $htmlfile -replace 'html',($datestamp+".html")

    $Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

    $PreContent = @"
<p>This report was generated at $datestamp.</p>
<p>OK/NOTOK = server has or has not been rebooted in the past $days days.</p>
<p>OFFLINE = server appears to be offline</p>
<p>ERROR = server does not respond to WMI request and is likely Non-Windows.</p>

"@

    ($ServerInfo | ConvertTo-Html -Property Status,OU,Computer,Rebooted -Head $Header -PreContent $PreContent) `
        -replace '<tr><td>OFFLINE</td>','<tr style="background-color:gray"><td>OFFLINE</td>' `
        -replace '<tr><td>NOTOK</td>','<tr style="background-color:red"><td>NOTOK</td>' `
        -replace '<tr><td>OK</td>','<tr style="background-color:green"><td>OK</td>' `
        -replace '<tr><td>ERROR</td>','<tr style="background-color:yellow"><td>ERROR</td>'|
        Out-File -FilePath $htmlfile

} else {
    # print results to the screen

    "" # new line to make room after progress dots
    foreach ($server in $ServerInfo) {
        $string = $server.Status + "|" + $server.OU + " | " + $server.Computer + " | " + $server.Rebooted
        write-host $string -BackgroundColor $server.BGcolor
    }
}

##### ##### ##### ##### ##### ##### ##### ##### ##### ##### ##### #####
# email the results
##### ##### ##### ##### ##### ##### ##### ##### ##### ##### ##### #####
if ($mailto) {
    Send-MailMessage -to $mailto -from $mailfrom -Subject "Report: Server Reboot Times" -SmtpServer outlook.domain.org -Attachments $htmlfile
}
