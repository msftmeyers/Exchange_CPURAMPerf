<#
.SYNOPSIS
    CPU RAM Performance Snapshot/overview

.DESCRIPTION
    This script will show you the concurrent CPU and RAM data of every Exchange
    server in your organization.
    
    You can use the script to get a fast overview of CPU Cores and installed RAM
    and furthermore the concurrent utilization of both.
                                  |"
    The network availability of all servers will be checked first and
    only available servers data will be collected.
                                              
.PERMISSIONS
    You need to be at least a member of Exchange view-only and the local administrator
    group to run the script.

.ENVIRONMENTS
    Script is optimized to run in Exchange on premise environments, because it will only
    collect data from Exchange Servers.

.EXAMPLE
    .\cpuramperf.ps1

 .VERSIONS
    12.02.2025 V1.0 Initial version
    14.02.2025 V1.1 Minor fixes

.AUTHOR/COPYRIGHT: Steffen Meyer
.ROLE: Cloud Solution Architect
.COMPANY: Microsoft Deutschland GmbH

#>
$scriptversion = "V1.1_14.02.2025"

try
{
    $ScriptPath = Split-Path -parent $MyInvocation.MyCommand.Path -ErrorAction Stop
}
catch
{
    Write-Host "`nDo not forget to save the script!" -ForegroundColor Red
}

$now = Get-Date -Format G

#Check if Exchange SnapIn is available and load it
if (!(Get-PSSession).ConfigurationName -eq "Microsoft.Exchange")
{
    if ((Get-PSSnapin -Registered).name -contains "Microsoft.Exchange.Management.PowerShell.SnapIn")
    {
        Write-Host "`nLoading the Exchange Powershell SnapIn..." -ForegroundColor Yellow
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
        . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
        Connect-ExchangeServer -auto -AllowClobber
    }
    else
    {
        write-host "`nExchange Management Tools are not installed. Run the script on a different machine." -ForegroundColor Red
        Return
    }
}

#Detect, where the script is executed
if (!(Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue))
{
    write-host "`nATTENTION: Script is executed on a non-Exchangeserver...`n" -ForegroundColor Cyan
}

Write-Host "----------------------------------------------"
Write-Host "Current/Reference time stamp: $(Get-Date $now -Format "dd.MM.yyyy HH:mm")"
Write-Host "Script version: $scriptversion"
Write-Host "----------------------------------------------`n"

$srvrs = Get-ExchangeServer | sort name

#Check server availability
$servers = @()
$data = @()
$i = 0

foreach ( $srvr in $srvrs )
{
    #Progress bar
    $i++
    $status = "{0:N0}" -f ($i / $srvrs.count * 100)
    Write-Progress -Activity "Checking Server connectivity to $srvr..." -Status "Trying to reach server $i of $($srvrs.count) : $status% Completed" -PercentComplete ($i / $srvrs.count * 100)
    
    $testconnect = Test-Connection -ComputerName $srvr.fqdn -Quiet -Count 1
    
    if ($testconnect)
    {
        $winrm = Test-WSMan -ComputerName $srvr.fqdn -ErrorAction SilentlyContinue
        
        If ($winrm)
        {
            $servers += $srvr
        }
        else
        {
            Write-Host "$srvr WINRM NOT RUNNING`n" -ForegroundColor Red
            $data += New-Object -Type PSObject -Prop @{Servername=$srvr;Comment="WINRM NOT RUNNING"}
        }
    }
    else
    {
        $ip = (Test-NetConnection -ComputerName $srvr.fqdn -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).remoteaddress.ipaddresstostring
        if ($ip)
        {
            Write-Host "$srvr NOT REACHABLE`n" -ForegroundColor Cyan
            $data += New-Object -Type PSObject -Prop @{Servername=$srvr;Comment=$ip}
        }
        else
        {
            Write-Host "$srvr NOT RESOLVABLE`n" -ForegroundColor Red
            $data += New-Object -Type PSObject -Prop @{Servername=$srvr;Comment="NOT RESOLVABLE"}
        }
    }
}
Write-Progress -Completed -Activity "Done!"


foreach ($server in $servers)
{

    $totalCores = (Get-CimInstance Win32_Computersystem -ComputerName $server.fqdn | Measure-Object -Property NumberOfLogicalProcessors -Sum).sum
    $totalRam = (Get-CimInstance Win32_PhysicalMemory -ComputerName $server.fqdn | Measure-Object -Property capacity -Sum).Sum
           
    $cpuTime = (Get-Counter -ComputerName $server.fqdn '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue
    $availMem = (Get-Counter -ComputerName $server.fqdn '\Memory\Available MBytes').CounterSamples.CookedValue
    
    $server.Name + ': Cores: ' + ($totalCores).ToString("#,00") + ', RAM: ' + ($totalRam/1073741824).ToString("#,000") + ' GB, CPU %: ' + $cpuTime.ToString("#,00.00") + ' %, Avail.Mem.: ' + ($availMem/1024).ToString("#,000") + ' GB (' + (104857600 * $availMem / $totalRam).ToString("#,00.0") + ' %)'
}