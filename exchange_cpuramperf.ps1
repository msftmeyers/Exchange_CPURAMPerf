<#
.SYNOPSIS
    CPU, RAM, Performance snapshot and pagefile overview

.DESCRIPTION
    This script will show you the concurrent CPU and RAM numbers, its 
    concurrent utilization and pagefile data of every Exchange server
    in your organization.
    
    You can use this script to get a fast overview of basic recommended
    CPU Core, RAM and pagefile settings and furthermore the concurrent CPU/RAM utilization.

    The network availability of all servers will be checked first and
    only available servers data will be collected.
                                              
.PERMISSIONS
    You need to be at least a member of Exchange view-only and the local administrator
    group to run the script.

.ENVIRONMENTS
    Script is optimized to run in Exchange on premise environments, because it will only
    collect data from Exchange Servers.

    DO NOT RUN THIS SCRIPT in Powershell ISE.

.EXAMPLE
    .\exchange_cpuramperf.ps1

 .VERSIONS
    12.02.2025 V1.0 Initial version
    14.02.2025 V1.1 Minor fixes
    19.02.2025 V1.3 Pagefile, LogicalCores added, Output format changed
    27.02.2025 V1.5 PerfCounter IDs
    22.05.2025 V1.6 Exchange CPU/RAM counter are collected by invoking a scriptblock

.AUTHOR/COPYRIGHT: Steffen Meyer
.ROLE: Cloud Solution Architect
.COMPANY: Microsoft Deutschland GmbH

#>
$scriptversion = "V1.6_22.05.2025"

function Get-Counters009
{
   param (
      [Parameter(Mandatory=$false)]
      $ServerFQDN=([System.Net.Dns]::GetHostByName($env:computerName)).hostname
    )
    
    $counters009 = Invoke-Command -ComputerName $serverFQDN {(Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\009" -Name Counter).counter.tolower()}
    Return $counters009
}

function Get-CountersLocal
{
   param (
      [Parameter(Mandatory=$false)]
      $ServerFQDN=([System.Net.Dns]::GetHostByName($env:computerName)).hostname
    )
    
    $counterslocal = Invoke-Command -ComputerName $serverFQDN {(Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\CurrentLanguage" -Name Counter).Counter }
    Return $counterslocal
}

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
if ((Get-Host).name -ne 'Windows PowerShell ISE Host')
{
    if (!(Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue))
    {
        write-host "`nATTENTION: Exchange Admin machine detected, this script can only be executed on an Exchangeserver!`n" -ForegroundColor Cyan
        Return
    }
}
else
{
    Write-Host "`nATTENTION: Do not run this script in Powershell ISE, please use Windows- or Exchange Powershell!"
    Return
}

Write-Host "`nCurrent time stamp: $(Get-Date $now -Format "dd.MM.yyyy HH:mm")"
Write-Host "Script version:     $scriptversion"

Write-Host "`n----------------------------------------------------------------------------"
Write-Host "|           This script will check the CPU Core/RAM numbers,               |"
Write-Host "|         concurrent CPU/RAM utilization and pagefile size and             |"
Write-Host "|                 settings of all Exchange servers                         |" 
Write-Host "|                  in this Exchange organization.                          |"
Write-Host "----------------------------------------------------------------------------"

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

$result = @()
$i = 0

foreach ($server in $servers)
{
    #Progress bar
    $i++
    $status = "{0:N0}" -f ($i / $servers.count * 100)
    Write-Progress -Activity "Analyzing server $server..." -Status "Processing server $i of $($servers.count) : $status% Completed" -PercentComplete ($i / $servers.count * 100)
        
    #Collect CPU / RAM information
    $CPUProc = Get-CimInstance CIM_Processor -ComputerName $server.fqdn

    $totalphysCores = ($CPUProc | Measure-Object -Property NumberOfEnabledCore -Sum).sum
    $totallogiCores = ($CPUProc | Measure-Object -Property NumberOfLogicalProcessors -Sum).sum

    $totalRam = (Get-CimInstance Win32_PhysicalMemory -ComputerName $server.fqdn | Measure-Object -Property capacity -Sum).Sum/1048576
    
    #Collect english counters
    $counters009 = Get-Counters009 -ServerFQDN $server.fqdn

    #Filter counter indexes
    $procobjectindex = $counters009.IndexOf("Processor".tolower())
    $proccounterindex = $counters009.IndexOf("% Processor Time".tolower())
    
    $memobjectindex = $counters009.IndexOf("Memory".tolower())
    $memcounterindex = $counters009.IndexOf("Available MBytes".tolower())

    #Collect language specific counters
    $counterslocal = Get-CountersLocal -ServerFQDN $server.fqdn
    
    #Match localized counter - Proc
    $object = $counterslocal[$procobjectindex]
    $counter = $counterslocal[$proccounterindex]
    $cpuTime = (Get-Counter -ComputerName $server.fqdn "\$object(_total)\$counter").CounterSamples.CookedValue

    #Match localized counter - Memory
    $object = $counterslocal[$memobjectindex]
    $counter = $counterslocal[$memcounterindex]
    $availMem = (Get-Counter -ComputerName $server.fqdn "\$object\$counter").CounterSamples.CookedValue

    #Collect Pagefile information
    $auto = (Get-CimInstance -ComputerName $server.fqdn Win32_ComputerSystem).AutomaticManagedPagefile
    $initial = (Get-WmiObject -ComputerName $server.fqdn WIN32_Pagefile).initialsize
    $maximum = (Get-WmiObject -ComputerName $server.fqdn WIN32_Pagefile).maximumsize
    
    #Check Pagefile for Ex2019
    if ($server.AdminDisplayVersion -like "*15.2*")
    {
        if (($totalRam -le "262144") -and ($totalRam -ge "131072"))
        {
            $RAMDiff = "OK"
        }
        else
        {
            $RAMDiff = "WRONG"
        }
        
        if (($totallogiCores -le "48") -and ($totallogiCores -eq $totalphysCores))
        {
            $coreDiff = "OK"
        }
        else
        {
            $coreDiff = "WRONG"
        }
        
        if (($initial -eq ($totalRam/4)) -and ($maximum -eq ($totalRam/4)))
        {
            $pagefile = "OK"
        }
        else
        {
            $pagefile = "WRONG"
        }
    }

    #Check Pagefile for Ex2013/2016
    else
    {
        if ($totalRam -le "196608")
        {
            $RAMDiff = "OK"
        }
        else
        {
            $RAMDiff = "WRONG"
        }
        
        if (($totallogiCores -le "24") -and ($totallogiCores -eq $totalphysCores))
        {
            $coreDiff = "OK"
        }
        else
        {
            $coreDiff = "WRONG"
        }
        
        if ($totalRam -ge "32768")
        {
            if (($initial -eq "32778") -and ($maximum -eq "32778"))
            {
                $pagefile = "OK"
            }
            else
            {
                $pagefile = "WRONG"
            }
        }
        else
        {
            if (($initial -eq ($totalRam + 10)) -and ($maximum -eq ($totalRam + 10)))
            {
                $pagefile = "OK"
            }
            else
            {
                $pagefile = "WRONG"
            }
        }   
    
    }
    
    $result += New-Object -Type PSObject -Prop @{Servername=$server.name;PhysCores=$totalphysCores;LogiCores=$totallogiCores;CoreDiff=$coreDiff;'RAM GB'=$totalRam/1024;RAMDiff=$RAMDiff;'CPU Util %'=$cpuTime;'Avail.Mem GB'=$availMem/1024;'Avail.Mem %'=100*$availMem/$totalRam;SysManaged=$auto;InitSize=$initial;MaxSize=$maximum;Pagefile=$pagefile}
}
Write-Progress -Completed -Activity "Done!"

#Output incl. ESCAPE chars for different colours
$result | format-table Servername,PhysCores,@{n="LogiCores";e={if($_.CoreDiff -eq 'OK'){"$([char]27)[32m$($_.LogiCores)$([char]27)[0m"}else{"$([char]27)[31m$($_.LogiCores)$([char]27)[0m"}};a='right'},@{n='RAM GB';e={if($_.RAMDiff -eq 'OK'){"$([char]27)[32m$("{0:N0}" -f $_.'RAM GB')$([char]27)[0m"}else{"$([char]27)[31m$("{0:N0}" -f $_.'RAM GB')$([char]27)[0m"}};a='right'},@{n='CPU Util %';e={if($_.'CPU Util %' -le 40){"$([char]27)[32m$("{0:N2}" -f $_.'CPU Util %')$([char]27)[0m"}else{"$([char]27)[31m$("{0:N2}" -f $_.'CPU Util %')$([char]27)[0m"}};a="right"},@{n='Avail.Mem GB';e={"{0:N0}" -f $_.'Avail.Mem GB'};a='right'},@{n='Avail.Mem %';e={if($_.'Avail.Mem %' -gt 25){"$([char]27)[32m$("{0:N1}" -f $_.'Avail.Mem %')$([char]27)[0m"}else{"$([char]27)[31m$("{0:N1}" -f $_.'Avail.Mem %')$([char]27)[0m"}};a="right"},SysManaged,InitSize,MaxSize,@{n="Pagefile";e={if($_.pagefile -eq 'OK'){"$([char]27)[32m$($_.pagefile)$([char]27)[0m"}else{"$([char]27)[31m$($_.pagefile)$([char]27)[0m"}};a='right'}
#END