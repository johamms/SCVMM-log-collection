#Requires -Version 5.1
#Requires -Modules VirtualMachineManager
#Requires -RunAsAdministrator

<# 
.SYNOPSIS
This script will generate a html-based test report on SCVMM preflight collection and the etwork connectivity/ports/WinRM/WS-Man/WMI

.DESCRIPTION
This script collects various SCVMM server, host, VM, network, and job information, and performs network connectivity tests to the managed hosts, including TCP port tests, WinRM configuration retrieval, WS-Man connectivity, VMM agent version check, and WMI connectivity tests.
The output is saved as an HTML report along with CSV files for each data section.
- Required ports: 5985/5986 (WinRM), 443 (HTTPS/BITS), 445 (SMB), 135 (RPC), [optional] 139 (NetBIOS)
- WS-Man agent version check: root/scvmm/AgentManagement (Invoke-WSManAction)

.PARAMETER RefreshLLDP
If specified, the script will attempt to refresh LLDP information on each host network adapter.
.PARAMETER JobHistoryHours
Specifies the time window (in hours) for collecting recent job history. Default is 24 hours.
.PARAMETER IncludeLegacyNetBIOS
If specified, the script will include tests for legacy NetBIOS port 139/tcp.    
.PARAMETER Credential
Specifies the credentials to use for remote host tests. If not provided, the script will prompt for credentials.

.OUTPUTS
HTML report file, CSV files and WinRM configugration files per host in a timestamped output folder under C:\VMMReports\SCVMM_Diag_.

.LINK
https://github.com/johamms/SCVMM-log-collection

#>

[CmdletBinding()]
param(
    [switch]$RefreshLLDP,
    [int]$JobHistoryHours = 24,
    [switch]$IncludeLegacyNetBIOS,          # include 139/tcp test
    [System.Management.Automation.PSCredential]$Credential
)

# --- Version & cmdlet capability helpers (PowerShell 5.1-safe) ---

function Get-VmmVersionInfo {
    param([Parameter(Mandatory)]$VmmServer)
    $ver   = $VmmServer.ProductVersion
    $build = $VmmServer.ProductVersion
    $major = 'Unknown'; $label = 'Unknown'

    switch -Regex ($ver) {
        '^10\.25' { $major = '2025'; $label = 'SCVMM 2025' }
        '^10\.22' { $major = '2022'; $label = 'SCVMM 2022' }
        '^10\.19' { $major = '2019'; $label = 'SCVMM 2019' }
        '^(4\.0|3\.2)' { $major = '2016'; $label = 'SCVMM 2016/2012R2' }
    }

    [pscustomobject]@{
        RawVersion = $ver
        Build      = $build
        Major      = $major
        Label      = $label
    }
}

function Get-VmmIpPools {
    param($Vmm, [switch]$PreferStatic)

    # Prefer Get-SCStaticIPAddressPool (SCVMM 2016+), fallback to Get-SCIPPool
    if (Get-Command 'Get-SCStaticIPAddressPool' -Module VirtualMachineManager -ErrorAction SilentlyContinue) {
        return Get-SCStaticIPAddressPool -VMMServer $Vmm -ErrorAction SilentlyContinue
    }
    elseif (-not $PreferStatic -and (Get-Command 'Get-SCIPPool' -Module VirtualMachineManager -ErrorAction SilentlyContinue)) {
        return Get-SCIPPool -VMMServer $Vmm -ErrorAction SilentlyContinue
    }
    else {
        Write-Warning "No IP pool cmdlet available. Skipping IP pool collection."
        return @()
    }
}

function New-OutputFolder {
    $root = 'C:\VMMReports'
    $null = New-Item -Path $root -ItemType Directory -Force -ErrorAction SilentlyContinue
    $out = Join-Path $root ("SCVMM_Diag_{0:yyyyMMdd_HHmmss}" -f (Get-Date))
    $null = New-Item -Path $out -ItemType Directory -Force
    return $out
}

function Export-ToCsv {
    param([Parameter(ValueFromPipeline)]$Data, [string]$Path)
    begin { $items = @() }
    process { if ($Data) { $items += $Data } }
    end { if ($items.Count) { $items | Export-Csv -Path $Path -Encoding UTF8 -NoTypeInformation } }
}

function Html-Style {
@"
<style>
body { font-family: Segoe UI, Arial, sans-serif; color:#222; }
h1 { border-bottom: 3px solid #005ea5; padding-bottom:4px; }
h2 { margin-top:28px; color:#005ea5; }
.tablebox { margin: 10px 0 20px 0; }
table { border-collapse: collapse; width: 100%; table-layout: fixed; }
th, td { border: 1px solid #ddd; padding: 6px 8px; word-wrap: break-word; }
th { background: #f3f7fb; text-align: left; }
caption { text-align:left; font-weight:bold; margin:6px 0; }
footer { margin-top:20px; font-size:12px; color:#666; }
.small { font-size: 12px; color:#666; }
.code { font-family: Consolas, monospace; background:#fafafa; border:1px solid #eee; padding:6px; }
</style>
"@
}

function To-HtmlSection { param([string]$Title, $Data)
    if ($null -eq $Data -or $Data.Count -eq 0) { return "<h2>$Title</h2><div class='small'>No data</div>" }
    return ($Data | ConvertTo-Html -As Table -PreContent "<h2>$Title</h2><div class='tablebox'>" -PostContent "</div>" -Fragment)
}

# ------------------- Output/Module preparation -------------------
$OUT  = New-OutputFolder
$HTMLpath = Join-Path $OUT 'SCVMM_Diag_Report.html'
Write-Host "[INFO] Output directory: $OUT"

$vmm = Get-SCVMMServer -ComputerName $env:COMPUTERNAME
if (-not $vmm) { throw "Get-SCVMMServer failed. Make sure you're running in the VMM PowerShell console." }

if (-not $Credential) {
    Write-Host "[INFO] Enter credentials for remote tests (use the host management/Run As account)."
    $Credential = Get-Credential
}

# ------------------- Overview section -------------------
$versionInfo = Get-VmmVersionInfo -VmmServer $vmm
$serverInfo = @([pscustomobject][ordered]@{
    VMMServer       = if ($vmm.Name) { $vmm.Name } else { $env:COMPUTERNAME }
    FQDN            = $vmm.FullyQualifiedDomainName
    Version         = if ($vmm.Version) { $vmm.Version } else { $vmm.ServerVersion }
    Build           = $vmm.BuildNumber
    DBServer        = $vmm.DatabaseServerName
    DBName          = $vmm.DatabaseName
    ServiceAccount  = $vmm.ServiceAccountName
    TimeCollected   = Get-Date
    VMMVersionRaw   = $versionInfo.RawVersion
    VMMBuild        = $versionInfo.Build
    VMMVersionMajor = $versionInfo.Major
    VMMVersionLabel = $versionInfo.Label
})




$clouds     = Get-SCCloud -VMMServer $vmm -ErrorAction SilentlyContinue
$hostGroups = Get-SCVMHostGroup -VMMServer $vmm -ErrorAction SilentlyContinue
$jobsRecent = Get-SCJob -VMMServer $vmm -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -gt (Get-Date).AddHours(-$JobHistoryHours) }

# ------------------- Hosts/Agent -------------------
$hosts = Get-SCVMHost -VMMServer $vmm
$hostRows = foreach ($h in $hosts) {
    $mc = $h.ManagedComputer
    [pscustomobject][ordered]@{
        Host             = $h.Name
        FQDN             = $h.FullyQualifiedDomainName
        Cluster          = $h.HostCluster.Name
        Group            = $h.VMHostGroup.Name
        OS               = $h.OperatingSystem
        Status           = if ($h.Status) { $h.Status } else { $h.ServerState }
        Health           = $h.HealthStatus
        LogicalCPU       = $h.LogicalProcessorCount
        MemoryGB         = if ($h.PhysicalMemory) { [Math]::Round($h.PhysicalMemory / 1GB, 1) } else { $null }
        AgentVersion     = $mc.AgentVersion
        AgentCommState   = $mc.CommunicationState
        AgentResponding  = $mc.IsResponding
        AgentLastContact = $mc.LastContact
        LastRefresh      = $h.LastRefreshTime
    }
}

# ------------------- VM inventory -------------------
$vms = Get-SCVirtualMachine -VMMServer $vmm
$vmRows = foreach ($vm in $vms) {
    $vnics = Get-SCVirtualNetworkAdapter -VM $vm -ErrorAction SilentlyContinue
    $nicSummary = ($vnics | ForEach-Object { "{0} -> {1}" -f $_.Name, $_.VMNetwork.Name }) -join '; '

    [pscustomobject][ordered]@{
        VM              = $vm.Name
        Status          = $vm.Status
        Host            = $vm.VMHost.Name
        Cloud           = $vm.Cloud.Name
        CPUCount        = $vm.CPUCount
        MemoryMB        = $vm.MemoryMB
        OperatingSystem = $vm.OperatingSystem
        VMNetworks      = $nicSummary
        LastRefresh     = $vm.LastRefreshTime
    }
}

# ------------------- Network (Logical/VM/Pools/Switches) -------------------
$logicalNetworks = Get-SCLogicalNetwork -VMMServer $vmm -ErrorAction SilentlyContinue
$vmNetworks      = Get-SCVMNetwork      -VMMServer $vmm -ErrorAction SilentlyContinue
#$ipPools         = Get-SCIPPool         -VMMServer $vmm -ErrorAction SilentlyContinue
$ipPools = Get-VmmIpPools -Vmm $vmm -PreferStatic
$macPools        = Get-SCMACAddressPool -VMMServer $vmm -ErrorAction SilentlyContinue
$logicalSwitches = Get-SCLogicalSwitch  -VMMServer $vmm -ErrorAction SilentlyContinue

$lnRows = $logicalNetworks | ForEach-Object {
    [pscustomobject]@{
        LogicalNetwork = $_.Name
        IsolationType  = $_.NetworkIsolation
        Sites          = ($_.NetworkSites.Name) -join ', '
    }
}

$vmnRows = $vmNetworks | ForEach-Object {
    [pscustomobject]@{
        VMNetwork  = $_.Name
        LogicalNet = $_.LogicalNetwork.Name
        Subnets    = ($_.Subnets.Name) -join ', '
        Isolation  = $_.NetworkIsolation
    }
}

$ipPoolRows = $ipPools | ForEach-Object {
    [pscustomobject]@{
        IPPool     = $_.Name
        LogicalNet = $_.LogicalNetwork.Name
        StartIP    = $_.StartIPAddress
        EndIP      = $_.EndIPAddress
        Subnet     = $_.Subnet
        Gateway    = $_.DefaultGateway
        DNS        = $_.DnsServers -join ', '
        InUse      = $_.AddressesInUse
        Available  = $_.AddressesAvailable
    }
}


$macPoolRows = $macPools | ForEach-Object {
    [pscustomobject]@{
        MACPool   = $_.Name
        StartMAC  = $_.StartMacAddress
        EndMAC    = $_.EndMacAddress
        Allocated = $_.AllocatedMacAddresses
        Available = $_.AvailableMacAddresses
    }
}

$lswRows = $logicalSwitches | ForEach-Object {
    [pscustomobject]@{
        LogicalSwitch  = $_.Name
        UplinkProfile  = $_.UplinkPortProfiles.Name -join ', '
        PortClass      = $_.PortClassifications.Name -join ', '
        CompliantHosts = $_.CompliantVMHosts.Name -join ', '
    }
}

# ------------------- Host NIC–vSwitch–Logical Network mapping -------------------
$nicMapRows = foreach ($h in $hosts) {
    $nics = Get-SCVMHostNetworkAdapter -VMHost $h -ErrorAction SilentlyContinue
    foreach ($nic in $nics) {
        [pscustomobject][ordered]@{
            Host            = $h.Name
            AdapterName     = if ($nic.ConnectionName) { $nic.ConnectionName } else { $nic.Name }
            VSwitch         = $nic.VirtualSwitch.Name
            LogicalSwitch   = $nic.LogicalSwitch.Name
            BoundLogicalNet = $nic.LogicalNetworks.Name -join ', '
            VLanMode        = $nic.VlanMode
            VLanID          = $nic.VlanID
            IsTeamed        = $nic.IsTeamed
            MACAddress      = $nic.MacAddress
            LinkSpeed       = $nic.LinkSpeed
            LastRefresh     = $nic.LastRefreshTime
        }

        if ($RefreshLLDP) {
            try { $null = Set-SCVMHostNetworkAdapter -VMHostNetworkAdapter $nic -RefreshLLDP -ErrorAction Stop }
            catch { Write-Warning "Failed to refresh LLDP: $($_.Exception.Message)" }
        }
    }
}

# ------------------- Network connectivity/Ports/WinRM/WMI tests -------------------
# Reference ports: 5985/5986 (WinRM), 443 (BITS/HTTPS), 445 (SMB), 135 (RPC), [optional] 139 (NetBIOS)
$BASE_PORTS = @(443,445,5985,5986,135)
if ($IncludeLegacyNetBIOS) { $BASE_PORTS += 139 }

function Test-TcpPort {
    param([string]$ComputerName,[int]$Port)
    try {
        $res = Test-NetConnection -ComputerName $ComputerName -Port $Port -InformationLevel Detailed -WarningAction SilentlyContinue
        [pscustomobject]@{
            Port             = $Port
            TcpTestSucceeded = $res.TcpTestSucceeded
            RemoteAddress    = $res.RemoteAddress
            RoundTripTimeMs  = if ($res.PingReplyDetails) { $res.PingReplyDetails.RoundtripTime } else { $null }
        }
    } catch {
        [pscustomobject]@{ Port=$Port; TcpTestSucceeded=$false; RemoteAddress=$null; RoundTripTimeMs=$null }
    }
}

function Get-RemoteWinRMDetails {
    param([string]$ComputerName,[System.Management.Automation.PSCredential]$Credential,[string]$OutDir)
    $file = Join-Path $OutDir ("WinRM_"+$ComputerName+".txt")
    $sb = {
        $out = @()
        $out += "=== winrm get winrm/config ==="
        try { $out += (winrm get winrm/config) } catch { $out += ("ERROR: winrm/config - {0}" -f $_.Exception.Message) }
        $out += "`n=== winrm enum winrm/config/listener ==="
        try { $out += (winrm enum winrm/config/listener) } catch { $out += ("ERROR: listeners - {0}" -f $_.Exception.Message) }
        $out += "`n=== netsh http show iplisten ==="
        try { $out += (netsh http show iplisten) } catch { $out += ("ERROR: iplisten - {0}" -f $_.Exception.Message) }
        $out -join "`n"
    }
    try {
        $txt = Invoke-Command -ComputerName $ComputerName -Credential $Credential -ScriptBlock $sb -ErrorAction Stop
        #Set-Content -Path $file -Value $txt -Encoding UTF8
        $txt | Set-Content -Path $file -Encoding UTF8
        return $file
    } catch {
        #Set-Content -Path $file -Value ("REMOTE EXECUTION FAILED: "+$_.Exception.Message) -Encoding UTF8
        ("REMOTE EXECUTION FAILED: " + $_.Exception.Message) |
          Out-File -FilePath $file -Encoding UTF8

        return $file
    }
}

function Test-WsmanAndAgent {
    param([string]$ComputerName,[System.Management.Automation.PSCredential]$Credential)
    $wsmanOk = $false; $agentVersion=$null; $wsmanErr=$null
    try {
        $null = Test-WSMan -ComputerName $ComputerName -Authentication Default -ErrorAction Stop
        $wsmanOk = $true
    } catch { $wsmanErr = $_.Exception.Message }
    # Query agent version from VMM AgentManagement (via WS-Man)
    try {
        $resp = Invoke-WSManAction -Action GetVersion -ComputerName $ComputerName `
                -ResourceURI "http://schemas.microsoft.com/wbem/wsman/1/wmi/root/scvmm/AgentManagement" `
                -Authentication Default -Credential $Credential -ErrorAction Stop
        $agentVersion = $resp.Version
    } catch { $wsmanErr = "AgentVersion error: $($_.Exception.Message)" }
    [pscustomobject]@{ WSManOk=$wsmanOk; AgentVersion=$agentVersion; WSManError=$wsmanErr }
}

function Test-WMI {
    param([string]$ComputerName,[System.Management.Automation.PSCredential]$Credential)
    # Try both CIM (WS-Man) and DCOM paths
    $wsCimOk=$false; $dcCimOk=$false; $err1=$null; $err2=$null
    try { 
        $null = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
        $wsCimOk=$true 
    } catch { $err1 = $_.Exception.Message }
    try { 
        $opt = New-CimSessionOption -Protocol Dcom
        $sess = New-CimSession -ComputerName $ComputerName -Credential $Credential -SessionOption $opt -ErrorAction Stop
        $null = Get-CimInstance -CimSession $sess -ClassName Win32_OperatingSystem -ErrorAction Stop
        $dcCimOk=$true
        $sess | Remove-CimSession
    } catch { $err2 = $_.Exception.Message }
    [pscustomobject]@{ WMI_WSMan=$wsCimOk; WMI_DCOM=$dcCimOk; WMI_WSMan_Error=$err1; WMI_DCOM_Error=$err2 }
}

# Host-wise network tests
$netTestRows = foreach ($h in $hosts) {
    $target = if ($h.FullyQualifiedDomainName) { $h.FullyQualifiedDomainName } else { $h.Name }

    # DNS/IP/Ping
    $dnsOk = $false; $resolvedIPs = $null; $pingMs = $null
    try {
        $r = Resolve-DnsName -Name $target -ErrorAction Stop
        $dnsOk = $true
        $resolvedIPs = ($r | Where-Object Type -eq 'A').IPAddress -join ', '
    } catch { }
    try {
        $pingMs = [Math]::Round((Test-Connection -ComputerName $target -Count 2 -ErrorAction Stop |
            Measure-Object -Property ResponseTime -Average).Average, 1)
    } catch { }

    # TCP ports - build hashtable for easy lookup
    $portResults = @{}
    foreach ($port in $BASE_PORTS) {
        $portResults[$port] = (Test-TcpPort -ComputerName $target -Port $port).TcpTestSucceeded
    }

    # WS-Man, WMI & WinRM details
    $ws = Test-WsmanAndAgent -ComputerName $target -Credential $Credential
    $wmi = Test-WMI -ComputerName $target -Credential $Credential
    $winrmFile = Get-RemoteWinRMDetails -ComputerName $target -Credential $Credential -OutDir $OUT

    [pscustomobject][ordered]@{
        Host              = $h.Name
        TargetName        = $target
        DNS_OK            = $dnsOk
        DNS_ResolvedIPs   = $resolvedIPs
        Ping_AvgMs        = $pingMs
        TCP_443_OK        = $portResults[443]
        TCP_445_OK        = $portResults[445]
        TCP_5985_OK       = $portResults[5985]
        TCP_5986_OK       = $portResults[5986]
        TCP_135_OK        = $portResults[135]
        TCP_139_OK        = $portResults[139]
        WSMan_OK          = $ws.WSManOk
        VMM_AgentVersion  = $ws.AgentVersion
        WSMan_Error       = $ws.WSManError
        WMI_WSMan_OK      = $wmi.WMI_WSMan
        WMI_DCOM_OK       = $wmi.WMI_DCOM
        WMI_WSMan_Error   = $wmi.WMI_WSMan_Error
        WMI_DCOM_Error    = $wmi.WMI_DCOM_Error
        WinRM_DetailsFile = $winrmFile
    }
}

# ------------------- Recent jobs -------------------
$jobRows = foreach ($j in $jobsRecent) {
    [pscustomobject]@{
        JobId       = $j.ID
        Name        = $j.Name
        Status      = $j.Status
        Owner       = $j.Owner
        StartTime   = $j.StartTime
        EndTime     = $j.EndTime
        Result      = $j.Result
        Error       = if ($j.ErrorCode) { ("{0} - {1}" -f $j.ErrorCode, $j.ErrorMessage) } else { $null }
    }
}

# ------------------- Save CSVs -------------------
$csvExports = @{
    'server_info.csv'         = $serverInfo
    'hosts.csv'               = $hostRows
    'vms.csv'                 = $vmRows
    'logical_networks.csv'    = $lnRows
    'vm_networks.csv'         = $vmnRows
    'ip_pools.csv'            = $ipPoolRows
    'mac_pools.csv'           = $macPoolRows
    'logical_switches.csv'    = $lswRows
    'host_nic_switch_map.csv' = $nicMapRows
    'host_network_tests.csv'  = $netTestRows
    'jobs_recent.csv'         = $jobRows
}
$csvExports.GetEnumerator() | ForEach-Object { $_.Value | Export-ToCsv -Path (Join-Path $OUT $_.Key) }

# ------------------- HTML report -------------------
$sections = @()
$sections += "<h2>SCVMM Diagnostic report</h2><div class='small'>Creation Time: $(Get-Date) / Server: $($serverInfo[0].VMMServer)</div>"
#$sections += "<div class='small'>SCVMM Version: $($versionInfo.Label) (raw: $($versionInfo.RawVersion), build: $($versionInfo.Build))</div>"
$sections += To-HtmlSection -Title 'VMM Server information'                        -Data $serverInfo
$sections += To-HtmlSection -Title 'Cloud & Host Group'                            -Data (@([pscustomobject]@{ Clouds = ($clouds | ForEach-Object Name) -join ', ' ; HostGroups = ($hostGroups | ForEach-Object Name) -join ', ' }))
$sections += To-HtmlSection -Title 'Host & VMM Agent status'                       -Data $hostRows
$sections += To-HtmlSection -Title 'VM Inventory & vNIC-VMNetwork'                 -Data $vmRows
$sections += To-HtmlSection -Title 'Logical Network'                               -Data $lnRows
$sections += To-HtmlSection -Title 'VM Network'                                    -Data $vmnRows
$sections += To-HtmlSection -Title 'IP Pool'                                       -Data $ipPoolRows
$sections += To-HtmlSection -Title 'MAC Pool'                                      -Data $macPoolRows
$sections += To-HtmlSection -Title 'Logical Switch'                                -Data $lswRows
$sections += To-HtmlSection -Title 'Host NIC-Virtual Switch-Logical Network mapping' -Data $nicMapRows
$sections += To-HtmlSection -Title ('Recent jobs (past {0} hour)' -f $JobHistoryHours) -Data $jobRows
$sections += To-HtmlSection -Title 'Host Network Connection Test (Port/WS-Man/WMI)'    -Data $netTestRows

$html = ConvertTo-Html -Head (Html-Style) -Body ($sections -join "`n")
$null = $html | Set-Content -Path $HTMLpath -Encoding UTF8

Write-Host "[INFO] Report generated: $HTMLpath"
Write-Host "[INFO] Source CSV folder: $OUT"
Write-Host "[TIP ] Use -RefreshLLDP to refresh LLDP details if you need latest switch/port info."
