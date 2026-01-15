
<# SCVMM preflight collection + Network connectivity/ports/WinRM/WS-Man/WMI test report
- Required ports: 5985/5986 (WinRM), 443 (HTTPS/BITS), 445 (SMB), 135 (RPC), [optional] 139 (NetBIOS)
- WS-Man agent version check: root/scvmm/AgentManagement (Invoke-WSManAction)
#>

[CmdletBinding()]
param(
    [switch]$RefreshLLDP,
    [int]$JobHistoryHours = 24,
    [switch]$IncludeLegacyNetBIOS,          # include 139/tcp test
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credential
)

# ------------------- Common utilities -------------------

# --- Version & cmdlet capability helpers (PowerShell 5.1-safe) ---

function Supports-Cmdlet {
    param([Parameter(Mandatory)][string]$Name)
    $cmd = Get-Command -Name $Name -Module VirtualMachineManager -ErrorAction SilentlyContinue
    return ($null -ne $cmd)
}

function Get-VmmVersionInfo {
    param([Parameter(Mandatory)]$VmmServer)
    $ver   = $VmmServer.ProductVersion
    $build = $VmmServer.ProductVersion
    $major = 'Unknown'; $label = 'Unknown'

    if ($ver -match '^10\.25') { $major = '2025'; $label = 'SCVMM 2025' }
    elseif ($ver -match '^10\.22') { $major = '2022'; $label = 'SCVMM 2022' }
    elseif ($ver -match '^10\.19') { $major = '2019'; $label = 'SCVMM 2019' }
    elseif ($ver -match '^(4\.0|3\.2)') { $major = '2016'; $label = 'SCVMM 2016/2012R2' }

    [pscustomobject]@{
        RawVersion = $ver
        Build      = $build
        Major      = $major
        Label      = $label
    }
}

function Get-VmmIpPools {
    param(
        $Vmm,
        [switch]$PreferStatic
    )

    # Prefer using Get-SCStaticIPAddressPool when available (SCVMM 2016+)
    if (Supports-Cmdlet 'Get-SCStaticIPAddressPool') {
        return Get-SCStaticIPAddressPool -VMMServer $Vmm -ErrorAction SilentlyContinue
    }
    # Fallback: older environments that still expose Get-SCIPPool
    elseif (Supports-Cmdlet 'Get-SCIPPool' -and (-not $PreferStatic)) {
        return Get-SCIPPool -VMMServer $Vmm -ErrorAction SilentlyContinue
    }
    else {
        Write-Warning "No IP pool cmdlet available in this environment. Skipping IP pool collection."
        return @()
    }
}

function New-OutputFolder {
    $root = 'C:\VMMReports'
    if (-not (Test-Path $root)) { New-Item -Path $root -ItemType Directory -Force | Out-Null }
    $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $out = Join-Path $root ("SCVMM_Diag_{0}" -f $stamp)
    New-Item -Path $out -ItemType Directory -Force | Out-Null
    return $out
}

function Get-Prop { param($Object, [string]$Name) try { $Object | Select-Object -ExpandProperty $Name -ErrorAction Stop } catch { $null } }
function To-Csv   { param($Data, [string]$Path) if ($null -ne $Data -and $Data.Count -gt 0) { $Data | Export-Csv -Path $Path -Encoding UTF8 -NoTypeInformation } }

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
$ErrorActionPreference = 'Stop'
$OUT  = New-OutputFolder
$HTMLpath = Join-Path $OUT 'SCVMM_Diag_Report.html'
Write-Host "[INFO] Output directory: $OUT"

Import-Module VirtualMachineManager -ErrorAction Stop
$vmm = Get-SCVMMServer -ComputerName $env:COMPUTERNAME
if (-not $vmm) { throw "Get-SCVMMServer failed. Make sure you're running in the VMM PowerShell console." }

if (-not $Credential) {
    Write-Host "[INFO] Enter credentials for remote tests (use the host management/Run As account)."
    $Credential = Get-Credential
}

# ------------------- Overview section -------------------
$serverInfo = @()
$serverRow = [ordered]@{
    VMMServer      = if ($null -ne (Get-Prop $vmm 'Name')) { (Get-Prop $vmm 'Name') } else { $env:COMPUTERNAME }
    FQDN           = (Get-Prop $vmm 'FullyQualifiedDomainName')
    Version        = if ($null -ne (Get-Prop $vmm 'Version')) { (Get-Prop $vmm 'Version') } else { (Get-Prop $vmm 'ServerVersion') }
    Build          = (Get-Prop $vmm 'BuildNumber')
    DBServer       = (Get-Prop $vmm 'DatabaseServerName')
    DBName         = (Get-Prop $vmm 'DatabaseName')
    ServiceAccount = (Get-Prop $vmm 'ServiceAccountName')
    TimeCollected  = (Get-Date)
}
# After $serverRow …
$versionInfo = Get-VmmVersionInfo -VmmServer $vmm
$serverRow['VMMVersionRaw'] = $versionInfo.RawVersion
$serverRow['VMMBuild']      = $versionInfo.Build
$serverRow['VMMVersionMajor'] = $versionInfo.Major
$serverRow['VMMVersionLabel'] = $versionInfo.Label
$serverInfo += [pscustomobject]$serverRow




$clouds     = Get-SCCloud -VMMServer $vmm -ErrorAction SilentlyContinue
$hostGroups = Get-SCVMHostGroup -VMMServer $vmm -ErrorAction SilentlyContinue
$jobsRecent = Get-SCJob -VMMServer $vmm -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -gt (Get-Date).AddHours(-$JobHistoryHours) }

# ------------------- Hosts/Agent -------------------
$hosts = Get-SCVMHost -VMMServer $vmm
$hostRows = foreach ($h in $hosts) {
    $mc = Get-Prop $h 'ManagedComputer'
    $clusterObj = Get-Prop $h 'HostCluster'
    $clusterName = if ($clusterObj) { $clusterObj.Name } else { $null }
    $groupObj    = Get-Prop $h 'VMHostGroup'
    $groupName   = if ($groupObj) { $groupObj.Name } else { $null }
    $statusVal   = if ($null -ne (Get-Prop $h 'Status')) { (Get-Prop $h 'Status') } else { (Get-Prop $h 'ServerState') }

    [pscustomobject][ordered]@{
        Host             = $h.Name
        FQDN             = (Get-Prop $h 'FullyQualifiedDomainName')
        Cluster          = $clusterName
        Group            = $groupName
        OS               = (Get-Prop $h 'OperatingSystem')
        Status           = $statusVal
        Health           = (Get-Prop $h 'HealthStatus')
        LogicalCPU       = (Get-Prop $h 'LogicalProcessorCount')
        MemoryGB         = if ($mem = Get-Prop $h 'PhysicalMemory') { [Math]::Round(($mem/1GB),1) } else { $null }
        AgentVersion     = (Get-Prop $mc 'AgentVersion')
        AgentCommState   = (Get-Prop $mc 'CommunicationState')
        AgentResponding  = (Get-Prop $mc 'IsResponding')
        AgentLastContact = (Get-Prop $mc 'LastContact')
        LastRefresh      = (Get-Prop $h 'LastRefreshTime')
    }
}

# ------------------- VM inventory -------------------
$vms = Get-SCVirtualMachine -VMMServer $vmm
$vmRows = foreach ($vm in $vms) {
    $vmHostObj  = Get-Prop $vm 'VMHost'
    $cloudObj   = Get-Prop $vm 'Cloud'
    $vmHostName = if ($vmHostObj) { $vmHostObj.Name } else { $null }
    $cloudName  = if ($cloudObj) { $cloudObj.Name } else { $null }

    $vnics = Get-SCVirtualNetworkAdapter -VM $vm -ErrorAction SilentlyContinue

    # >>> FIXED (PowerShell 5.1-safe): build pairs WITHOUT inline 'if' in the -f arguments <<<
    $nicSummary = $null
    if ($vnics -and $vnics.Count -gt 0) {
        $pairs = @()
        foreach ($vnic in $vnics) {
            $vmNetName = $null
            # Be defensive: some objects may not have VMNetwork; check existence and value
            if ($vnic -and ($vnic.PSObject.Properties.Match('VMNetwork').Count -gt 0) -and $vnic.VMNetwork) {
                $vmNetName = $vnic.VMNetwork.Name
            }
            $pairs += ("{0} -> {1}" -f $vnic.Name, $vmNetName)
        }
        $nicSummary = ($pairs -join "; ")
    }

    [pscustomobject][ordered]@{
        VM              = $vm.Name
        Status          = (Get-Prop $vm 'Status')
        Host            = $vmHostName
        Cloud           = $cloudName
        CPUCount        = (Get-Prop $vm 'CPUCount')
        MemoryMB        = (Get-Prop $vm 'MemoryMB')
        OperatingSystem = (Get-Prop $vm 'OperatingSystem')
        VMNetworks      = $nicSummary
        LastRefresh     = (Get-Prop $vm 'LastRefreshTime')
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
        IsolationType  = (Get-Prop $_ 'NetworkIsolation')
        Sites          = ($_.NetworkSites | ForEach-Object Name) -join ', '
    }
}
$vmnRows = $vmNetworks | ForEach-Object {
    $lnObj = Get-Prop $_ 'LogicalNetwork'
    [pscustomobject]@{
        VMNetwork  = $_.Name
        LogicalNet = if ($lnObj) { $lnObj.Name } else { $null }
        Subnets    = ($_.Subnets | ForEach-Object Name) -join ', '
        Isolation  = (Get-Prop $_ 'NetworkIsolation')
    }
}

$ipPoolRows = $ipPools | ForEach-Object {
    # Output only properties common to both cmdlets (Static IP Pool standard; some environments may not have dynamic pool objects)
    [pscustomobject]@{
        IPPool     = $_.Name
        LogicalNet = if ($null -ne (Get-Prop $_ 'LogicalNetwork')) { (Get-Prop $_ 'LogicalNetwork').Name } else { $null }
        StartIP    = $_.StartIPAddress
        EndIP      = $_.EndIPAddress
        Subnet     = $_.Subnet
        Gateway    = $_.DefaultGateway
        DNS        = ($_.DnsServers) -join ', '
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
        LogicalSwitch   = $_.Name
        UplinkProfile   = ($_.UplinkPortProfiles | ForEach-Object Name) -join ', '
        PortClass       = ($_.PortClassifications | ForEach-Object Name) -join ', '
        CompliantHosts  = ($_.CompliantVMHosts | ForEach-Object Name) -join ', '
    }
}

# ------------------- Host NIC–vSwitch–Logical Network mapping -------------------
$nicMapRows = New-Object System.Collections.Generic.List[object]
foreach ($h in $hosts) {
    $nics = Get-SCVMHostNetworkAdapter -VMHost $h -ErrorAction SilentlyContinue
    foreach ($nic in $nics) {
        $vsObj = Get-Prop $nic 'VirtualSwitch'
        $lsObj = Get-Prop $nic 'LogicalSwitch'
        $lns   = $nic.LogicalNetworks

        $vsName = if ($vsObj) { $vsObj.Name } else { $null }
        $lsName = if ($lsObj) { $lsObj.Name } else { $null }
        $connName = if ($null -ne $nic.ConnectionName) { $nic.ConnectionName } else { $nic.Name }
        $boundLn  = if ($lns) { ($lns | ForEach-Object Name) -join ', ' } else { $null }

        $nicMapRows.Add([pscustomobject][ordered]@{
            Host            = $h.Name
            AdapterName     = $connName
            VSwitch         = $vsName
            LogicalSwitch   = $lsName
            BoundLogicalNet = $boundLn
            VLanMode        = (Get-Prop $nic 'VlanMode')
            VLanID          = (Get-Prop $nic 'VlanID')
            IsTeamed        = (Get-Prop $nic 'IsTeamed')
            MACAddress      = (Get-Prop $nic 'MacAddress')
            LinkSpeed       = (Get-Prop $nic 'LinkSpeed')
            LastRefresh     = (Get-Prop $nic 'LastRefreshTime')
        }) | Out-Null

        if ($RefreshLLDP) {
            try { Set-SCVMHostNetworkAdapter -VMHostNetworkAdapter $nic -RefreshLLDP -ErrorAction Stop | Out-Null }
            catch { Write-Warning ("Failed to refresh LLDP: {0}" -f $_.Exception.Message) }
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
        # Query agent version from VMM AgentManagement (via WS-Man)
        try {
            $resp = Invoke-WSManAction -Action GetVersion -ComputerName $ComputerName `
                    -ResourceURI "http://schemas.microsoft.com/wbem/wsman/1/wmi/root/scvmm/AgentManagement" `
                    -Authentication Default -Credential $Credential -ErrorAction Stop
            $agentVersion = ($resp | Out-String).Trim()
        } catch { $wsmanErr = "AgentVersion error: $($_.Exception.Message)" }
    } catch { $wsmanErr = $_.Exception.Message }
    [pscustomobject]@{ WSManOk=$wsmanOk; AgentVersion=$agentVersion; WSManError=$wsmanErr }
}

function Test-WMI {
    param([string]$ComputerName,[System.Management.Automation.PSCredential]$Credential)
    # Try both CIM (WS-Man) and DCOM paths
    $wsCimOk=$false; $dcCimOk=$false; $err1=$null; $err2=$null
    try { $null = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop; $wsCimOk=$true } catch { $err1 = $_.Exception.Message }
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
$netTestRows = New-Object System.Collections.Generic.List[object]
foreach ($h in $hosts) {
    $cn = $h.Name
    $fqdn = (Get-Prop $h 'FullyQualifiedDomainName')
    $target = if ($fqdn) { $fqdn } else { $cn }

    # DNS/IP/Ping
    $dnsOk=$null; $resolvedIPs=$null; $pingMs=$null
    try {
        $r = Resolve-DnsName -Name $target -ErrorAction Stop
        $dnsOk = $true
        $resolvedIPs = ($r | Where-Object { $_.Type -eq 'A' } | ForEach-Object IPAddress) -join ', '
    } catch { $dnsOk = $false }
    try {
        $p = Test-Connection -ComputerName $target -Count 2 -ErrorAction Stop
        $avg = ($p | Measure-Object -Property ResponseTime -Average).Average
        $pingMs = if ($avg) { [Math]::Round($avg,1) } else { $null }
    } catch { $pingMs = $null }

    # TCP ports
    $portResults = foreach ($port in $BASE_PORTS) { Test-TcpPort -ComputerName $target -Port $port }
    $p443  = $portResults | Where-Object { $_.Port -eq 443 }
    $p445  = $portResults | Where-Object { $_.Port -eq 445 }
    $p5985 = $portResults | Where-Object { $_.Port -eq 5985 }
    $p5986 = $portResults | Where-Object { $_.Port -eq 5986 }
    $p135  = $portResults | Where-Object { $_.Port -eq 135 }
    $p139  = $portResults | Where-Object { $_.Port -eq 139 }

    $p139Ok = if ($p139) { $p139.TcpTestSucceeded } else { $null }

    # WS-Man & Agent
    $ws = Test-WsmanAndAgent -ComputerName $target -Credential $Credential

    # WMI (DCOM/WS-Man)
    $wmi = Test-WMI -ComputerName $target -Credential $Credential

    # Remote WinRM config snapshot (file)
    $winrmFile = Get-RemoteWinRMDetails -ComputerName $target -Credential $Credential -OutDir $OUT

    $netTestRows.Add([pscustomobject][ordered]@{
        Host              = $cn
        TargetName        = $target
        DNS_OK            = $dnsOk
        DNS_ResolvedIPs   = $resolvedIPs
        Ping_AvgMs        = $pingMs
        TCP_443_OK        = $p443.TcpTestSucceeded
        TCP_445_OK        = $p445.TcpTestSucceeded
        TCP_5985_OK       = $p5985.TcpTestSucceeded
        TCP_5986_OK       = $p5986.TcpTestSucceeded
        TCP_135_OK        = $p135.TcpTestSucceeded
        TCP_139_OK        = $p139Ok
        WSMan_OK          = $ws.WSManOk
        VMM_AgentVersion  = $ws.AgentVersion
        WSMan_Error       = $ws.WSManError
        WMI_WSMan_OK      = $wmi.WMI_WSMan
        WMI_DCOM_OK       = $wmi.WMI_DCOM
        WMI_WSMan_Error   = $wmi.WMI_WSMan_Error
        WMI_DCOM_Error    = $wmi.WMI_DCOM_Error
        WinRM_DetailsFile = $winrmFile
    }) | Out-Null
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
To-Csv $serverInfo  (Join-Path $OUT 'server_info.csv')
To-Csv $hostRows    (Join-Path $OUT 'hosts.csv')
To-Csv $vmRows      (Join-Path $OUT 'vms.csv')
To-Csv $lnRows      (Join-Path $OUT 'logical_networks.csv')
To-Csv $vmnRows     (Join-Path $OUT 'vm_networks.csv')
To-Csv $ipPoolRows  (Join-Path $OUT 'ip_pools.csv')
To-Csv $macPoolRows (Join-Path $OUT 'mac_pools.csv')
To-Csv $lswRows     (Join-Path $OUT 'logical_switches.csv')
To-Csv $nicMapRows  (Join-Path $OUT 'host_nic_switch_map.csv')
To-Csv $netTestRows (Join-Path $OUT 'host_network_tests.csv')
To-Csv $jobRows     (Join-Path $OUT 'jobs_recent.csv')

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
