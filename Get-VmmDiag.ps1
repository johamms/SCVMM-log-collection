#Requires -Version 5.1
#Requires -Modules VirtualMachineManager
#Requires -RunAsAdministrator

<# 
.SYNOPSIS
This script will generate a html-based test report on SCVMM preflight collection and the network connectivity/ports/WinRM/WS-Man/WMI

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

.EXAMPLE
PS> .\Get-VmmDiag.ps1

.EXAMPLE
PS> .\Get-VmmDiag.ps1 -RefreshLLDP -JobHistoryHours 24 -Credential (Get-Credential)

.EXAMPLE
PS> .\Get-VmmDiag.ps1 -RefreshLLDP -JobHistoryHours 24 -IncludeLegacyNetBIOS -Credential "Contoso\administrator"

.LINK
https://github.com/johamms/SCVMM-log-collection

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
    $major = 'Unknown'; $label = 'Unknown'

    if ($ver -match '^10\.25') { $major = '2025'; $label = 'SCVMM 2025' }
    elseif ($ver -match '^10\.22') { $major = '2022'; $label = 'SCVMM 2022' }
    elseif ($ver -match '^10\.19') { $major = '2019'; $label = 'SCVMM 2019' }
    elseif ($ver -match '^(4\.0|3\.2)') { $major = '2016'; $label = 'SCVMM 2016/2012R2' }

    [pscustomobject]@{
        RawVersion = $ver
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
        Write-Warning "[WARN] No IP pool cmdlet available in this environment. Skipping IP pool collection."
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
/* Severity coloring for VMM Error Log */
.sev-Critical { background: #ffe6e6; border-left: 4px solid #b30000; }
.sev-High     { background: #fff0e6; border-left: 4px solid #b35900; }
.sev-Medium   { background: #fff9e6; border-left: 4px solid #b38f00; }
.sev-Low      { color: #666; }
.badge        { display:inline-block; padding:2px 6px; font-size:12px; border-radius:3px; background:#eee; margin-right:6px; }
.badge-critical { background:#b30000; color:#fff; }
.badge-high     { background:#b35900; color:#fff; }
.badge-medium   { background:#b38f00; color:#fff; }
.badge-low      { background:#999; color:#fff; }
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

$vmm = Get-SCVMMServer -ComputerName $env:COMPUTERNAME
if (-not $vmm) { throw "Get-SCVMMServer failed. Make sure you're running in the VMM PowerShell console." }

if (-not $Credential) {
    $msgString = "Enter credentials for remote tests (use the host management/Run As account)."
    Write-Host "[INFO] $msgString"
    $Credential = Get-Credential -UserName "$env:USERDOMAIN\$env:USERNAME" -Message $msgString
    Write-Host "[INFO] Using credentials for user: $($Credential.UserName)"
}

# ------------------- Overview section -------------------
$serverInfo = @()
$serverRow = [ordered]@{
    VMMServer      = if ($null -ne (Get-Prop $vmm 'Name')) { (Get-Prop $vmm 'Name') } else { $env:COMPUTERNAME }
    FQDN           = (Get-Prop $vmm 'FullyQualifiedDomainName')
    # Version        = if ($null -ne (Get-Prop $vmm 'Version')) { (Get-Prop $vmm 'Version') } else { (Get-Prop $vmm 'ServerVersion') }
    # Build          = (Get-Prop $vmm 'BuildNumber')
    DBServer       = (Get-Prop $vmm 'DatabaseServerName')
    DBInstance     = (Get-Prop $vmm 'DatabaseInstanceName')
    DBName         = (Get-Prop $vmm 'DatabaseName')
    ServiceAccount = (Get-Prop $vmm 'VMMServiceAccount')
    TimeCollected  = (Get-Date)
}
# After $serverRow …
$versionInfo = Get-VmmVersionInfo -VmmServer $vmm
$serverRow['VMMProductVersion'] = $versionInfo.RawVersion
$serverRow['VMMVersionLabel'] = $versionInfo.Label
$serverRow['VMMVersionMajor'] = $versionInfo.Major
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
    $statusVal   = if ($null -ne (Get-Prop $h 'OverallState')) { (Get-Prop $h 'OverallState') } else { (Get-Prop $h 'ComputerState') }

    [pscustomobject][ordered]@{
        Host             = $h.Name
        FQDN             = (Get-Prop $h 'FullyQualifiedDomainName')
        Cluster          = $clusterName
        Group            = $groupName
        OS               = (Get-Prop $h 'OperatingSystem')
        Type             = (Get-Prop $h 'ObjectType')
        Status           = $statusVal
        #Health           = (Get-Prop $h 'HealthStatus')
        LogicalCPU       = (Get-Prop $h 'LogicalProcessorCount')
        TotalMemoryGB         = if ($memByte = Get-Prop $h 'TotalMemory') { [Math]::Round(($memByte/1GB),1) } else { $null }
        AvailableMemoryGB         = if ($memByte = Get-Prop $h 'AvailableMemory') { [Math]::Round(($memByte/1GB),2) } else { $null }
        AgentVersion     = (Get-Prop $mc 'AgentVersion')
        AgentCommState   = (Get-Prop $mc 'State')
        #AgentResponding  = (Get-Prop $mc 'IsResponding')
        AgentLastContact = (Get-Prop $mc 'UpdatedDate')
        #LastRefresh      = (Get-Prop $h 'LastRefreshTime')
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
        MemoryMB        = if ($memMB = Get-Prop $vm 'Memory') { [Math]::Round($memMB) } else { $null }
        OperatingSystem = (Get-Prop $vm 'OperatingSystem')
        VMNetworks      = $nicSummary
        LastRefresh     = (Get-Prop $vm 'ModifiedTime')
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
    $sites = (Get-SCLogicalNetworkDefinition -LogicalNetwork $_ | ForEach-Object Name) -join ', '
    $DefaultNetworkManagerForVMNetworkCreation = Get-Prop $_ 'DefaultNetworkManagerForVMNetworkCreation'
    [pscustomobject]@{
        LogicalNetwork = $_.Name
        SupportsExternalVMNetworkProvisioning = $_.SupportsExternalVMNetworkProvisioning
        DefaultNetworkManagerForVMNetworkCreation = if ($null -ne $DefaultNetworkManagerForVMNetworkCreation) { $DefaultNetworkManagerForVMNetworkCreation } else { 'Hyper-V Network Virtualization' }
        IsManagedByNetworkController = $_.IsManagedByNetworkController
        IsPublicIPNetwork = $_.IsPublicIPNetwork
        Sites          = $sites
    }
}
$vmnRows = $vmNetworks | ForEach-Object {
    $lnObj = Get-Prop $_ 'LogicalNetwork'
    [pscustomobject]@{
        VMNetwork  = $_.Name
        LogicalNetwork = if ($lnObj) { $lnObj.Name } else { $null }
        Subnets    = ($_.VMSubnet | ForEach-Object Name) -join ', '
        IsolationType  = (Get-Prop $_ 'IsolationType')
        GatewayConnection = (Get-Prop $_ 'HasGatewayConnection')
    }
}

$ipPoolRows = $ipPools | ForEach-Object {
    $site = $_.LogicalNetworkDefinition
    # Output only properties common to both cmdlets (Static IP Pool standard; some environments may not have dynamic pool objects)
    [pscustomobject]@{
        IPPool     = $_.Name
        LogicalNetwork = Get-Prop $site 'LogicalNetwork'
        NetworkSite = (Get-Prop $site 'Name')
        StartIP    = $_.IPAddressRangeStart
        EndIP      = $_.IPAddressRangeEnd
        Subnet     = $_.Subnet
        DefaultGateway    = ($_.DefaultGateways) -join ', '
        DNS        = ($_.DnsServers) -join ', '
        InUse      = $_.TotalAddresses - $_.AvailableAddresses
        Available  = $_.AvailableAddresses
    }
}


$macPoolRows = $macPools | ForEach-Object {
    [pscustomobject]@{
        MACPool   = $_.Name
        StartMAC  = $_.MACAddressRangeStart
        EndMAC    = $_.MACAddressRangeEnd
        Allocated = $_.TotalAddresses - $_.AvailableAddresses
        Available = $_.AvailableAddresses
    }
}
$lswRows = $logicalSwitches | ForEach-Object {
    $uplinks = Get-SCUplinkPortProfileSet -LogicalSwitch $_
    $vnic = (Get-SCLogicalSwitchVirtualNetworkAdapter -UplinkPortProfileSet $uplinks | ForEach-Object Name) -join ', '
    [pscustomobject]@{
        LogicalSwitch   = $_.Name
        UplinkProfile   = $uplinks
        VirtualNetworkAdapters = $vnic
        PortClass       = ($_.PortClassifications | ForEach-Object Name) -join ', ' #Need Review
        CompliantHosts  = ($_.CompliantVMHosts | ForEach-Object Name) -join ', '    #Need Review
    }
}

# ------------------- Host NIC–vSwitch–Logical Network mapping -------------------
$nicMapRows = New-Object System.Collections.Generic.List[object]
foreach ($h in $hosts) {
    $nics = Get-SCVMHostNetworkAdapter -VMHost $h -ErrorAction SilentlyContinue
    foreach ($nic in $nics) {
        $vsObj = Get-Prop $nic 'VirtualSwitch'          #Need Review
        $lsObj = Get-Prop $nic 'LogicalSwitch'          #Need Review
        $lns   = $nic.LogicalNetworks

        $vsName = if ($vsObj) { $vsObj.Name } else { $null }
        $lsName = if ($lsObj) { $lsObj.Name } else { $null }
        $connName = if ($null -ne $nic.ConnectionName) { $nic.ConnectionName } else { $nic.Name }
        $boundLn  = if ($lns) { ($lns | ForEach-Object Name) -join ', ' } else { $null }

        $nicMapRows.Add([pscustomobject][ordered]@{
            Host            = $h.Name
            AdapterName     = $connName
            VirtualNetwork         = $vsName
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
            catch { Write-Warning ("[WARN] Failed to refresh LLDP: {0}" -f $_.Exception.Message) }
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
            $agentVersion = ($resp.Version).Trim()
        } catch { $wsmanErr = "AgentVersion error: $($_.Exception.Message)"; $agentVersion=$null }
    } catch {
        [xml]$rawXMLError = $_.Exception.Message
        # Map the namespace used by WSManFault
        $ns = @{ f = "http://schemas.microsoft.com/wbem/wsman/1/wsmanfault" }
        
        $wsmanErr = (Select-Xml -Xml $rawXMLError -XPath "//f:Message" -Namespace $ns).Node.InnerText
        $wsmanOk = $false
    }
    [pscustomobject]@{ WSManOk=$wsmanOk; AgentVersion=$agentVersion; WSManError=$wsmanErr }
}

function Test-WMI {
    param([string]$ComputerName,[System.Management.Automation.PSCredential]$Credential)
    # Try both CIM (WS-Man) and DCOM paths
    $wsCimOk=$false; $dcCimOk=$false; $wsCimerr=$null; $dcomerr=$null
    try { 
        $null = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop
        $wsCimOk=$true 
    } catch { $wsCimerr = $_.Exception.Message; $wsCimOk=$false }
    try { 
        $opt = New-CimSessionOption -Protocol Dcom
        $sess = New-CimSession -ComputerName $ComputerName -Credential $Credential -SessionOption $opt -ErrorAction Stop
        $null = Get-CimInstance -CimSession $sess -ClassName Win32_OperatingSystem -ErrorAction Stop
        $dcCimOk=$true
        $sess | Remove-CimSession
    } catch {
        $dcomerr = $_.Exception.MessageId + "`n" + $_.Exception.Message
        $dcCimOk=$false 
    }
    [pscustomobject]@{ WMI_WSMan=$wsCimOk; WMI_DCOM=$dcCimOk; WMI_WSMan_Error=$wsCimerr; WMI_DCOM_Error=$dcomerr }
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
        $pingMs = if ($null -ne $avg) { [Math]::Round($avg,1) } else { $null }
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


<# =======================================================================
SCVMM report.txt Severity Classification Policy (English)

Overall rule:
- We assign a 4-level severity based on regex/keyword hits in each line.
- If multiple patterns match, the line is promoted to the highest matched severity.
- Priority: Critical (4) > High (3) > Medium (2) > Low (1). Empty lines get 0.

Levels and examples:
1) Critical (4): Service or platform-wide failures, fatal conditions, or HRESULTs
   - Examples: 0x8xxxxxxx (e.g., 0x8033802A), "Fatal", "Database is unavailable",
               "connection refused", "SSL/TLS handshake failure",
               "VMM service ... failed/unavailable", agent authentication/certificate failures

2) High (3): Operation-level failures that can block important tasks
   - Examples: "ErrorCode=<digits>", "Exception:", "Agent not responding",
               "Access is denied", "RPC server is unavailable", "timeout",
               "WinRM cannot complete the operation"

3) Medium (2): Performance or environmental warnings that may cause delays
   - Examples: "Warning:", "Retrying", "throttling", "insufficient resources",
               "rate limit"

4) Low (1): Informational or minor notices
   - Examples: "Info:", "Note:", "Succeeded with warnings", "skipped"

Extensibility:
- Add per-environment keywords (English/Korean) to the matcher.
- Consider collapsing duplicates by HResult/ErrorCode for top-N summaries.
======================================================================= #>


# ---------- VMM report.txt collector + severity classifier (PS 5.1 safe) ----------

# Returns numeric score: 0(None), 1(Low), 2(Medium), 3(High), 4(Critical)
function Get-SeverityScore {
    param([string]$Text)

    # Default Low if we have any content; 0 will be used for blanks
    $score = 1

    if ($null -eq $Text -or $Text.Trim().Length -eq 0) { return 0 }

    # Normalize for case-insensitive keyword search
    $t = $Text.ToLowerInvariant()

    # --- Critical (4): fatal/system-wide issues or 0x8xxxxxxx HRESULTs ---
    # Examples: 0x8033802A, "fatal", "database is unavailable", "connection refused",
    # "ssl/tls handshake failure", "vmm service ... failed", agent auth/cert failures
    if ($t -match '0x8[0-9a-f]{7}\b' -or
        $t -match '\bfatal\b' -or
        $t -match 'database is unavailable' -or
        $t -match 'connection refused' -or
        $t -match 'ssl|tls.*(handshake|failure)' -or
        $t -match 'vmm service.*(stopped|unavailable|failed)' -or
        $t -match 'agent.*(authentication|certificate).*(fail|invalid)') {
        return 4
    }

    # --- High (3): blocking operation failures / hard errors ---
    # Examples: "ErrorCode=####", "Exception:", "Agent not responding",
    # "Access is denied", "RPC server is unavailable", "timeout",
    # "WinRM cannot complete the operation"
    if ($t -match 'errorcode=\d+' -or
        $t -match '\bexception\b' -or
        $t -match 'agent not responding' -or
        $t -match 'access is denied' -or
        $t -match 'rpc server is unavailable' -or
        $t -match '\btimeout\b' -or
        $t -match 'winrm.*cannot complete the operation') {
        # use helper to keep the max of current vs target level (3)
        $score = [Math]::Max($score, 3)
    }

    # --- Medium (2): degradations / warnings / transient throttling ---
    # Examples: "Warning:", "Retrying", "throttling", "insufficient resources", "rate limit"
    if ($t -match '\bwarning\b' -or
        $t -match 'retrying' -or
        $t -match 'throttling' -or
        $t -match 'insufficient resources' -or
        $t -match 'rate limit') {
        $score = [Math]::Max($score, 2)
    }

    # --- Low (1): informational or minor notices ---
    # Examples: "Info:", "Note:", "Succeeded with warnings", "skipped"
    return $score
}

# Maps numeric score to label; used for HTML/CSV output
function Get-SeverityLabel {
    param([int]$Score)
    switch ($Score) {
        4 { 'Critical' }
        3 { 'High' }
        2 { 'Medium' }
        1 { 'Low' }
        default { 'None' }
    }
}

# Parse C:\ProgramData\VMMLogs\report.txt and extract fields/patterns
# Returns a list of PSCustomObject with (Line, Timestamp, HResult, ErrorCode, Exception, Message, Severity, SeverityScore)
function Parse-VmmReportTxt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Path
    )

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($file in $Path) {

        if (-not (Test-Path -LiteralPath $file)) {
            continue
        }

        # --- Determine report-level timestamp ---
        $reportTime = $null

        # Try to extract "Error report created <date time>" from header
        $headerLines = Get-Content -LiteralPath $file -TotalCount 20 -ErrorAction SilentlyContinue
        foreach ($h in $headerLines) {
            if ($h -match 'Error report created\s+(.+)$') {
                try {
                    $reportTime = [datetime]::Parse($matches[1])
                } catch {
                    $reportTime = $null
                }
                break
            }
        }

        # Fallback to file LastWriteTime if header timestamp is not available
        if (-not $reportTime) {
            try {
                $reportTime = (Get-Item -LiteralPath $file).LastWriteTime
            } catch {
                $reportTime = $null
            }
        }

        # --- Parse file content ---
        $ln = 0
        Get-Content -LiteralPath $file -ErrorAction Continue | ForEach-Object {
            $ln += 1
            $line = $_

            if ($null -eq $line -or $line.Trim().Length -eq 0) {
                return
            }

            # Extract fields (existing logic)
            $hr = $null
            $ecode = $null
            $ex = $null

            if ($line -match '(0x[0-9a-fA-F]{8})') { $hr = $matches[1] }
            if ($line -match 'ErrorCode=(\d+)') { $ecode = $matches[1] }
            if ($line -match '(?i)exception:? ([^;]+)') { $ex = $matches[1].Trim() }

            $score = Get-SeverityScore -Text $line

            $rows.Add([pscustomobject]@{
                Line          = $ln
                Timestamp     = $reportTime
                HResult       = $hr
                ErrorCode     = $ecode
                Exception     = $ex
                Message       = $line
                Severity      = Get-SeverityLabel -Score $score
                SeverityScore = $score
                SourceFile    = $file
                SourceFolder  = Split-Path -Path $file -Parent
            }) | Out-Null
        }
    }

    return $rows
}

# ---------- end: VMM report.txt collector ----------

# Build HTML fragment with severity styling
function To-HtmlSection-VmmReport {
    param($Rows)

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        return "<h2>VMM Error Log (report.txt)</h2><div class='small'>No data</div>"
    }

    # Top badges summary
    $countCritical = ($Rows | Where-Object { $_.Severity -eq 'Critical' }).Count
    $countHigh     = ($Rows | Where-Object { $_.Severity -eq 'High' }).Count
    $countMedium   = ($Rows | Where-Object { $_.Severity -eq 'Medium' }).Count
    $countLow      = ($Rows | Where-Object { $_.Severity -eq 'Low' }).Count

    $summary = "<div class='small'>
    <span class='badge badge-critical'>Critical: $countCritical</span>
    <span class='badge badge-high'>High: $countHigh</span>
    <span class='badge badge-medium'>Medium: $countMedium</span>
    <span class='badge badge-low'>Low: $countLow</span>
    </div>"

    # Render table (manual to attach row classes)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<h2>VMM Error Log (report.txt)</h2>$summary<div class='tablebox'><table><thead><tr>
        <th>#</th><th>Time</th><th>Severity</th><th>HResult</th><th>ErrorCode</th><th>Exception</th><th>Message</th>
    </tr></thead><tbody>")

    foreach ($r in $Rows) {
        $cls = "sev-$($r.Severity)"
        $msgEsc = [System.Web.HttpUtility]::HtmlEncode([string]$r.Message)
        $exEsc  = [System.Web.HttpUtility]::HtmlEncode([string]$r.Exception)
        $hrEsc  = [System.Web.HttpUtility]::HtmlEncode([string]$r.HResult)
        $ecEsc  = [System.Web.HttpUtility]::HtmlEncode([string]$r.ErrorCode)
        $tEsc   = [System.Web.HttpUtility]::HtmlEncode([string]$r.Timestamp)

        [void]$sb.Append("<tr class='$cls'>
            <td>$($r.Line)</td>
            <td>$tEsc</td>
            <td>$($r.Severity)</td>
            <td>$hrEsc</td>
            <td>$ecEsc</td>
            <td>$exEsc</td>
            <td class='code'>$msgEsc</td>
        </tr>")
    }

    [void]$sb.Append("</tbody></table></div>")
    return $sb.ToString()
}


# =================== NEW: report.txt discovery across SCVMM.* folders ===================

function Get-RecentVmmReportFiles {
    [CmdletBinding()]
    param(
        [string]$Root = 'C:\ProgramData\VMMLogs',
        [string]$ChildPattern = 'SCVMM.*',
        [string]$ReportName = 'report.txt',
        [int]$ModifiedWithinDays = 7,

        # Fallback option:
        # Used only when there are NO report.txt files modified within the last N days.
        # If set to 10, return the most recent 10 report.txt files by LastWriteTime.
        # If set to 0 (or negative), return ALL report.txt files sorted by LastWriteTime desc.
        [int]$FallbackTop = 10
    )

    # Calculate the cutoff timestamp (e.g., "now - 7 days")
    $cutoff = (Get-Date).AddDays(-[math]::Abs($ModifiedWithinDays))

    # If the root path doesn't exist, return an empty list
    if (-not (Test-Path $Root)) { return @() }

    # 1) Enumerate SCVMM.* folders and collect report.txt files under each folder
    $allReports = @()
    $folders = Get-ChildItem -Path $Root -Directory -Filter $ChildPattern -ErrorAction SilentlyContinue

    foreach ($f in $folders) {
        $p = Join-Path $f.FullName $ReportName
        if (Test-Path $p) {
            $allReports += Get-Item -LiteralPath $p -ErrorAction SilentlyContinue
        }
    }

    # If no report.txt files were found at all, return an empty list
    if (-not $allReports -or $allReports.Count -eq 0) {
        return @()
    }

    # 2) Primary filter:
    # Return only report.txt files whose LastWriteTime is within the last N days.
    $recent = $allReports | Where-Object { $_.LastWriteTime -ge $cutoff }

    if ($recent -and $recent.Count -gt 0) {
        return $recent
    }

    # 3) Fallback behavior:
    # If there are NO report.txt files modified within the last N days,
    # then return the most recent report.txt files by LastWriteTime.
    $sorted = $allReports | Sort-Object LastWriteTime -Descending

    if ($FallbackTop -le 0) {
        # If FallbackTop is 0 or negative, return ALL files (sorted)
        return $sorted
    } else {
        # Otherwise return Top N most recent files
        return ($sorted | Select-Object -First $FallbackTop)
    }
}


# NEW: HTML rendering with Source columns shown and per-folder badges
function To-HtmlSection-VmmReport {
    param($Rows)

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        return "<h2>VMM Error Log (report.txt, merged)</h2><div class='small'>No data</div>"
    }

    # Overall counts
    $countCritical = ($Rows | Where-Object { $_.Severity -eq 'Critical' }).Count
    $countHigh     = ($Rows | Where-Object { $_.Severity -eq 'High'     }).Count
    $countMedium   = ($Rows | Where-Object { $_.Severity -eq 'Medium'   }).Count
    $countLow      = ($Rows | Where-Object { $_.Severity -eq 'Low'      }).Count

    $summary = "<div class='small'>
    <span class='badge badge-critical'>Critical: $countCritical</span>
    <span class='badge badge-high'>High: $countHigh</span>
    <span class='badge badge-medium'>Medium: $countMedium</span>
    <span class='badge badge-low'>Low: $countLow</span>
</div>"

    # Per-source mini-summary
    $perFolder = ($Rows | Group-Object SourceFolder | ForEach-Object {
        $f = $_.Name
        $c4 = ($_.Group | Where-Object Severity -eq 'Critical').Count
        $c3 = ($_.Group | Where-Object Severity -eq 'High').Count
        $c2 = ($_.Group | Where-Object Severity -eq 'Medium').Count
        $c1 = ($_.Group | Where-Object Severity -eq 'Low').Count
        "<div class='small'><span class='badge'>$f</span> <span class='badge badge-critical'>$c4</span> <span class='badge badge-high'>$c3</span> <span class='badge badge-medium'>$c2</span> <span class='badge badge-low'>$c1</span></div>"
    }) -join "`n"

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<h2>VMM Error Log (report.txt, merged)</h2>$summary$perFolder<div class='tablebox'><table><thead><tr>
<th>#</th><th>Time</th><th>Severity</th><th>HResult</th><th>ErrorCode</th><th>Exception</th><th>Message</th><th>SourceFolder</th>
</tr></thead><tbody>")

    foreach ($r in $Rows) {
        $cls   = "sev-$($r.Severity)"
        $msg   = [System.Web.HttpUtility]::HtmlEncode([string]$r.Message)
        $exEsc = [System.Web.HttpUtility]::HtmlEncode([string]$r.Exception)
        $hrEsc = [System.Web.HttpUtility]::HtmlEncode([string]$r.Hikari)
        $hrEsc = [System.Web.HttpUtility]::HtmlEncode([string]$r.HResult) # fix
        $ecEsc = [System.Web.HttpUtility]::HtmlEncode([string]$r.ErrorCode)
        $tEsc  = [System.Web.HttpUtility]::HtmlEncode([string]$r.Timestamp)
        $src   = [System.Web.HttpUtility]::HtmlEncode([string]$r.SourceFolder)

        [void]$sb.Append("<tr class='$cls'>
<td>$($r.Line)</td>
<td>$tEsc</td>
<td>$($r.Severity)</td>
<td>$hrEsc</td>
<td>$ecEsc</td>
<td>$exEsc</td>
<td class='code'>$msg</td>
<td>$src</td>
</tr>")
    }
    [void]$sb.Append("</tbody></table></div>")
    return $sb.ToString()
}

function New-VmmReportSummaryBadge {
    param(
        [Parameter(Mandatory)]
        [object[]]$ReportRows
    )

    if (-not $ReportRows -or $ReportRows.Count -eq 0) {
        return "<div class='small'>No report.txt data</div>"
    }

    # Use the earliest Report Created Time as representative
    $times = $ReportRows |
        Where-Object { $_.Timestamp } |
        Select-Object -ExpandProperty Timestamp

    $minTime = ($times | Sort-Object | Select-Object -First 1)

    $sourceCount = ($ReportRows |
        Select-Object -ExpandProperty SourceFolder -Unique).Count

    return @"
<div class='tablebox'>
  <span class='badge badge-high'>Report Created Time</span>
  <span class='badge'>$($minTime.ToString('yyyy-MM-dd HH:mm:ss'))</span>
  <span class='badge badge-low'>Sources: $sourceCount report.txt files</span>
</div>
"@
}

function To-HtmlSection-VmmReport-Grouped {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        return "<h2>VMM Error Log (report.txt)</h2><div class='small'>No data</div>"
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<h2>VMM Error Log (report.txt)</h2>")

    # Group by SourceFolder + Timestamp
    $groups = $Rows | Group-Object {
        "{0}|{1}" -f $_.SourceFolder, $_.Timestamp
    }

    foreach ($g in $groups) {

        $sample = $g.Group | Select-Object -First 1
        $folder = $sample.SourceFolder
        $time   = if ($sample.Timestamp) {
            $sample.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')
        } else {
            'Unknown'
        }
	$sevOrder = @{
	    'Critical' = 1
	    'High'     = 2
	    'Medium'   = 3
	    'Low'      = 4
	    'None'     = 5
	}
	
	# Severity summary per group
	$sevSummary = (
	    $g.Group |
	        Group-Object Severity |
	        Sort-Object { $sevOrder[$_.Name] } |
	        ForEach-Object { "$($_.Name): $($_.Count)" }
	) -join ', '
       
        [void]$sb.Append(@"
<h3>Source: $folder</h3>
<div class='small'>
  <span class='badge'>$time</span>
  <span class='badge badge-medium'>$sevSummary</span>
</div>
<div class='tablebox'>
<table>
<thead>
<tr>
  <th>#</th>
  <th>Time</th>
  <th>Severity</th>
  <th>HResult</th>
  <th>ErrorCode</th>
  <th>Exception</th>
  <th>Message</th>
</tr>
</thead>
<tbody>
"@)

        foreach ($r in $g.Group) {
            $cls = "sev-$($r.Severity)"

            $msg = [System.Web.HttpUtility]::HtmlEncode($r.Message)
            $ex  = [System.Web.HttpUtility]::HtmlEncode($r.Exception)
            $hr  = [System.Web.HttpUtility]::HtmlEncode($r.HResult)
            $ec  = [System.Web.HttpUtility]::HtmlEncode($r.ErrorCode)
            $ts  = if ($r.Timestamp) {
                $r.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')
            } else { '' }

            [void]$sb.Append(@"
<tr class='$cls'>
  <td>$($r.Line)</td>
  <td>$ts</td>
  <td>$($r.Severity)</td>
  <td>$hr</td>
  <td>$ec</td>
  <td>$ex</td>
  <td class='code'>$msg</td>
</tr>
"@)
        }

        [void]$sb.Append("</tbody></table></div>")
    }

    return $sb.ToString()
}

# ------------------- Save CSVs -------------------

# Collect VMM report.txt rows
# Recent 7 days SCVMM.*\report.txt
$reportFiles = Get-RecentVmmReportFiles -Root 'C:\ProgramData\VMMLogs' -ChildPattern 'SCVMM.*' -ReportName 'report.txt' -ModifiedWithinDays 7

# Parsing the files (SourceFolder/SourceFile included)
$reportRows  = Parse-VmmReportTxt -Path ($reportFiles.FullName)

# CSV save for the files
To-Csv $reportRows (Join-Path $OUT 'report_log.csv')

# Update HTML section
#$sections += To-HtmlSection-VmmReport -Rows $reportRows
$sections += To-HtmlSection-VmmReport-Grouped -Rows $reportRows

# Save CSV
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
To-Csv $reportRows (Join-Path $OUT 'report_log.csv')

# ------------------- HTML report -------------------
$sections = @()
$sections += "<h2>SCVMM Diagnostic report</h2><div class='small'>Creation Time: $(Get-Date) / Server: $($serverInfo[0].FQDN) / Created By: $($Credential.UserName)</div>"
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
# --- Report.txt summary badge (TOP of HTML) ---
$sections += New-VmmReportSummaryBadge -ReportRows $reportRows

$sections += To-HtmlSection-VmmReport -Rows $reportRows


$html = ConvertTo-Html -Head (Html-Style) -Body ($sections -join "`n")
$null = $html | Set-Content -Path $HTMLpath -Encoding UTF8

Write-Host "[INFO] Report generated: $HTMLpath"
Write-Host "[INFO] Source CSV folder: $OUT"
Write-Host "[TIP ] Use -RefreshLLDP to refresh LLDP details if you need latest switch/port info."
