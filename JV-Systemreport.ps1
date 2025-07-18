# Check for Administrator rights
If (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script must be run as Administrator." -ForegroundColor Red
    Exit 1
}

# Output path
$ScriptPath = $MyInvocation.MyCommand.Path
$OutputFolder = Split-Path $ScriptPath
$ReportPath = Join-Path $OutputFolder "report.html"

function Convert-ToHtmlTable {
    param ([object[]]$Data)
    if (-not $Data) {
        return "<div style='color: red; font-weight: bold;'>Not Installed</div>"
    }
    return $Data | ConvertTo-Html -Fragment | Out-String
}

# System Info via systeminfo
$SysInfoRaw = systeminfo | ForEach-Object {
    if ($_ -match "^(.*?):\s+(.*)$") {
        [PSCustomObject]@{
            Property = $matches[1].Trim()
            Value    = $matches[2].Trim()
        }
    }
}
$SystemInfoHtml = Convert-ToHtmlTable $SysInfoRaw

# Network info (correcte IP weergave + IPv6 per interface)
$Adapters = Get-NetIPConfiguration | ForEach-Object {
    $ipv4 = $_.IPv4Address.IPAddress -join ', '
    $ipv6 = $_.IPv6Address.IPAddress -join ', '
    $dns  = $_.DNSServer.ServerAddresses -join ', '
    $iface = $_.InterfaceAlias
    $ipv6enabled = (Get-NetAdapterBinding -InterfaceAlias $iface -ComponentID ms_tcpip6).Enabled

    [PSCustomObject]@{
        InterfaceAlias       = $iface
        IPv4Address          = $ipv4
        IPv6Address          = $ipv6
        DNSServer            = $dns
        IPv6Enabled          = if ($ipv6enabled) { "Yes" } else { "No" }
        InterfaceDescription = $_.InterfaceDescription
    }
}
$NetworkHtml = Convert-ToHtmlTable $Adapters

# Firewall status overview
$FWStatus = Get-NetFirewallProfile | Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction
$FirewallHtml = "<h3>Firewall Profiles</h3>" + (Convert-ToHtmlTable $FWStatus)

# Custom firewall rules with port info
$CustomRules = Get-NetFirewallRule |
    Where-Object { $_.Enabled -eq "True" -and $_.PolicyStoreSourceType -eq "Local" } |
    Sort-Object DisplayName |
    ForEach-Object {
        $rule = $_
        $filter = Get-NetFirewallPortFilter -AssociatedNetFirewallRule $rule

        [PSCustomObject]@{
            Name        = $rule.DisplayName
            Direction   = $rule.Direction
            Action      = $rule.Action
            Profile     = $rule.Profile
            Protocol    = $filter.Protocol
            LocalPort   = if ($filter.LocalPort -eq 'Any') { 'Any' } else { $filter.LocalPort }
            RemotePort  = if ($filter.RemotePort -eq 'Any') { 'Any' } else { $filter.RemotePort }
        }
    }

$FirewallHtml += "<h3>Custom Enabled Rules (Local Only)</h3>" + (Convert-ToHtmlTable $CustomRules)

# Storage
$Drives = Get-PSDrive -PSProvider 'FileSystem' | Select-Object Name, @{Name='Free(GB)';Expression={[math]::Round($_.Free/1GB,2)}}, @{Name='Used(GB)';Expression={[math]::Round(($_.Used/1GB),2)}}, @{Name='Total(GB)';Expression={[math]::Round($_.Used/1GB + $_.Free/1GB,2)}}
$StorageHtml = Convert-ToHtmlTable $Drives

# Placeholder tabs
$PlaceholderHtml = "<div style='color: red; font-weight: bold;'>Not Implemented Yet</div>"

# Tabs
$Tabs = @(
    @{ Name = "System_Info"; Label = "System Info"; Content = $SystemInfoHtml },
    @{ Name = "Network"; Content = $NetworkHtml },
    @{ Name = "Firewall"; Content = $FirewallHtml },
    @{ Name = "Storage"; Content = $StorageHtml },
    @{ Name = "Applications"; Content = $PlaceholderHtml },
    @{ Name = "Server_Roles"; Label = "Server Roles"; Content = $PlaceholderHtml },
    @{ Name = "Shares"; Content = $PlaceholderHtml },
    @{ Name = "Printers"; Content = $PlaceholderHtml }
)

# HTML layout
$HtmlHeader = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>System Report</title>
<style>
body { font-family: Arial; margin: 20px; }
.tab { display: none; padding: 10px; border: 1px solid #ccc; border-top: none; }
.tab-buttons button { background: #eee; border: none; padding: 10px; cursor: pointer; }
.tab-buttons button.active { background: #ccc; }
table { border-collapse: collapse; width: 100%; margin-top: 10px; }
th, td { border: 1px solid #ddd; padding: 8px; vertical-align: top; }
tr:nth-child(even) { background-color: #f2f2f2; }
</style>
<script>
function showTab(name) {
    var tabs = document.getElementsByClassName('tab');
    for (var t of tabs) { t.style.display = 'none'; }
    var btns = document.getElementsByClassName('tab-btn');
    for (var b of btns) { b.classList.remove('active'); }
    document.getElementById(name).style.display = 'block';
    document.getElementById('btn_' + name).classList.add('active');
}
window.onload = function() {
    showTab('System_Info');
};
</script>
</head>
<body>
<h1>System Report</h1>
<div class='tab-buttons'>
"@

# Knoppen
$HtmlButtons = ""
foreach ($tab in $Tabs) {
    $label = if ($tab.Label) { $tab.Label } else { $tab.Name }
    $HtmlButtons += "<button class='tab-btn' id='btn_$($tab.Name)' onclick=`"showTab('$($tab.Name)')`">$label</button>`n"
}

# Tab inhoud
$HtmlTabs = ""
foreach ($tab in $Tabs) {
    $HtmlTabs += "<div id='$($tab.Name)' class='tab'>$($tab.Content)</div>`n"
}

$HtmlFooter = "</body></html>"

# Save report
$FinalHtml = $HtmlHeader + $HtmlButtons + "</div>" + $HtmlTabs + $HtmlFooter
$FinalHtml | Out-File -FilePath $ReportPath -Encoding UTF8

Write-Host "Report generated at: $ReportPath" -ForegroundColor Green