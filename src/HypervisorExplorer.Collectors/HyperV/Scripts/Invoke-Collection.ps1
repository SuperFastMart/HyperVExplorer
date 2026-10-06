<#
.SYNOPSIS
    Local launcher used by Hypervisor Explorer to run Collect-HyperV.ps1 against a Hyper-V host.

.DESCRIPTION
    Receives the request object (decoded from stdin by the launcher bootstrap; credentials are never on
    the command line or in the environment). Runs the collection script locally when the target is this
    computer, otherwise with Invoke-Command over WinRM. When the host is a Failover Cluster node and
    expandCluster is set, every other Up/Paused node is collected in parallel and merged.

    Output: exactly one single-line JSON envelope on stdout:
        {"ok":true,"data":{"schemaVersion":1,"kind":"multi","nodes":[...],"warnings":[...]}}
        {"ok":false,"stage":"connect|auth|collect","error":"...","category":"..."}
    Progress: "PROGRESS: text" lines on stderr.

    Request fields: computer, port, useSsl, skipCertificateCheck, user, password, authentication,
    expandCluster, nodeTimeoutSeconds, script (the text of Collect-HyperV.ps1).
#>
param([Parameter(Mandatory = $true)]$Request)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
# Remote and job verbose records carry the PROGRESS lines; they are captured with 4>&1 below.
$VerbosePreference = 'Continue'

$utf8 = New-Object System.Text.UTF8Encoding($false)
$stdout = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $utf8)
$stderr = New-Object System.IO.StreamWriter([Console]::OpenStandardError(), $utf8)
$stderr.AutoFlush = $true
$warnings = New-Object 'System.Collections.Generic.List[string]'

function Send-Progress([string]$Text) {
    try { $stderr.WriteLine('PROGRESS: ' + $Text) } catch { }
}

function Write-Envelope([string]$Json) {
    $stdout.WriteLine($Json)
    $stdout.Flush()
}

function ConvertTo-JsonString([string]$Text) {
    if ($null -eq $Text) { return 'null' }
    return (ConvertTo-Json -InputObject $Text -Compress)
}

function Send-Failure([string]$Stage, $Err) {
    $msg = ''
    $cat = ''
    if ($Err -is [System.Management.Automation.ErrorRecord]) {
        $msg = [string]$Err.Exception.Message
        $cat = [string]$Err.FullyQualifiedErrorId
        if ($Err.CategoryInfo) { $cat = $cat + ' | ' + $Err.CategoryInfo.ToString() }
    } else {
        $msg = [string]$Err
    }
    $o = [ordered]@{ ok = $false; stage = $Stage; error = $msg; category = $cat }
    Write-Envelope (ConvertTo-Json -InputObject $o -Compress)
}

function Get-FailureStage($Err, [string]$Default) {
    $t = [string]$Err.Exception.Message + ' ' + [string]$Err.FullyQualifiedErrorId
    if ($t -match 'Access is denied|AccessDenied|logon failure|LogonFailure|user name or password|0x80090311|0x8009030e|0x8009030c') { return 'auth' }
    return $Default
}

# Forwards PROGRESS verbose records to stderr and passes ordinary output through as strings.
function Select-CollectorOutput {
    process {
        $item = $_
        if ($item -is [System.Management.Automation.VerboseRecord]) {
            if ($item.Message -like 'PROGRESS:*') { Send-Progress ($item.Message.Substring(9).Trim()) }
        } elseif ($item -is [System.Management.Automation.WarningRecord] -or $item -is [System.Management.Automation.DebugRecord]) {
        } elseif ($null -ne $item) {
            [string]$item
        }
    }
}

function Test-IsLocalComputer([string]$Name) {
    if (-not $Name) { return $true }
    $n = $Name.Trim().TrimEnd('.')
    $local = @('localhost', '.', '127.0.0.1', '::1', $env:COMPUTERNAME)
    try { $local += [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName } catch { }
    foreach ($l in $local) {
        if ($l -and [string]::Equals($n, [string]$l, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

# ---------------------------------------------------------------- request

try {
    $target = ([string]$Request.computer).Trim()
    $collectText = [string]$Request.script
    if (-not $collectText) { throw 'No collection script was supplied.' }
    $collect = [scriptblock]::Create($collectText)

    $icm = @{ ErrorAction = 'Stop' }
    if ($Request.port) { $icm['Port'] = [int]$Request.port }
    if ($Request.useSsl) {
        $icm['UseSSL'] = $true
        if ($Request.skipCertificateCheck) {
            $icm['SessionOption'] = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
        }
    }
    if ($Request.user) {
        $sec = New-Object System.Security.SecureString
        foreach ($ch in ([string]$Request.password).ToCharArray()) { $sec.AppendChar($ch) }
        $sec.MakeReadOnly()
        $icm['Credential'] = New-Object System.Management.Automation.PSCredential([string]$Request.user, $sec)
        $auth = 'Negotiate'
        if ($Request.authentication) { $auth = [string]$Request.authentication }
        $icm['Authentication'] = $auth
    } elseif ($Request.authentication) {
        $icm['Authentication'] = [string]$Request.authentication
    }

    $isLocal = Test-IsLocalComputer $target
    $nodeTimeout = 900
    if ($Request.nodeTimeoutSeconds) { $nodeTimeout = [int]$Request.nodeTimeoutSeconds }
} catch {
    Send-Failure 'connect' $_
    return
}

# ---------------------------------------------------------------- discovery (also proves connectivity + auth)

$discover = {
    $lines = @([string]$env:COMPUTERNAME)
    if (Get-Command -Name Get-ClusterNode -ErrorAction SilentlyContinue) {
        try {
            foreach ($n in @(Get-ClusterNode -ErrorAction Stop)) { $lines += ([string]$n.Name + '|' + [string]$n.State) }
        } catch { }
    }
    $lines
}

try {
    if ($isLocal) {
        Send-Progress ('Collecting the local computer (' + $env:COMPUTERNAME + ')')
        $disc = @(& $discover)
    } else {
        Send-Progress ('Connecting to ' + $target + ' over WinRM')
        $disc = @(Invoke-Command @icm -ComputerName $target -ScriptBlock $discover)
    }
} catch {
    Send-Failure (Get-FailureStage $_ 'connect') $_
    return
}

$primaryName = [string]$disc[0]
$others = @()
foreach ($line in @($disc | Select-Object -Skip 1)) {
    $parts = ([string]$line).Split('|')
    $nodeName = $parts[0]
    $state = ''
    if ($parts.Length -gt 1) { $state = $parts[1] }
    if (-not $Request.expandCluster) { continue }
    if ([string]::Equals($nodeName, $primaryName, [StringComparison]::OrdinalIgnoreCase)) { continue }
    if ($state -eq 'Up' -or $state -eq 'Paused') { $others += $nodeName }
    else { $warnings.Add('Cluster node ' + $nodeName + ' is ' + $state + '; not collected.') }
}

# ---------------------------------------------------------------- other cluster nodes (parallel, background)

$job = $null
if ($others.Count -gt 0) {
    Send-Progress ('Failover cluster detected; collecting ' + $others.Count + ' other node(s) in parallel: ' + ($others -join ', '))
    $jp = @{}
    foreach ($k in $icm.Keys) { $jp[$k] = $icm[$k] }
    $jp['ErrorAction'] = 'Continue'
    try {
        $job = Invoke-Command @jp -ComputerName $others -ScriptBlock $collect -ArgumentList $primaryName -ThrottleLimit 8 -AsJob
    } catch {
        $warnings.Add('Could not start collection of other cluster nodes: ' + $_.Exception.Message)
        $job = $null
    }
}

# ---------------------------------------------------------------- primary host

$primaryOut = New-Object 'System.Collections.Generic.List[string]'
try {
    if ($isLocal) {
        & $collect 4>&1 | Select-CollectorOutput | ForEach-Object { $primaryOut.Add($_) }
    } else {
        Invoke-Command @icm -ComputerName $target -ScriptBlock $collect 4>&1 | Select-CollectorOutput | ForEach-Object { $primaryOut.Add($_) }
    }
} catch {
    if ($job) { try { Stop-Job -Job $job; Remove-Job -Job $job -Force } catch { } }
    Send-Failure (Get-FailureStage $_ 'collect') $_
    return
}

$primaryJson = $null
foreach ($o in $primaryOut) { if ($o.TrimStart().StartsWith('{')) { $primaryJson = $o.Trim() } }
if (-not $primaryJson) {
    if ($job) { try { Stop-Job -Job $job; Remove-Job -Job $job -Force } catch { } }
    Send-Failure 'collect' ('The collection script returned no data from ' + $primaryName + '.')
    return
}

$nodeJsons = New-Object 'System.Collections.Generic.List[string]'
$nodeJsons.Add($primaryJson)

if ($job) {
    $jobErrors = @()
    $deadline = (Get-Date).AddSeconds($nodeTimeout)
    while ($true) {
        $errs = $null
        try {
            Receive-Job -Job $job -ErrorAction SilentlyContinue -ErrorVariable errs 4>&1 | Select-CollectorOutput |
                ForEach-Object { if ($_.TrimStart().StartsWith('{')) { $nodeJsons.Add($_.Trim()) } }
        } catch { $warnings.Add('Cluster node collection: ' + $_.Exception.Message) }
        if ($errs) { $jobErrors += @($errs) }
        $st = [string]$job.State
        if ($st -ne 'Running' -and $st -ne 'NotStarted') { break }
        if ((Get-Date) -gt $deadline) {
            foreach ($c in @($job.ChildJobs)) {
                if ([string]$c.State -eq 'Running') { $warnings.Add('Cluster node ' + $c.Location + ' timed out after ' + $nodeTimeout + ' s; not collected.') }
            }
            try { Stop-Job -Job $job } catch { }
            break
        }
        Start-Sleep -Milliseconds 500
    }
    foreach ($e in $jobErrors) {
        if ($null -eq $e) { continue }
        $who = ''
        try { if ($e.OriginInfo -and $e.OriginInfo.PSComputerName) { $who = '[' + $e.OriginInfo.PSComputerName + '] ' } } catch { }
        $m = ''
        if ($e -is [System.Management.Automation.ErrorRecord]) { $m = [string]$e.Exception.Message } else { $m = [string]$e }
        $warnings.Add('Cluster node collection failed: ' + $who + $m)
    }
    try { Remove-Job -Job $job -Force } catch { }
}

# ---------------------------------------------------------------- envelope

$warnParts = @()
foreach ($w in $warnings) { $warnParts += (ConvertTo-JsonString $w) }
$data = '{"schemaVersion":1,"kind":"multi","requestedComputer":' + (ConvertTo-JsonString $target) +
    ',"primary":' + (ConvertTo-JsonString $primaryName) +
    ',"nodes":[' + ($nodeJsons -join ',') + '],"warnings":[' + ($warnParts -join ',') + ']}'
Write-Envelope ('{"ok":true,"data":' + $data + '}')
