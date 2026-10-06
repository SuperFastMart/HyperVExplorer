# Hypervisor Explorer installer / updater for Windows.
#
#   irm https://raw.githubusercontent.com/SuperFastMart/HyperVisorExplorer/main/install.ps1 | iex
#
# Files downloaded by PowerShell don't get the "downloaded from the internet" mark, so SmartScreen doesn't block the
# unsigned app. Installs per user (no admin): %LOCALAPPDATA%\Programs\HypervisorExplorer, a Start menu shortcut, and
# hvexplorer.exe on your user PATH. Re-run at any time to update. Options (environment variables):
#   $env:HVE_VERSION = 'v3.0.1'      install a specific release instead of the latest
#   $env:HVE_INSTALL_DIR = 'C:\...'   install somewhere else
#   $env:HVE_NO_LAUNCH = '1'          don't start the app afterwards
#   $env:HVE_WAIT_PID = '1234'        wait for this process to exit first (used by the app's own updater)
# Works in Windows PowerShell 5.1 and PowerShell 7.

& {
    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'   # the progress bar makes Invoke-WebRequest very slow in 5.1
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    $repo = 'SuperFastMart/HyperVisorExplorer'
    function Say([string]$Text) { Write-Host "==> $Text" -ForegroundColor Cyan }

    if ($env:OS -ne 'Windows_NT') { throw 'This installer is for Windows. On macOS use install.sh.' }

    # ------------------------------------------------------------ find the release asset
    $headers = @{ 'User-Agent' = 'HypervisorExplorer-Installer'; 'Accept' = 'application/vnd.github+json' }
    if ($env:HVE_VERSION) { $api = "https://api.github.com/repos/$repo/releases/tags/$($env:HVE_VERSION)" }
    else { $api = "https://api.github.com/repos/$repo/releases/latest" }
    Say 'Looking up the release...'
    $release = Invoke-RestMethod -Uri $api -Headers $headers -UseBasicParsing
    $asset = $release.assets | Where-Object { $_.name -like '*-win-x64.zip' } | Select-Object -First 1
    if (-not $asset) { throw "No Windows download found in release $($release.tag_name)." }

    $dest = $env:HVE_INSTALL_DIR
    if (-not $dest) { $dest = Join-Path $env:LOCALAPPDATA 'Programs\HypervisorExplorer' }

    # ------------------------------------------------------------ download and unpack
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ('hve-install-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    try {
        $zip = Join-Path $tmp $asset.name
        Say "Downloading Hypervisor Explorer $($release.tag_name)..."
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zip -Headers $headers -UseBasicParsing
        $unpacked = Join-Path $tmp 'x'
        Expand-Archive -Path $zip -DestinationPath $unpacked -Force
        if (-not (Test-Path (Join-Path $unpacked 'HypervisorExplorer.exe'))) { throw 'The download does not contain HypervisorExplorer.exe.' }

        # -------------------------------------------------------- close the running app
        if ($env:HVE_WAIT_PID) {
            Say 'Waiting for Hypervisor Explorer to close...'
            try { Wait-Process -Id ([int]$env:HVE_WAIT_PID) -Timeout 30 -ErrorAction SilentlyContinue } catch { }
        }
        $exePath = Join-Path $dest 'HypervisorExplorer.exe'
        $running = @(Get-Process -Name HypervisorExplorer -ErrorAction SilentlyContinue |
            Where-Object { $_.Path -and ([IO.Path]::GetFullPath($_.Path) -eq [IO.Path]::GetFullPath($exePath)) })
        if ($running.Count -gt 0) {
            Say 'Closing the running copy of Hypervisor Explorer...'
            foreach ($p in $running) { [void]$p.CloseMainWindow() }
            $running | Wait-Process -Timeout 10 -ErrorAction SilentlyContinue
            $running | Where-Object { -not $_.HasExited } | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 500
        }

        # -------------------------------------------------------- install
        Say "Installing to $dest"
        New-Item -ItemType Directory -Path $dest -Force | Out-Null
        Copy-Item -Path (Join-Path $unpacked '*') -Destination $dest -Recurse -Force
        Get-ChildItem -Path $dest -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

        # Start menu shortcut
        $programs = [Environment]::GetFolderPath('Programs')
        $shell = New-Object -ComObject WScript.Shell
        $lnk = $shell.CreateShortcut((Join-Path $programs 'Hypervisor Explorer.lnk'))
        $lnk.TargetPath = $exePath
        $lnk.WorkingDirectory = $dest
        $lnk.Description = 'Hypervisor Explorer - Hyper-V, Proxmox and VMware inventory'
        $lnk.Save()

        # hvexplorer.exe on the user PATH
        $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
        $parts = @($userPath -split ';' | Where-Object { $_ })
        if ($parts -notcontains $dest) {
            [Environment]::SetEnvironmentVariable('Path', (($parts + $dest) -join ';'), 'User')
            Say 'Added the install folder to your PATH (open a new terminal to use hvexplorer).'
        }

        Say "Hypervisor Explorer $($release.tag_name) is installed. Find it in the Start menu."
        if (-not $env:HVE_NO_LAUNCH) { Start-Process -FilePath $exePath -WorkingDirectory $dest }
    }
    finally {
        Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }
}
