#Requires -Version 5.1
# =============================================================================
#  PC BUILD TOOLKIT (v1.0.3)
#  - Self-elevating PowerShell + WPF
#  - Multi-source installers (winget / choco / direct / zip)
#  - Streaming download + tar.exe extraction (handles Deflate64/LZMA zips)
#
#  Launch locally:  powershell -ExecutionPolicy Bypass -File .\deployer.ps1
#  Launch from web: irm https://fay.digital/pbt | iex
# =============================================================================

$SCRIPT_VERSION = 'v1.0.3'
$SCRIPT_RAW_URL = 'https://raw.githubusercontent.com/fay-digital/pc-build-toolkit/main/deployer.ps1'

# --- Self-elevate if not admin -----------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
            ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    if ($PSCommandPath) {
        Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    }
    exit
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# --- App catalog -------------------------------------------------------------
$script:AppCatalog = @(
    @{ Id='FinalWire.AIDA64.Extreme';        Name='AIDA64 Extreme';  Category='Diagnostics'; Source='winget' }
    @{ Id='REALiX.HWiNFO';                   Name='HWiNFO';          Category='Diagnostics'; Source='winget' }
    @{ Id='CrystalDewWorld.CrystalDiskMark'; Name='CrystalDiskMark'; Category='Benchmark';   Source='winget' }
    @{ Id='Maxon.CinebenchR23';              Name='Cinebench R23';   Category='Benchmark';   Source='winget' }
    @{ Id='3dmark-bundled';                  Name='3DMark (Steel Nomad)'; Category='Benchmark'; Source='zip'
       DownloadUrl='https://github.com/fay-digital/pc-build-toolkit/releases/download/v1.0.0/3dmark-bundle.zip'
       SetupExecutable='3dmark-setup.exe'
       SilentArgs='/S'
       UninstallRegistryMatch='3DMark' }
    @{ Id='Geeks3D.FurMark.2'; Name='FurMark 2'; Category='GPU stress';  Source='winget' }
    @{ Id='OCCT.OCCT';         Name='OCCT';      Category='Stability';   Source='winget' }
    @{ Id='CPUID.CPU-Z';       Name='CPU-Z';     Category='Info';        Source='winget' }
    @{ Id='TechPowerUp.GPU-Z'; Name='GPU-Z';     Category='Info';        Source='winget' }
    @{ Id='GIMPS.Prime95';     Name='Prime95';   Category='CPU torture'; Source='winget' }
)

$script:DefaultChecked = @(
    'FinalWire.AIDA64.Extreme','REALiX.HWiNFO','CrystalDewWorld.CrystalDiskMark',
    'Maxon.CinebenchR23','3dmark-bundled'
)

# --- XAML UI -----------------------------------------------------------------
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PC Build Toolkit" Height="760" Width="960"
        WindowStartupLocation="CenterScreen" Background="#0F1115">
    <Window.Resources>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#E6E8EB"/>
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        <Style TargetType="TextBlock"><Setter Property="Foreground" Value="#E6E8EB"/></Style>
    </Window.Resources>
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="210"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0" Margin="0,0,0,14">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
                <TextBlock Text="PC BUILD TOOLKIT" FontSize="22" FontWeight="Bold"/>
                <TextBlock Text="Build validation deployer" Foreground="#7F8793" FontSize="12"/>
            </StackPanel>
            <TextBlock Grid.Column="1" Name="VersionLabel" Text="" Foreground="#7F8793"
                       FontSize="11" VerticalAlignment="Bottom"/>
        </Grid>

        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Background="#1A1D24" CornerRadius="6" Padding="18" Margin="0,0,7,0">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel>
                        <TextBlock Text="APPLICATIONS" Foreground="#7F8793" FontSize="11" FontWeight="Bold" Margin="0,0,0,10"/>
                        <StackPanel Name="AppPanel"/>
                        <Separator Margin="0,12,0,8" Background="#2A2F38"/>
                        <Button Name="BtnSelectAllApps" Content="Toggle all" Height="26"
                                Background="Transparent" Foreground="#7F8793" BorderThickness="0"
                                HorizontalAlignment="Left" Cursor="Hand"/>
                    </StackPanel>
                </ScrollViewer>
            </Border>

            <Border Grid.Column="1" Background="#1A1D24" CornerRadius="6" Padding="18" Margin="7,0,0,0">
                <StackPanel>
                    <TextBlock Text="SYSTEM TWEAKS" Foreground="#7F8793" FontSize="11" FontWeight="Bold" Margin="0,0,0,10"/>
                    <CheckBox Name="TweakPowerNever"       Content="Set power plan: display/sleep/hibernate -> never"/>
                    <CheckBox Name="TweakDisableHibernate" Content="Disable hibernation (powercfg -h off)"/>
                    <CheckBox Name="TweakClearDownloads"   Content="Clear Downloads folder"/>
                    <CheckBox Name="TweakEmptyRecycle"     Content="Empty Recycle Bin"/>
                    <CheckBox Name="TweakClearBrowser"     Content="Clear browser history (Edge, Chrome, Firefox)"/>
                    <Separator Margin="0,14,0,10" Background="#2A2F38"/>
                    <TextBlock Text="REPORT" Foreground="#7F8793" FontSize="11" FontWeight="Bold" Margin="0,0,0,6"/>
                    <CheckBox Name="OptBenchReport" Content="Generate bench report on Desktop after run" IsChecked="True"/>
                    <Separator Margin="0,14,0,10" Background="#2A2F38"/>
                    <TextBlock Text="DEBUG" Foreground="#7F8793" FontSize="11" FontWeight="Bold" Margin="0,0,0,6"/>
                    <CheckBox Name="OptKeepTemp" Content="Keep downloaded/extracted files for debugging"/>
                </StackPanel>
            </Border>
        </Grid>

        <Border Grid.Row="2" Background="#0A0C10" CornerRadius="6" Margin="0,14,0,0">
            <ScrollViewer Name="LogScroll" VerticalScrollBarVisibility="Auto">
                <TextBlock Name="LogOutput" Foreground="#B8C0CC" FontFamily="Consolas"
                           FontSize="12" Padding="14" TextWrapping="Wrap"/>
            </ScrollViewer>
        </Border>

        <Grid Grid.Row="3" Margin="0,14,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
                <TextBlock Name="StatusText" Text="Ready." Foreground="#B8C0CC" FontSize="12" Margin="0,0,0,4"/>
                <ProgressBar Name="ProgressBar" Height="8" Minimum="0" Maximum="100" Value="0"
                             Background="#1A1D24" Foreground="#3B82F6" BorderThickness="0"/>
            </StackPanel>
            <TextBlock Grid.Column="1" Name="ProgressPct" Text="" Foreground="#7F8793" FontSize="12"
                       Margin="12,0,0,0" VerticalAlignment="Center"/>
        </Grid>

        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
            <Button Name="BtnQuit"      Content="Quit"          Width="100" Height="34" Margin="0,0,10,0"
                    Background="#2A2F38" Foreground="#E6E8EB" BorderThickness="0" FontSize="13"/>
            <Button Name="BtnUninstall" Content="Uninstall all"  Width="140" Height="34" Margin="0,0,10,0"
                    Background="#7F1D1D" Foreground="White"     BorderThickness="0" FontSize="13"/>
            <Button Name="BtnRun"       Content="Run"            Width="130" Height="34"
                    Background="#3B82F6" Foreground="White"     BorderThickness="0"
                    FontWeight="Bold" FontSize="13"/>
        </StackPanel>
    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$controls = @{}
foreach ($n in 'AppPanel','LogOutput','LogScroll','BtnRun','BtnQuit','BtnUninstall','BtnSelectAllApps',
               'StatusText','ProgressBar','ProgressPct','VersionLabel',
               'TweakPowerNever','TweakDisableHibernate','TweakClearDownloads',
               'TweakEmptyRecycle','TweakClearBrowser','OptBenchReport','OptKeepTemp') {
    $controls[$n] = $window.FindName($n)
}
$controls.VersionLabel.Text = $SCRIPT_VERSION

# --- Synchronized state shared with the pipeline runspace --------------------
$sync = [hashtable]::Synchronized(@{})
$sync.Window         = $window
$sync.Log            = $controls.LogOutput
$sync.LogScroll      = $controls.LogScroll
$sync.Status         = $controls.StatusText
$sync.Progress       = $controls.ProgressBar
$sync.ProgressPct    = $controls.ProgressPct
$sync.BtnRun         = $controls.BtnRun
$sync.BtnUninst      = $controls.BtnUninstall
$sync.LogPath        = Join-Path $env:TEMP 'pcbt.log'
$sync.AppCatalog     = $script:AppCatalog
$sync.Mode           = $null
$sync.SelectedApps   = @()
$sync.SelectedTweaks = @{}
$sync.GenerateReport = $true
$sync.KeepTemp       = $false
$sync.RunResults     = @()

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $stamp = Get-Date -Format 'HH:mm:ss'
    $line  = "[$stamp] [$Level] $Message"
    Add-Content -Path $sync.LogPath -Value $line -ErrorAction SilentlyContinue
    $sync.Log.Dispatcher.Invoke([action]{
        $sync.Log.Text += ($line + "`n")
        $sync.LogScroll.ScrollToBottom()
    })
}

$appCheckboxes = @{}
foreach ($app in $script:AppCatalog) {
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.Content   = "$($app.Name)   —   $($app.Category)"
    $cb.IsChecked = ($script:DefaultChecked -contains $app.Id)
    $controls.AppPanel.AddChild($cb) | Out-Null
    $appCheckboxes[$app.Id] = $cb
}

function Invoke-PreflightChecks {
    param([int]$MinFreeGB = 10)
    Write-Log "Running pre-flight checks..."
    $problems = @()

    try {
        $ok = Test-Connection -ComputerName '1.1.1.1' -Count 1 -Quiet -ErrorAction Stop
        if ($ok) { Write-Log "Internet: reachable." 'OK' } else { $problems += "No internet connectivity." }
    } catch { $problems += "Internet check failed: $_" }

    try {
        $vol = Get-Volume -DriveLetter C -ErrorAction Stop
        $freeGB = [Math]::Round($vol.SizeRemaining / 1GB, 1)
        if ($freeGB -lt $MinFreeGB) { $problems += "Only ${freeGB} GB free on C: (need ${MinFreeGB}+ GB)." }
        else { Write-Log "Disk C: ${freeGB} GB free." 'OK' }
    } catch { Write-Log "Disk check skipped: $_" 'WARN' }

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        try {
            $null = & winget source list 2>&1
            if ($LASTEXITCODE -eq 0) { Write-Log "winget sources: healthy." 'OK' }
            else { Write-Log "winget sources returned exit $LASTEXITCODE." 'WARN' }
        } catch { Write-Log "winget source check failed: $_" 'WARN' }
    } else { $problems += "winget not found on PATH." }

    if ($problems.Count -gt 0) {
        foreach ($p in $problems) { Write-Log $p 'ERROR' }
        return $false
    }
    return $true
}

function Invoke-SelfUpdateCheck {
    if (-not $PSCommandPath -or -not (Test-Path $PSCommandPath)) {
        Write-Log "Loaded from web stream; skipping self-update check."; return
    }
    try {
        $local  = Get-FileHash -Path $PSCommandPath -Algorithm SHA256
        $remote = Invoke-WebRequest -UseBasicParsing -Uri $SCRIPT_RAW_URL -TimeoutSec 10
        $bytes  = [System.Text.Encoding]::UTF8.GetBytes($remote.Content)
        $sha    = [System.Security.Cryptography.SHA256]::Create()
        $rHash  = -join ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') })
        if ($local.Hash.ToLower() -ne $rHash.ToLower()) {
            Write-Log "A newer version is available at $SCRIPT_RAW_URL" 'INFO'
        } else { Write-Log "Running latest version ($SCRIPT_VERSION)." 'OK' }
    } catch { Write-Log "Self-update check skipped: $_" 'WARN' }
}

$pipelineCode = @'
function Write-UiLog {
    param([string]$Message, [string]$Level = 'INFO')
    $stamp = Get-Date -Format 'HH:mm:ss'
    $line  = "[$stamp] [$Level] $Message"
    Add-Content -Path $sync.LogPath -Value $line -ErrorAction SilentlyContinue
    $sync.Log.Dispatcher.Invoke([action]{
        $sync.Log.Text += ($line + "`n")
        $sync.LogScroll.ScrollToBottom()
    })
}
function Write-UiLogReplace {
    param([string]$Message)
    $stamp = Get-Date -Format 'HH:mm:ss'
    $line  = "[$stamp] [INFO] $Message"
    $sync.Log.Dispatcher.Invoke([action]{
        $text = $sync.Log.Text
        if ($text.EndsWith("`n")) { $text = $text.Substring(0, $text.Length - 1) }
        $lastNewline = $text.LastIndexOf("`n")
        if ($lastNewline -ge 0) { $text = $text.Substring(0, $lastNewline + 1) } else { $text = '' }
        $sync.Log.Text = $text + $line + "`n"
        $sync.LogScroll.ScrollToBottom()
    })
}
function Set-UiStatus   { param([string]$t) $sync.Status.Dispatcher.Invoke([action]{ $sync.Status.Text = $t }) }
function Set-UiProgress {
    param([double]$v)
    $sync.Progress.Dispatcher.Invoke([action]{
        $sync.Progress.Value = $v
        $sync.ProgressPct.Text = "{0:N0}%" -f $v
    })
}
function Set-UiBusy {
    param([bool]$busy)
    $sync.BtnRun.Dispatcher.Invoke([action]{
        $sync.BtnRun.IsEnabled    = -not $busy
        $sync.BtnUninst.IsEnabled = -not $busy
    })
}
function Add-Result { param($App, $Action, $Status, $Detail='')
    $sync.RunResults += [pscustomobject]@{ App=$App.Name; Action=$Action; Status=$Status; Detail=$Detail }
}

function Invoke-StreamingDownload {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$OutFile,
        [string]$AppName = 'file'
    )

    Add-Type -AssemblyName System.Net.Http
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072

    $handler = New-Object System.Net.Http.HttpClientHandler
    $handler.AllowAutoRedirect = $true
    $client = New-Object System.Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromMinutes(30)

    try {
        $response = $client.GetAsync($Url, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
        if (-not $response.IsSuccessStatusCode) { throw "HTTP $([int]$response.StatusCode) $($response.ReasonPhrase)" }
        $ct = $response.Content.Headers.ContentType
        if ($ct) {
            $mt = $ct.MediaType
            if ($mt -match '^(text/|application/json|application/xml)') {
                throw "Server returned $mt (likely an error page, not a file). Check the URL."
            }
        }
        $totalBytes = $response.Content.Headers.ContentLength
        $totalMB    = if ($totalBytes) { [Math]::Round($totalBytes / 1MB, 1) } else { $null }
        if ($totalBytes) { Write-UiLog "Downloading $AppName ($totalMB MB)..." }
        else             { Write-UiLog "Downloading $AppName (size unknown)..." }

        $sourceStream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        $destStream   = [System.IO.File]::Create($OutFile)
        $buffer       = New-Object byte[] (1MB)
        $totalRead    = 0L
        $lastReport   = [DateTime]::UtcNow
        $startTime    = [DateTime]::UtcNow
        $firstLine    = $true

        try {
            while ($true) {
                $read = $sourceStream.Read($buffer, 0, $buffer.Length)
                if ($read -le 0) { break }
                $destStream.Write($buffer, 0, $read)
                $totalRead += $read
                $now = [DateTime]::UtcNow
                if (($now - $lastReport).TotalMilliseconds -ge 1000) {
                    $elapsed = ($now - $startTime).TotalSeconds
                    $speedMB = if ($elapsed -gt 0) { [Math]::Round(($totalRead / 1MB) / $elapsed, 1) } else { 0 }
                    $doneMB  = [Math]::Round($totalRead / 1MB, 1)
                    if ($totalBytes) {
                        $pct = ($totalRead / $totalBytes) * 100
                        $msg = "Downloading ${AppName}: $doneMB / $totalMB MB ({0:N1}%) - $speedMB MB/s" -f $pct
                        Set-UiProgress $pct
                    } else {
                        $msg = "Downloading ${AppName}: $doneMB MB - $speedMB MB/s"
                    }
                    if ($firstLine) { Write-UiLog $msg; $firstLine = $false }
                    else            { Write-UiLogReplace $msg }
                    $lastReport = $now
                }
            }
        }
        finally {
            $destStream.Flush()
            $destStream.Close()
            $sourceStream.Close()
        }
        if ($totalBytes -and $totalRead -ne $totalBytes) {
            throw "Download truncated: got $totalRead bytes, expected $totalBytes."
        }
        $finalMB = [Math]::Round($totalRead / 1MB, 1)
        Write-UiLog "Download complete: $finalMB MB." 'OK'
    }
    finally {
        $client.Dispose()
        $handler.Dispose()
    }
}

function Test-IsZipFile {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $false }
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        $hdr = New-Object byte[] 4
        $n = $fs.Read($hdr, 0, 4)
        $fs.Close()
        return ($n -eq 4 -and $hdr[0] -eq 0x50 -and $hdr[1] -eq 0x4B -and $hdr[2] -eq 0x03 -and $hdr[3] -eq 0x04)
    } catch { return $false }
}

# Extract a zip. Tries tar.exe (ships with Win10 1803+, handles Deflate64,
# LZMA, ZIP64, etc) first; falls back to Expand-Archive only if tar is absent.
function Expand-ZipArchive {
    param(
        [Parameter(Mandatory)][string]$ZipPath,
        [Parameter(Mandatory)][string]$DestinationPath
    )
    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null

    $tar = Get-Command tar.exe -ErrorAction SilentlyContinue
    if ($tar) {
        Write-UiLog "Extracting with tar.exe..."
        $stderrFile = Join-Path $env:TEMP ("pbt_tar_err_" + [Guid]::NewGuid().ToString('N').Substring(0,8) + ".log")
        $p = Start-Process -FilePath $tar.Source `
            -ArgumentList @('-xf', "`"$ZipPath`"", '-C', "`"$DestinationPath`"") `
            -Wait -PassThru -NoNewWindow `
            -RedirectStandardError $stderrFile
        $err = if (Test-Path $stderrFile) { Get-Content $stderrFile -Raw -ErrorAction SilentlyContinue } else { $null }
        Remove-Item $stderrFile -Force -ErrorAction SilentlyContinue
        if ($p.ExitCode -eq 0) {
            Write-UiLog "Extraction OK (tar.exe)." 'OK'
            return
        }
        Write-UiLog "tar.exe failed (exit $($p.ExitCode)). $($err -replace '\s+', ' ')" 'WARN'
        Write-UiLog "Falling back to Expand-Archive..."
    } else {
        Write-UiLog "tar.exe not found. Using Expand-Archive..."
    }

    # Fallback - works for standard Deflate zips only
    Expand-Archive -Path $ZipPath -DestinationPath $DestinationPath -Force -ErrorAction Stop
    Write-UiLog "Extraction OK (Expand-Archive)." 'OK'
}

function Start-SilentInstaller {
    param(
        [Parameter(Mandatory)][string]$ExePath,
        [Parameter(Mandatory)][string]$Arguments,
        [Parameter(Mandatory)][string]$AppName
    )
    $workDir  = Split-Path -Parent $ExePath
    $stdoutFile = Join-Path $env:TEMP ("pbt_stdout_" + [Guid]::NewGuid().ToString('N').Substring(0,8) + ".log")
    $stderrFile = Join-Path $env:TEMP ("pbt_stderr_" + [Guid]::NewGuid().ToString('N').Substring(0,8) + ".log")

    Write-UiLog "Running $([IO.Path]::GetFileName($ExePath)) in '$workDir' with args '$Arguments'..."
    $p = Start-Process -FilePath $ExePath -ArgumentList $Arguments `
        -WorkingDirectory $workDir `
        -RedirectStandardOutput $stdoutFile `
        -RedirectStandardError  $stderrFile `
        -Wait -PassThru -NoNewWindow

    $stdout = if (Test-Path $stdoutFile) { Get-Content $stdoutFile -Raw -ErrorAction SilentlyContinue } else { $null }
    $stderr = if (Test-Path $stderrFile) { Get-Content $stderrFile -Raw -ErrorAction SilentlyContinue } else { $null }
    Remove-Item $stdoutFile, $stderrFile -Force -ErrorAction SilentlyContinue

    if ($stdout -and $stdout.Trim()) { Write-UiLog "$AppName stdout: $($stdout.Trim() -replace "`r?`n", ' | ')" 'INFO' }
    if ($stderr -and $stderr.Trim()) { Write-UiLog "$AppName stderr: $($stderr.Trim() -replace "`r?`n", ' | ')" 'WARN' }
    return $p.ExitCode
}

function Test-ChocoInstalled { [bool](Get-Command choco.exe -ErrorAction SilentlyContinue) }
function Install-Chocolatey {
    Write-UiLog "Chocolatey not found. Installing..."
    Set-UiStatus "Installing Chocolatey..."
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path","User")
    if (-not (Test-ChocoInstalled)) { throw "Chocolatey install did not expose choco.exe." }
    Write-UiLog "Chocolatey installed." 'OK'
}

function Invoke-WingetInstall { param($App)
    Write-UiLog "Installing $($App.Name) via winget ($($App.Id))..."
    $p = Start-Process winget -ArgumentList @('install','--id',$App.Id,'--silent',
        '--accept-package-agreements','--accept-source-agreements','--disable-interactivity') `
        -Wait -PassThru -NoNewWindow
    switch ($p.ExitCode) {
        0           { Write-UiLog "$($App.Name) installed." 'OK';                    Add-Result $App 'install' 'OK' }
        -1978335189 { Write-UiLog "$($App.Name) already installed." 'OK';             Add-Result $App 'install' 'OK' 'already installed' }
        -1978335212 { Write-UiLog "$($App.Name) - id not in winget." 'ERROR';         Add-Result $App 'install' 'FAIL' 'id not found' }
        default     { Write-UiLog "$($App.Name) exit $($p.ExitCode)." 'WARN';         Add-Result $App 'install' 'WARN' "exit $($p.ExitCode)" }
    }
}
function Invoke-WingetUninstall { param($App)
    Write-UiLog "Uninstalling $($App.Name) via winget..."
    $p = Start-Process winget -ArgumentList @('uninstall','--id',$App.Id,'--silent',
        '--accept-source-agreements','--disable-interactivity') -Wait -PassThru -NoNewWindow
    switch ($p.ExitCode) {
        0           { Write-UiLog "$($App.Name) uninstalled." 'OK';                   Add-Result $App 'uninstall' 'OK' }
        -1978335212 { Write-UiLog "$($App.Name) not installed, skipped." 'INFO';      Add-Result $App 'uninstall' 'SKIP' 'not present' }
        default     { Write-UiLog "$($App.Name) exit $($p.ExitCode)." 'WARN';         Add-Result $App 'uninstall' 'WARN' "exit $($p.ExitCode)" }
    }
}

function Invoke-ChocoInstall { param($App)
    if (-not (Test-ChocoInstalled)) { Install-Chocolatey }
    Write-UiLog "Installing $($App.Name) via Chocolatey..."
    $p = Start-Process choco -ArgumentList @('install',$App.Id,'-y','--no-progress','--limit-output') -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -eq 0) { Write-UiLog "$($App.Name) installed." 'OK'; Add-Result $App 'install' 'OK' }
    else { Write-UiLog "$($App.Name) exit $($p.ExitCode)." 'WARN'; Add-Result $App 'install' 'WARN' "exit $($p.ExitCode)" }
}
function Invoke-ChocoUninstall { param($App)
    if (-not (Test-ChocoInstalled)) { Write-UiLog "Chocolatey not present, skipping $($App.Name)." 'INFO'; Add-Result $App 'uninstall' 'SKIP'; return }
    Write-UiLog "Uninstalling $($App.Name) via Chocolatey..."
    $p = Start-Process choco -ArgumentList @('uninstall',$App.Id,'-y','--no-progress','--limit-output') -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -eq 0) { Write-UiLog "$($App.Name) uninstalled." 'OK'; Add-Result $App 'uninstall' 'OK' }
    else { Write-UiLog "$($App.Name) exit $($p.ExitCode) (may be absent)." 'INFO'; Add-Result $App 'uninstall' 'SKIP' "exit $($p.ExitCode)" }
}

function Invoke-DirectInstall { param($App)
    if (-not $App.DownloadUrl) { Write-UiLog "$($App.Name): no DownloadUrl." 'ERROR'; Add-Result $App 'install' 'FAIL' 'no URL'; return }
    $tmp = Join-Path $env:TEMP ("pbt_" + [IO.Path]::GetFileName($App.DownloadUrl))
    Set-UiStatus "Downloading $($App.Name)..."
    try {
        Invoke-StreamingDownload -Url $App.DownloadUrl -OutFile $tmp -AppName $App.Name
    } catch {
        Write-UiLog "$($App.Name) download failed: $_" 'ERROR'
        Add-Result $App 'install' 'FAIL' 'download failed'
        if (-not $sync.KeepTemp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        return
    }
    Set-UiStatus "Installing $($App.Name)..."
    $silent = if ($App.SilentArgs) { $App.SilentArgs } else { '/S' }
    $exit = Start-SilentInstaller -ExePath $tmp -Arguments $silent -AppName $App.Name
    if (-not $sync.KeepTemp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    else { Write-UiLog "Kept installer: $tmp" 'INFO' }
    if ($exit -eq 0) { Write-UiLog "$($App.Name) installed." 'OK'; Add-Result $App 'install' 'OK' }
    else { Write-UiLog "$($App.Name) installer exit $exit (0x$('{0:X8}' -f $exit))." 'WARN'; Add-Result $App 'install' 'WARN' "exit $exit" }
}

function Invoke-ZipInstall { param($App)
    if (-not $App.DownloadUrl)     { Write-UiLog "$($App.Name): no DownloadUrl." 'ERROR';     Add-Result $App 'install' 'FAIL' 'no URL';       return }
    if (-not $App.SetupExecutable) { Write-UiLog "$($App.Name): no SetupExecutable." 'ERROR'; Add-Result $App 'install' 'FAIL' 'no setup exe'; return }

    $tmpZip = Join-Path $env:TEMP ("pbt_" + [IO.Path]::GetFileName($App.DownloadUrl))
    $tmpDir = Join-Path $env:TEMP ("pbt_extract_" + [Guid]::NewGuid().ToString('N').Substring(0,8))

    try {
        Set-UiStatus "Downloading $($App.Name)..."
        Invoke-StreamingDownload -Url $App.DownloadUrl -OutFile $tmpZip -AppName $App.Name

        if (-not (Test-IsZipFile $tmpZip)) {
            throw "Downloaded file is not a valid zip (wrong magic bytes)."
        }

        Set-UiStatus "Extracting $($App.Name)..."
        Write-UiLog "Extracting $($App.Name) to $tmpDir ..."
        Expand-ZipArchive -ZipPath $tmpZip -DestinationPath $tmpDir

        $setup = Get-ChildItem -Path $tmpDir -Filter $App.SetupExecutable -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $setup) {
            throw "'$($App.SetupExecutable)' not found in extracted content at $tmpDir"
        }

        Set-UiStatus "Installing $($App.Name)..."
        $silent = if ($App.SilentArgs) { $App.SilentArgs } else { '/S' }
        $exit = Start-SilentInstaller -ExePath $setup.FullName -Arguments $silent -AppName $App.Name

        if ($exit -eq 0) { Write-UiLog "$($App.Name) installed." 'OK'; Add-Result $App 'install' 'OK' }
        else {
            $hex = '0x{0:X8}' -f $exit
            Write-UiLog "$($App.Name) installer exit $exit ($hex). Setup ran from '$($setup.DirectoryName)'." 'WARN'
            Add-Result $App 'install' 'WARN' "exit $exit ($hex)"
        }
    }
    catch {
        Write-UiLog "$($App.Name) install error: $_" 'ERROR'
        Add-Result $App 'install' 'FAIL' $_.Exception.Message
    }
    finally {
        if ($sync.KeepTemp) {
            Write-UiLog "Kept zip:     $tmpZip" 'INFO'
            Write-UiLog "Kept extract: $tmpDir" 'INFO'
        } else {
            Remove-Item $tmpZip -Force -ErrorAction SilentlyContinue
            Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-RegistryUninstall { param($App)
    $match = $App.UninstallRegistryMatch
    if (-not $match) { Write-UiLog "$($App.Name): no UninstallRegistryMatch." 'WARN'; return }
    $keys = @('HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
              'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')
    $found = $false
    foreach ($k in $keys) {
        if (-not (Test-Path $k)) { continue }
        Get-ChildItem $k -ErrorAction SilentlyContinue | ForEach-Object {
            $dn = $_.GetValue('DisplayName')
            if ($dn -and $dn -like "*$match*") {
                $us = $_.GetValue('QuietUninstallString')
                if (-not $us) { $us = $_.GetValue('UninstallString') }
                if ($us) {
                    Write-UiLog "Uninstalling '$dn'..."
                    Start-Process -FilePath cmd.exe -ArgumentList "/c",($us + " /S") -Wait -NoNewWindow
                    Write-UiLog "$dn uninstalled." 'OK'
                    Add-Result $App 'uninstall' 'OK' $dn
                    $found = $true
                }
            }
        }
    }
    if (-not $found) { Write-UiLog "$($App.Name) not found in registry, skipped." 'INFO'; Add-Result $App 'uninstall' 'SKIP' 'not present' }
}

function Invoke-TweakPowerNever {
    Write-UiLog "Setting power timeouts to never..."
    powercfg -change -monitor-timeout-ac 0
    powercfg -change -monitor-timeout-dc 0
    powercfg -change -standby-timeout-ac 0
    powercfg -change -standby-timeout-dc 0
    powercfg -change -hibernate-timeout-ac 0
    powercfg -change -hibernate-timeout-dc 0
    Write-UiLog "Power timeouts set." 'OK'
}
function Invoke-TweakDisableHibernate { Write-UiLog "Disabling hibernation..."; powercfg -h off; Write-UiLog "Done." 'OK' }
function Invoke-TweakClearDownloads {
    $dl = Join-Path $env:USERPROFILE 'Downloads'
    if (Test-Path $dl) {
        Get-ChildItem $dl -Recurse -Force -EA SilentlyContinue | Remove-Item -Recurse -Force -EA SilentlyContinue
        Write-UiLog "Downloads cleared." 'OK'
    } else { Write-UiLog "Downloads folder not found." 'WARN' }
}
function Invoke-TweakEmptyRecycle {
    try { Clear-RecycleBin -Force -EA Stop; Write-UiLog "Recycle Bin emptied." 'OK' }
    catch { Write-UiLog "Recycle Bin: $_" 'WARN' }
}
function Invoke-TweakClearBrowser {
    Write-UiLog "Clearing browser history (close browsers first)..."
    $t = @(
        @{ N='Edge';   P="$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\History" }
        @{ N='Chrome'; P="$env:LOCALAPPDATA\Google\Chrome\User Data\Default\History"  }
    )
    foreach ($i in $t) {
        if (Test-Path $i.P) {
            try { Remove-Item $i.P -Force -EA Stop; Write-UiLog "Cleared $($i.N)." 'OK' }
            catch { Write-UiLog "$($i.N) locked." 'WARN' }
        }
    }
    $ff = Join-Path $env:APPDATA 'Mozilla\Firefox\Profiles'
    if (Test-Path $ff) {
        Get-ChildItem $ff -Directory | ForEach-Object {
            $p = Join-Path $_.FullName 'places.sqlite'
            if (Test-Path $p) {
                try { Remove-Item $p -Force -EA Stop; Write-UiLog "Cleared Firefox ($($_.Name))." 'OK' }
                catch { Write-UiLog "Firefox profile $($_.Name) locked." 'WARN' }
            }
        }
    }
}

function New-BenchReport {
    try {
        $cpu  = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name
        $gpu  = ((Get-CimInstance Win32_VideoController) | Where-Object { $_.Name -notmatch 'Virtual|Basic' } | ForEach-Object Name) -join ', '
        $ram  = [Math]::Round(((Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum / 1GB), 0)
        $mb   = Get-CimInstance Win32_BaseBoard | Select-Object Manufacturer, Product
        $disks = (Get-CimInstance Win32_DiskDrive | ForEach-Object { "$($_.Model) ($([Math]::Round($_.Size/1GB,0)) GB)" }) -join "`n             "

        $sb = New-Object Text.StringBuilder
        [void]$sb.AppendLine("PC Build Validation Report")
        [void]$sb.AppendLine("Generated:   $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
        [void]$sb.AppendLine("Hostname:    $env:COMPUTERNAME")
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("SYSTEM")
        [void]$sb.AppendLine("  CPU:         $cpu")
        [void]$sb.AppendLine("  GPU:         $gpu")
        [void]$sb.AppendLine("  RAM:         ${ram} GB")
        [void]$sb.AppendLine("  Motherboard: $($mb.Manufacturer) $($mb.Product)")
        [void]$sb.AppendLine("  Storage:     $disks")
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("ACTIONS ($($sync.Mode.ToUpper()))")
        foreach ($r in $sync.RunResults) {
            [void]$sb.AppendLine(("  [{0,-4}] {1,-22} {2,-10} {3}" -f $r.Status, $r.App, $r.Action, $r.Detail))
        }
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("Full log: $($sync.LogPath)")

        $desk = [Environment]::GetFolderPath('Desktop')
        $fn   = "PCBT_Report_{0}_{1}.txt" -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd_HHmm')
        $path = Join-Path $desk $fn
        Set-Content -Path $path -Value $sb.ToString() -Encoding UTF8
        Write-UiLog "Bench report saved: $path" 'OK'
    } catch { Write-UiLog "Report generation failed: $_" 'WARN' }
}

try {
    Set-UiBusy $true
    Set-UiProgress 0
    $sync.RunResults = @()
    $apps   = @($sync.SelectedApps)
    $tweaks = $sync.SelectedTweaks
    $mode   = $sync.Mode
    Write-UiLog "Pipeline dispatching ($mode, $($apps.Count) apps)..."
    if ($sync.KeepTemp) { Write-UiLog "Debug mode: temp files will be kept." 'INFO' }

    if ($mode -eq 'Install') {
        Write-UiLog "=============== INSTALL RUN STARTED ==============="
        $tweakCount = (($tweaks.Values | Where-Object { $_ }) | Measure-Object).Count
        $total = [Math]::Max(1, $apps.Count + $tweakCount)
        $done  = 0

        if (($apps | Where-Object { $_.Source -eq 'choco' }).Count -gt 0 -and -not (Test-ChocoInstalled)) {
            Install-Chocolatey
        }
        $i = 0
        foreach ($app in $apps) {
            $i++
            Set-UiStatus "Installing $($app.Name)  ($i of $($apps.Count))..."
            try {
                switch ($app.Source) {
                    'winget' { Invoke-WingetInstall $app }
                    'choco'  { Invoke-ChocoInstall  $app }
                    'direct' { Invoke-DirectInstall $app }
                    'zip'    { Invoke-ZipInstall    $app }
                    default  { Write-UiLog "Unknown source '$($app.Source)'." 'ERROR' }
                }
            } catch { Write-UiLog "Error on $($app.Name): $_" 'ERROR' }
            $done++; Set-UiProgress (($done / $total) * 100)
        }

        $tweakList = @(
            @{K='PowerNever';       L='Set power plan';    A={ Invoke-TweakPowerNever }}
            @{K='DisableHibernate'; L='Disable hibernate'; A={ Invoke-TweakDisableHibernate }}
            @{K='ClearDownloads';   L='Clear Downloads';   A={ Invoke-TweakClearDownloads }}
            @{K='EmptyRecycle';     L='Empty Recycle Bin'; A={ Invoke-TweakEmptyRecycle }}
            @{K='ClearBrowser';     L='Clear browser';     A={ Invoke-TweakClearBrowser }}
        )
        foreach ($t in $tweakList) {
            if ($tweaks[$t.K]) {
                Set-UiStatus "$($t.L)..."
                try { & $t.A } catch { Write-UiLog "Tweak $($t.L): $_" 'ERROR' }
                $done++; Set-UiProgress (($done / $total) * 100)
            }
        }

        Set-UiStatus "Install complete."
        Set-UiProgress 100
        Write-UiLog "=============== INSTALL RUN COMPLETE =============="
        if ($sync.GenerateReport) { New-BenchReport }
    }
    elseif ($mode -eq 'Uninstall') {
        Write-UiLog "=============== UNINSTALL RUN STARTED ============="
        $total = [Math]::Max(1, $apps.Count)
        $done = 0; $i = 0
        foreach ($app in $apps) {
            $i++
            Set-UiStatus "Uninstalling $($app.Name)  ($i of $($apps.Count))..."
            try {
                switch ($app.Source) {
                    'winget'         { Invoke-WingetUninstall   $app }
                    'choco'          { Invoke-ChocoUninstall    $app }
                    { $_ -in 'direct','zip' } { Invoke-RegistryUninstall $app }
                }
            } catch { Write-UiLog "Error on $($app.Name): $_" 'ERROR' }
            $done++; Set-UiProgress (($done / $total) * 100)
        }
        Set-UiStatus "Uninstall complete."
        Set-UiProgress 100
        Write-UiLog "=============== UNINSTALL RUN COMPLETE ============"
        if ($sync.GenerateReport) { New-BenchReport }
    }
}
catch { Write-UiLog "Pipeline fatal: $_" 'ERROR' }
finally { Set-UiBusy $false }
'@

function Start-Pipeline {
    param([string]$Mode, [array]$SelectedApps, [hashtable]$SelectedTweaks, [bool]$GenerateReport, [bool]$KeepTemp)
    $sync.Mode           = $Mode
    $sync.SelectedApps   = $SelectedApps
    $sync.SelectedTweaks = $SelectedTweaks
    $sync.GenerateReport = $GenerateReport
    $sync.KeepTemp       = $KeepTemp
    Write-Log "Starting $Mode pipeline..."
    $rs = [RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('sync', $sync)
    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    $null = $ps.AddScript($pipelineCode)
    $null = $ps.BeginInvoke()
}

$controls.BtnQuit.Add_Click({ $window.Close() })
$controls.BtnSelectAllApps.Add_Click({
    $anyUnchecked = $appCheckboxes.Values | Where-Object { -not $_.IsChecked }
    $newState = [bool]$anyUnchecked
    $appCheckboxes.Values | ForEach-Object { $_.IsChecked = $newState }
})
$controls.BtnRun.Add_Click({
    $selectedApps = @()
    foreach ($app in $script:AppCatalog) {
        if ($appCheckboxes[$app.Id].IsChecked) { $selectedApps += $app }
    }
    $selectedTweaks = @{
        PowerNever       = [bool]$controls.TweakPowerNever.IsChecked
        DisableHibernate = [bool]$controls.TweakDisableHibernate.IsChecked
        ClearDownloads   = [bool]$controls.TweakClearDownloads.IsChecked
        EmptyRecycle     = [bool]$controls.TweakEmptyRecycle.IsChecked
        ClearBrowser     = [bool]$controls.TweakClearBrowser.IsChecked
    }
    if ($selectedApps.Count -eq 0 -and (($selectedTweaks.Values | Where-Object { $_ }).Count -eq 0)) {
        [System.Windows.MessageBox]::Show('Select at least one app or tweak.','Nothing to do','OK','Information') | Out-Null
        return
    }
    if (-not (Invoke-PreflightChecks)) {
        [System.Windows.MessageBox]::Show('Pre-flight checks failed. See log.','Cannot run','OK','Error') | Out-Null
        return
    }
    Start-Pipeline -Mode 'Install' -SelectedApps $selectedApps -SelectedTweaks $selectedTweaks `
                   -GenerateReport ([bool]$controls.OptBenchReport.IsChecked) `
                   -KeepTemp       ([bool]$controls.OptKeepTemp.IsChecked)
})
$controls.BtnUninstall.Add_Click({
    $r = [System.Windows.MessageBox]::Show(
        "Uninstall every app in the catalog that is currently installed?`n`nApps not present will be skipped.",
        'Confirm uninstall', 'YesNo', 'Warning')
    if ($r -ne [System.Windows.MessageBoxResult]::Yes) { return }
    Start-Pipeline -Mode 'Uninstall' -SelectedApps $script:AppCatalog -SelectedTweaks @{} `
                   -GenerateReport ([bool]$controls.OptBenchReport.IsChecked) `
                   -KeepTemp       ([bool]$controls.OptKeepTemp.IsChecked)
})

Write-Log "PC Build Toolkit $SCRIPT_VERSION ready. Log: $($sync.LogPath)"
Invoke-SelfUpdateCheck

$window.ShowDialog() | Out-Null
