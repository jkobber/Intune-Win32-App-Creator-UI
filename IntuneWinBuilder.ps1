# IntuneWinBuilder.ps1
# Author: https://github.com/jkobber
# Production-ready, single-file GUI wizard for building .intunewin packages.

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$script:UiRunspace = [System.Management.Automation.Runspaces.Runspace]::DefaultRunspace
$script:AllowClose = $false
$script:LogFile = Join-Path $env:TEMP "IntuneWinBuilder.log"
try {
    "=== IntuneWinBuilder session $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" | Out-File -FilePath $script:LogFile -Encoding UTF8 -Append
} catch { }

# WPF requires STA; relaunch in Windows PowerShell if needed.
$scriptPath = $MyInvocation.MyCommand.Path
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Warning "WPF erfordert STA. Starte das Skript erneut in Windows PowerShell (STA)."
    if ($scriptPath) {
        $psExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
        Start-Process -FilePath $psExe -ArgumentList @("-STA", "-ExecutionPolicy", "Bypass", "-File", "`"$scriptPath`"") | Out-Null
    }
    return
}

# VSCode Host: launch external STA console to ensure WPF window appears reliably.
if (-not $env:INTUNEWINBUILDER_CHILD -and $Host.Name -like "*Visual Studio Code Host*") {
    Write-Warning "VSCode Host erkannt. Starte in externem Windows PowerShell (STA)."
    if ($scriptPath) {
        $env:INTUNEWINBUILDER_CHILD = "1"
        $psExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
        Start-Process -FilePath $psExe -ArgumentList @("-STA", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$scriptPath`"") | Out-Null
    }
    return
}

# Paths
$script:WorkRoot  = "C:\temp\IntuneWinBuilder"
$script:RepoRoot  = Join-Path $script:WorkRoot "Microsoft-Win32-Content-Prep-Tool-master"
$script:SetupRoot = Join-Path $script:WorkRoot "Setup-folder"
$script:OutputRoot = Join-Path $script:WorkRoot "output"
$script:ZipUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/archive/refs/heads/master.zip"
$script:ZipPath = Join-Path $script:WorkRoot "Microsoft-Win32-Content-Prep-Tool-master.zip"

# State
$script:EntryPointPath = $null
$script:IsBusy = $false
$script:CurrentProcess = $null
$script:LastError = $null
$script:SetupTimer = $null
$script:HasOutput = $false

# Load WPF assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

# Known folder helper (Downloads)
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace Win32 {
    public static class KnownFolder {
        [DllImport("shell32.dll")]
        public static extern int SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr pszPath);
    }
}
'@

function Get-DownloadsFolder {
    $guid = New-Object Guid "374DE290-123F-4565-9164-39C4925E467B"
    $ptr = [IntPtr]::Zero
    $hr = [Win32.KnownFolder]::SHGetKnownFolderPath($guid, 0, [IntPtr]::Zero, [ref]$ptr)
    if ($hr -ne 0 -or $ptr -eq [IntPtr]::Zero) {
        return (Join-Path $env:USERPROFILE "Downloads")
    }
    $path = [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)
    [Runtime.InteropServices.Marshal]::FreeCoTaskMem($ptr)
    return $path
}

function Invoke-Ui {
    param([scriptblock]$Action)
    $null = $script:Window.Dispatcher.Invoke($Action)
}

function Invoke-UiDoEvents {
    $frame = New-Object System.Windows.Threading.DispatcherFrame
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
        [Action] { $frame.Continue = $false },
        [System.Windows.Threading.DispatcherPriority]::Background
    ) | Out-Null
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
}

function Bring-WindowToFront {
    try {
        $script:Window.ShowInTaskbar = $true
        $script:Window.WindowState = [System.Windows.WindowState]::Normal
        $script:Window.Topmost = $true
        $script:Window.Activate() | Out-Null
        $script:Window.Topmost = $false
    } catch { }
}

function Bring-WindowToFrontTemporarily {
    param([int]$Seconds = 3)
    try {
        $script:Window.ShowInTaskbar = $true
        $script:Window.WindowState = [System.Windows.WindowState]::Normal
        $script:Window.Topmost = $true
        $script:Window.Activate() | Out-Null
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromSeconds($Seconds)
        $timer.Add_Tick({
            param($s, $e)
            $script:Window.Topmost = $false
            $s.Stop()
        })
        $timer.Start()
    } catch { }
}

function Init-ButtonDefaults {
    $script:ButtonDefaults = @{}
    foreach ($b in @($script:BtnStart,$script:BtnWeiter,$script:BtnEntry,$script:BtnBuild,$script:BtnCancel,$script:BtnClose)) {
        if ($null -ne $b -and -not $script:ButtonDefaults.ContainsKey($b.Name)) {
            $script:ButtonDefaults[$b.Name] = @{
                Background = $b.Background
                Foreground = $b.Foreground
                BorderBrush = $b.BorderBrush
                FontWeight = $b.FontWeight
            }
        }
    }
}

function Clear-NextActionHighlight {
    foreach ($b in @($script:BtnStart,$script:BtnWeiter,$script:BtnEntry,$script:BtnBuild,$script:BtnCancel,$script:BtnClose)) {
        if ($null -eq $b) { continue }
        $def = $script:ButtonDefaults[$b.Name]
        if ($def) {
            $b.Background = $def.Background
            $b.Foreground = $def.Foreground
            $b.BorderBrush = $def.BorderBrush
            $b.FontWeight = $def.FontWeight
        }
    }
}

function Set-NextActionButton {
    param([System.Windows.Controls.Button]$Button)
    if ($null -eq $Button) { return }
    Clear-NextActionHighlight
    $Button.Background = [System.Windows.Media.Brushes]::DarkSeaGreen
    $Button.Foreground = [System.Windows.Media.Brushes]::Black
    $Button.BorderBrush = [System.Windows.Media.Brushes]::SeaGreen
    $Button.FontWeight = [System.Windows.FontWeights]::Bold
}

function Add-Log {
    param([string]$Message)
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Invoke-Ui {
        $script:LogBox.AppendText("[$ts] $Message`r`n")
        $script:LogBox.ScrollToEnd()
    }
    try {
        "[$ts] $Message" | Out-File -FilePath $script:LogFile -Encoding UTF8 -Append
    } catch { }
}

function Set-Status {
    param([string]$Message)
    Invoke-Ui { $script:StatusText.Text = $Message }
}

function Set-Progress {
    param([bool]$Indeterminate)
    Invoke-Ui {
        $script:ProgressBar.IsIndeterminate = $Indeterminate
        if (-not $Indeterminate) {
            $script:ProgressBar.Minimum = 0
            $script:ProgressBar.Maximum = 100
            $script:ProgressBar.Value = 0
        }
    }
}

function Set-ProgressValue {
    param([int]$Value)
    Invoke-Ui {
        if ($Value -lt 0) { $Value = 0 }
        if ($Value -gt 100) { $Value = 100 }
        $script:ProgressBar.IsIndeterminate = $false
        $script:ProgressBar.Value = $Value
    }
}

function Join-ProcessArguments {
    param([string[]]$Arguments)
    $escaped = foreach ($a in $Arguments) {
        if ($a -match '[\s"]') {
            '"' + ($a -replace '"', '\"') + '"'
        } else {
            $a
        }
    }
    return ($escaped -join ' ')
}

function Wait-ProcessWithUi {
    param([Parameter(Mandatory=$true)][System.Diagnostics.Process]$Process)
    while (-not $Process.HasExited) {
        Start-Sleep -Milliseconds 150
        Invoke-UiDoEvents
    }
}

function Read-NewLinesFromFile {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][ref]$Position
    )
    if (-not (Test-Path $Path)) { return @() }
    $fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $fs.Seek($Position.Value, [System.IO.SeekOrigin]::Begin) | Out-Null
        $sr = New-Object System.IO.StreamReader($fs)
        $text = $sr.ReadToEnd()
        $Position.Value = $fs.Position
    } finally {
        try { $sr.Close() } catch { }
        try { $fs.Close() } catch { }
    }
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    return ($text -split "(`r`n|`n|`r)") | Where-Object { $_ -ne "" }
}

function Run-ProcessWithRedirectAndTail {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [Parameter(Mandatory=$true)][string[]]$Arguments,
        [Parameter(Mandatory=$true)][string]$WorkingDirectory,
        [Parameter(Mandatory=$true)][string]$StdOutPath,
        [Parameter(Mandatory=$true)][string]$StdErrPath,
        [string]$Label = "Prozess"
    )
    if (Test-Path $StdOutPath) { Remove-Item $StdOutPath -Force -ErrorAction SilentlyContinue }
    if (Test-Path $StdErrPath) { Remove-Item $StdErrPath -Force -ErrorAction SilentlyContinue }

    $argString = Join-ProcessArguments -Arguments $Arguments
    $proc = Start-Process -FilePath $FilePath -ArgumentList $argString -WorkingDirectory $WorkingDirectory `
        -NoNewWindow -PassThru -RedirectStandardOutput $StdOutPath -RedirectStandardError $StdErrPath

    $script:CurrentProcess = $proc
    $outPos = 0L
    $errPos = 0L
    while (-not $proc.HasExited) {
        foreach ($line in (Read-NewLinesFromFile -Path $StdOutPath -Position ([ref]$outPos))) {
            Add-Log "${Label}: $line"
        }
        foreach ($line in (Read-NewLinesFromFile -Path $StdErrPath -Position ([ref]$errPos))) {
            Add-Log "${Label} ERR: $line"
        }
        Start-Sleep -Milliseconds 150
        Invoke-UiDoEvents
    }
    foreach ($line in (Read-NewLinesFromFile -Path $StdOutPath -Position ([ref]$outPos))) {
        Add-Log "${Label}: $line"
    }
    foreach ($line in (Read-NewLinesFromFile -Path $StdErrPath -Position ([ref]$errPos))) {
        Add-Log "${Label} ERR: $line"
    }
    $exitCode = 0
    try { $exitCode = [int]$proc.ExitCode } catch { $exitCode = -1 }
    $script:CurrentProcess = $null
    return $exitCode
}

function Set-StepState {
    param([string]$StepName, [string]$State)
    Invoke-Ui {
        switch ($StepName) {
            "1" { $script:Step1Status.Text = $State }
            "2" { $script:Step2Status.Text = $State }
            "3" { $script:Step3Status.Text = $State }
            "4" { $script:Step4Status.Text = $State }
            "5" { $script:Step5Status.Text = $State }
        }
    }
}

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Remove-PathSafe {
    param([string]$Path)
    $full = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetFullPath($script:WorkRoot)
    if ($full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        if (Test-Path $full) {
            Remove-Item $full -Recurse -Force -ErrorAction Stop
        }
    } else {
        throw "Refusing to delete path outside WorkRoot: $full"
    }
}

function Ensure-Tls12 {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    } catch {
        # If TLS12 is not available, let Invoke-WebRequest handle it.
    }
}

function Download-FileWithProgress {
    param(
        [Parameter(Mandatory=$true)][string]$Url,
        [Parameter(Mandatory=$true)][string]$Destination
    )
    Ensure-Tls12
    Add-Log "Download URL: $Url"
    Set-Status "Download laeuft..."
    Set-Progress $false

    Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue
    $handler = New-Object System.Net.Http.HttpClientHandler
    $client = New-Object System.Net.Http.HttpClient($handler)
    try {
        $response = $client.GetAsync($Url, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).Result
        if (-not $response.IsSuccessStatusCode) {
            throw "Download fehlgeschlagen. HTTP $($response.StatusCode)"
        }

        $total = $response.Content.Headers.ContentLength
        if (-not $total -or $total -le 0) {
            Set-Progress $true
        } else {
            Set-Progress $false
        }

        $stream = $response.Content.ReadAsStreamAsync().Result
        $fs = New-Object System.IO.FileStream($Destination, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $buffer = New-Object byte[] 65536
            $read = 0
            $received = 0L
            $lastUpdate = [DateTime]::UtcNow
            while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $fs.Write($buffer, 0, $read)
                $received += $read
                if ($total -and $total -gt 0) {
                    $now = [DateTime]::UtcNow
                    if (($now - $lastUpdate).TotalMilliseconds -ge 200) {
                        $pct = [int]([math]::Floor(($received / $total) * 100))
                        Set-ProgressValue $pct
                        $lastUpdate = $now
                        Invoke-UiDoEvents
                    }
                }
            }
            if ($total -and $total -gt 0) {
                Set-ProgressValue 100
            }
        } finally {
            $fs.Close()
            $stream.Close()
        }
    } finally {
        $client.Dispose()
    }
}

function Validate-EntryPointPath {
    param([string]$Path)
    if (-not (Test-Path $Path -PathType Leaf)) { return $false }
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -notin @(".exe", ".ps1", ".msi")) { return $false }
    $full = [IO.Path]::GetFullPath($Path)
    $setup = [IO.Path]::GetFullPath($script:SetupRoot)
    if (-not $full.StartsWith($setup, [StringComparison]::OrdinalIgnoreCase)) { return $false }
    return $true
}

function Update-SetupFolderState {
    $hasFiles = $false
    try {
        $file = Get-ChildItem -Path $script:SetupRoot -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        $hasFiles = $null -ne $file
    } catch {
        $hasFiles = $false
    }
    Invoke-Ui {
        $script:BtnWeiter.IsEnabled = $hasFiles -and -not $script:IsBusy
    }
}

function Start-SetupFolderWatcher {
    if ($script:SetupTimer -ne $null) { return }
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds(1)
    $timer.Add_Tick({ Update-SetupFolderState })
    $timer.Start()
    $script:SetupTimer = $timer
}

function Stop-SetupFolderWatcher {
    if ($script:SetupTimer -ne $null) {
        $script:SetupTimer.Stop()
        $script:SetupTimer = $null
    }
}

function Download-Repo {
    Set-Status "Repository wird geladen..."
    Add-Log "Repo-Quelle: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool"
    Ensure-Directory $script:WorkRoot

    if (Test-Path $script:RepoRoot) {
        Add-Log "Vorhandenes Repo wird entfernt..."
        Remove-PathSafe $script:RepoRoot
    }

    $git = Get-Command git -ErrorAction SilentlyContinue
    if ($null -ne $git) {
        Add-Log "git gefunden. Klone Repo nach $script:RepoRoot"
        Set-Status "Download laeuft (git)..."
        Set-Progress $true
        $args = @("clone", "--depth", "1", "--progress", "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool.git", $script:RepoRoot)
        $gitOut = Join-Path $script:WorkRoot "git_stdout.log"
        $gitErr = Join-Path $script:WorkRoot "git_stderr.log"
        $exitCode = Run-ProcessWithRedirectAndTail -FilePath $git.Path -Arguments $args -WorkingDirectory $script:WorkRoot `
            -StdOutPath $gitOut -StdErrPath $gitErr -Label "git"
        $gitExit = [int]0
        try { $gitExit = [int]$exitCode } catch { $gitExit = -1 }
        if ($gitExit -ne 0) {
            Add-Log "git clone fehlgeschlagen (ExitCode: $exitCode). Fallback auf ZIP."
            Remove-PathSafe $script:RepoRoot
            Download-FileWithProgress -Url $script:ZipUrl -Destination $script:ZipPath
            $hash = (Get-FileHash -Path $script:ZipPath -Algorithm SHA256).Hash
            Add-Log "ZIP SHA256: $hash"
            Set-Status "Entpacken..."
            Expand-Archive -Path $script:ZipPath -DestinationPath $script:WorkRoot -Force
            Remove-Item $script:ZipPath -Force -ErrorAction SilentlyContinue
        }
    } else {
        Add-Log "git nicht gefunden. ZIP-Download via HTTPS: $script:ZipUrl"
        Download-FileWithProgress -Url $script:ZipUrl -Destination $script:ZipPath
        $hash = (Get-FileHash -Path $script:ZipPath -Algorithm SHA256).Hash
        Add-Log "ZIP SHA256: $hash"
        Set-Status "Entpacken..."
        Expand-Archive -Path $script:ZipPath -DestinationPath $script:WorkRoot -Force
        Remove-Item $script:ZipPath -Force -ErrorAction SilentlyContinue
    }

    $exe = Join-Path $script:RepoRoot "IntuneWinAppUtil.exe"
    if (-not (Test-Path $exe -PathType Leaf)) {
        throw "IntuneWinAppUtil.exe nicht gefunden unter $exe"
    }
    $exeFull = [IO.Path]::GetFullPath($exe)
    $repoFull = [IO.Path]::GetFullPath($script:RepoRoot)
    if (-not $exeFull.StartsWith($repoFull, [StringComparison]::OrdinalIgnoreCase)) {
        throw "IntuneWinAppUtil.exe Pfadpruefung fehlgeschlagen."
    }

    Set-Status "Fertig: Tool bereit."
}

function Ensure-Folders {
    Ensure-Directory $script:WorkRoot
    Ensure-Directory $script:RepoRoot
    Ensure-Directory $script:SetupRoot
    Ensure-Directory $script:OutputRoot
    Add-Log "Ordner Setup-folder und output erstellt/vorhanden."
    Set-Status "Ordner Setup-folder und output erstellt/vorhanden."
}

function Run-IntuneWin {
    $exe = Join-Path $script:RepoRoot "IntuneWinAppUtil.exe"
    $entry = $script:EntryPointPath
    $args = @("-s", $entry, "-c", $script:SetupRoot, "-o", $script:OutputRoot, "-q")

    Add-Log "Starte IntuneWinAppUtil.exe"
    Add-Log "Command: $exe $($args -join ' ')"
    $stdout = Join-Path $script:WorkRoot "intunewin_stdout.log"
    $stderr = Join-Path $script:WorkRoot "intunewin_stderr.log"
    $exitCode = Run-ProcessWithRedirectAndTail -FilePath $exe -Arguments $args -WorkingDirectory $script:WorkRoot `
        -StdOutPath $stdout -StdErrPath $stderr -Label "IntuneWinAppUtil"

    if ($exitCode -ne 0) {
        throw "IntuneWinAppUtil.exe exit code $exitCode"
    }
}


function Find-IntuneWinFile {
    $file = Get-ChildItem -Path $script:OutputRoot -Filter "*.intunewin" -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    return $file
}

function Export-ToDownloads {
    $downloads = Get-DownloadsFolder
    $stamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $destRoot = Join-Path $downloads "IntuneWinBuilder_$stamp"

    Ensure-Directory $destRoot
    $destSetup = Join-Path $destRoot "Setup-folder"
    $destOutput = Join-Path $destRoot "output"
    Ensure-Directory $destSetup
    Ensure-Directory $destOutput

    Set-Status "Kopiere Dateien nach Downloads..."
    Add-Log "Exportziel: $destRoot"

    if (Test-Path $script:SetupRoot) {
        Copy-Item -Path (Join-Path $script:SetupRoot "*") -Destination $destSetup -Recurse -Force -ErrorAction Stop
    }
    if (Test-Path $script:OutputRoot) {
        Copy-Item -Path (Join-Path $script:OutputRoot "*") -Destination $destOutput -Recurse -Force -ErrorAction Stop
    }

    Set-Status "Export abgeschlossen: $destRoot"
    Add-Log "Export abgeschlossen: $destRoot"
    return $destRoot
}

function Cleanup-WorkRoot {
    Set-Status "Cleanup laeuft..."
    Add-Log "Cleanup startet."
    Stop-SetupFolderWatcher
    foreach ($p in @($script:RepoRoot, $script:SetupRoot, $script:OutputRoot)) {
        try {
            Remove-PathSafe $p
            Add-Log "Geloescht: $p"
        } catch {
            Add-Log "Cleanup Fehler: $($_.Exception.Message)"
        }
    }

    # Remove WorkRoot if empty
    try {
        if (Test-Path $script:WorkRoot) {
            $remaining = Get-ChildItem -Path $script:WorkRoot -Force -ErrorAction SilentlyContinue
            if (-not $remaining) {
                Remove-Item $script:WorkRoot -Force -ErrorAction SilentlyContinue
            }
        }
    } catch { }

    Set-Status "Cleanup abgeschlossen. Setup beendet."
    Add-Log "Cleanup abgeschlossen."
}

function Run-Background {
    param([scriptblock]$Work, [scriptblock]$OnSuccess, [scriptblock]$OnError)

    $script:IsBusy = $true
    try {
        & $Work
        & $OnSuccess
    } catch {
        $script:LastError = $_
        Add-Log "Fehler: $($_.Exception.Message)"
        if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {
            Add-Log "Fehler-Position: $($_.InvocationInfo.PositionMessage.Trim())"
        }
        & $OnError
    }
}

# XAML UI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="IntuneWin Builder (Company)" Height="720" Width="980"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" ShowInTaskbar="True">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Height" Value="44"/>
            <Setter Property="MinWidth" Value="140"/>
            <Setter Property="Margin" Value="6"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
    </Window.Resources>
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock x:Name="StatusText" Grid.Row="0" Text="Bereit." FontSize="16" FontWeight="SemiBold" Margin="0,0,0,8"/>

        <ProgressBar x:Name="ProgressBar" Grid.Row="1" Height="18" Margin="0,0,0,12"/>

        <Grid Grid.Row="2" Margin="0,0,0,12">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*"/>
                <ColumnDefinition Width="3*"/>
            </Grid.ColumnDefinitions>

            <GroupBox Header="Wizard Schritte (Step 1-5)" Margin="0,0,12,0">
                <StackPanel Margin="8">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Step 1: Tool laden" Width="220"/>
                        <TextBlock x:Name="Step1Status" Text="Offen"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Step 2: Dateien bereitstellen" Width="220"/>
                        <TextBlock x:Name="Step2Status" Text="Offen"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Step 3: EntryPoint waehlen" Width="220"/>
                        <TextBlock x:Name="Step3Status" Text="Offen"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Step 4: IntuneWin erstellen" Width="220"/>
                        <TextBlock x:Name="Step4Status" Text="Offen"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Step 5: Export + Cleanup" Width="220"/>
                        <TextBlock x:Name="Step5Status" Text="Offen"/>
                    </StackPanel>
                </StackPanel>
            </GroupBox>

            <GroupBox Header="EntryPoint" Grid.Column="1">
                <StackPanel Margin="8">
                    <TextBlock Text="Zulassige Endungen: .exe, .ps1, .msi" Margin="0,0,0,6"/>
                    <TextBox x:Name="EntryPointBox" Height="32" IsReadOnly="True" AllowDrop="True" Background="#FFF4F4F4" Padding="6" Text="Drag &amp; Drop oder per Button waehlen"/>
                </StackPanel>
            </GroupBox>
        </Grid>

        <GroupBox Header="Status und Log" Grid.Row="3">
            <Grid Margin="8">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Text="Live-Logausgabe:" Grid.Row="0" Margin="0,0,0,6"/>
                <TextBox x:Name="LogBox" Grid.Row="1" IsReadOnly="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" />
                <StackPanel Orientation="Horizontal" Grid.Row="2" HorizontalAlignment="Right">
                    <Button x:Name="BtnDetails" Content="Details anzeigen"/>
                    <Button x:Name="BtnCopyLog" Content="Log kopieren"/>
                </StackPanel>
            </Grid>
        </GroupBox>

        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,12,0,0">
            <Button x:Name="BtnStart" Content="Start"/>
            <Button x:Name="BtnWeiter" Content="Weiter" IsEnabled="False"/>
            <Button x:Name="BtnEntry" Content="EntryPoint waehlen" IsEnabled="False"/>
            <Button x:Name="BtnBuild" Content="IntuneWin erstellen" IsEnabled="False"/>
            <Button x:Name="BtnCancel" Content="Abbrechen"/>
            <Button x:Name="BtnClose" Content="Schliessen" IsEnabled="False"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$script:Window = [Windows.Markup.XamlReader]::Load($reader)

# Bind controls
$script:StatusText = $script:Window.FindName("StatusText")
$script:ProgressBar = $script:Window.FindName("ProgressBar")
$script:Step1Status = $script:Window.FindName("Step1Status")
$script:Step2Status = $script:Window.FindName("Step2Status")
$script:Step3Status = $script:Window.FindName("Step3Status")
$script:Step4Status = $script:Window.FindName("Step4Status")
$script:Step5Status = $script:Window.FindName("Step5Status")
$script:LogBox = $script:Window.FindName("LogBox")
$script:EntryPointBox = $script:Window.FindName("EntryPointBox")

$script:BtnStart = $script:Window.FindName("BtnStart")
$script:BtnWeiter = $script:Window.FindName("BtnWeiter")
$script:BtnEntry = $script:Window.FindName("BtnEntry")
$script:BtnBuild = $script:Window.FindName("BtnBuild")
$script:BtnCancel = $script:Window.FindName("BtnCancel")
$script:BtnClose = $script:Window.FindName("BtnClose")
$script:BtnCopyLog = $script:Window.FindName("BtnCopyLog")
$script:BtnDetails = $script:Window.FindName("BtnDetails")

# Button handlers
$script:BtnStart.Add_Click({
    try {
        if ($script:IsBusy) { return }

        Set-StepState "1" "Aktiv"
        Set-Status "Repository wird geladen..."
        Add-Log "Start geklickt. Initialisiere Download..."
        Set-Progress $true
        $script:BtnStart.IsEnabled = $false
        $script:BtnCancel.IsEnabled = $true
        Clear-NextActionHighlight

        Run-Background {
            Ensure-Folders
            Download-Repo
    } {
        Set-Progress $false
        Set-StepState "1" "Fertig"
        Set-StepState "2" "Aktiv"
        Set-Status "Bitte legen Sie Ihre Setup-Dateien in den Ordner Setup-folder."
        Add-Log "Bitte legen Sie Ihre Setup-Dateien in den Ordner Setup-folder."
        Start-Process -FilePath "explorer.exe" -ArgumentList $script:SetupRoot
        Bring-WindowToFrontTemporarily -Seconds 4
        Set-Status "Setup-folder geoeffnet. Fenster ggf. in der Taskleiste."
        Add-Log "Fensterstatus nach Explorer: State=$($script:Window.WindowState) Visible=$($script:Window.IsVisible)"
        Start-SetupFolderWatcher
        Update-SetupFolderState
        Set-NextActionButton -Button $script:BtnWeiter
        $script:IsBusy = $false
    } {
            Set-Progress $false
            Set-StepState "1" "Fehler"
            Set-Status "Fehler beim Laden des Tools. Details im Log."
            $script:IsBusy = $false
            $script:BtnStart.IsEnabled = $true
            $script:BtnClose.IsEnabled = $false
        }
    } catch {
        Add-Log "UI-Fehler (Start): $($_.Exception.Message)"
        if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {
            Add-Log "Fehler-Position: $($_.InvocationInfo.PositionMessage.Trim())"
        }
    }
})

$script:BtnWeiter.Add_Click({
    if ($script:IsBusy) { return }
    Set-StepState "2" "Fertig"
    Set-StepState "3" "Aktiv"
    $script:BtnEntry.IsEnabled = $true
    Set-Status "EntryPoint bitte waehlen oder per Drag & Drop setzen."
    Set-NextActionButton -Button $script:BtnEntry
})

$script:BtnEntry.Add_Click({
    if ($script:IsBusy) { return }

    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.InitialDirectory = $script:SetupRoot
    $dlg.Filter = "Install files (*.exe;*.ps1;*.msi)|*.exe;*.ps1;*.msi|All files (*.*)|*.*"
    $dlg.Multiselect = $false
    $dlg.Title = "EntryPoint waehlen"

    $result = $dlg.ShowDialog()
    if ($result -eq $true) {
        if (-not (Validate-EntryPointPath $dlg.FileName)) {
            [System.Windows.MessageBox]::Show("Ungueltiger EntryPoint. Datei muss in Setup-folder liegen und .exe/.ps1/.msi sein.", "Hinweis")
            return
        }
        $script:EntryPointPath = $dlg.FileName
        $script:EntryPointBox.Text = $script:EntryPointPath
        Set-StepState "3" "Fertig"
        Set-StepState "4" "Aktiv"
        $script:BtnBuild.IsEnabled = $true
        Set-Status "EntryPoint gesetzt. Bereit zum Erstellen."
        Set-NextActionButton -Button $script:BtnBuild
    }
})

$script:EntryPointBox.Add_Drop({
    if ($script:IsBusy) { return }
    $files = $args[1].Data.GetData("FileDrop")
    if ($files -and $files.Count -gt 0) {
        $path = $files[0]
        if (-not (Validate-EntryPointPath $path)) {
            [System.Windows.MessageBox]::Show("Ungueltiger EntryPoint. Datei muss in Setup-folder liegen und .exe/.ps1/.msi sein.", "Hinweis")
            return
        }
        $script:EntryPointPath = $path
        $script:EntryPointBox.Text = $script:EntryPointPath
        Set-StepState "3" "Fertig"
        Set-StepState "4" "Aktiv"
        $script:BtnBuild.IsEnabled = $true
        Set-Status "EntryPoint gesetzt. Bereit zum Erstellen."
        Set-NextActionButton -Button $script:BtnBuild
    }
})

$script:BtnBuild.Add_Click({
    try {
        if ($script:IsBusy) { return }
        if (-not (Validate-EntryPointPath $script:EntryPointPath)) {
            [System.Windows.MessageBox]::Show("EntryPoint ist ungueltig oder fehlt.", "Hinweis")
            return
        }

        Set-Progress $true
        $script:BtnBuild.IsEnabled = $false
        $script:BtnEntry.IsEnabled = $false
        $script:BtnWeiter.IsEnabled = $false
        $script:BtnStart.IsEnabled = $false
        $script:BtnCancel.IsEnabled = $true
        Clear-NextActionHighlight

        Run-Background {
            Run-IntuneWin
        } {
            Set-Progress $false
            Set-StepState "4" "Fertig"
            $file = Find-IntuneWinFile
            if ($null -ne $file) {
                Add-Log "Erfolg: $($file.FullName)"
                Set-Status "Erfolg: $($file.Name)"
                $script:HasOutput = $true
            } else {
                Add-Log "Erfolg, aber .intunewin Datei nicht gefunden."
                Set-Status "Erfolg, aber Ausgabe nicht gefunden."
            }

            $choice = [System.Windows.MessageBox]::Show("Moechten Sie eine weitere IntuneWin-Datei erstellen?", "Weitere Pakete", "YesNo")
            if ($choice -eq "Yes") {
                Set-StepState "4" "Fertig"
                Set-StepState "3" "Aktiv"
                $script:EntryPointPath = $null
                $script:EntryPointBox.Text = "Drag & Drop oder per Button waehlen"
                Start-Process -FilePath "explorer.exe" -ArgumentList $script:SetupRoot
                $script:BtnEntry.IsEnabled = $true
                $script:BtnBuild.IsEnabled = $false
                $script:BtnWeiter.IsEnabled = $true
                Set-NextActionButton -Button $script:BtnEntry
                $script:IsBusy = $false
                return
            }

            Set-StepState "5" "Aktiv"
            $exportPath = Export-ToDownloads
            Set-StepState "5" "Fertig"
        Cleanup-WorkRoot
        $script:BtnClose.IsEnabled = $true
        $script:BtnCancel.IsEnabled = $false
        Set-NextActionButton -Button $script:BtnClose
        $script:IsBusy = $false
    } {
            Set-Progress $false
            Set-StepState "4" "Fehler"
            Set-Status "Fehler beim Erstellen. Details im Log."
            $script:BtnBuild.IsEnabled = $true
            $script:BtnEntry.IsEnabled = $true
            $script:IsBusy = $false
        }
    } catch {
        Add-Log "UI-Fehler (Build): $($_.Exception.Message)"
        if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {
            Add-Log "Fehler-Position: $($_.InvocationInfo.PositionMessage.Trim())"
        }
    }
})

$script:BtnCancel.Add_Click({
    $shouldCleanup = $true
    if ($script:IsBusy -and $script:CurrentProcess -ne $null) {
        $confirm = [System.Windows.MessageBox]::Show("Ein Prozess laeuft. Abbrechen und Prozess beenden?", "Abbrechen", "YesNo")
        if ($confirm -eq "Yes") {
            try { $script:CurrentProcess.Kill() } catch { }
            Add-Log "Prozess abgebrochen."
            $script:IsBusy = $false
            Set-Progress $false
        } else {
            return
        }
    } else {
        $confirm = [System.Windows.MessageBox]::Show("Anwendung beenden?", "Abbrechen", "YesNo")
        if ($confirm -ne "Yes") { return }
    }

    $hasOutputFiles = $false
    try {
        $file = Get-ChildItem -Path $script:OutputRoot -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        $hasOutputFiles = $null -ne $file
    } catch {
        $hasOutputFiles = $false
    }

    if ($script:HasOutput -or $hasOutputFiles) {
        $choice = [System.Windows.MessageBox]::Show(
            "Es wurden bereits Dateien erstellt. Jetzt nach Downloads exportieren?",
            "Export vor Beenden",
            "YesNoCancel"
        )
        if ($choice -eq "Yes") {
            Export-ToDownloads | Out-Null
        } elseif ($choice -eq "Cancel") {
            return
        } else {
            Add-Log "Beenden ohne Export. Cleanup der Arbeitsordner wird ausgefuehrt."
        }
    }

    if ($shouldCleanup) {
        Cleanup-WorkRoot
    }
    $script:AllowClose = $true
    $script:Window.Close()
})

$script:BtnClose.Add_Click({
    $script:AllowClose = $true
    $script:Window.Close()
})

$script:BtnCopyLog.Add_Click({
    try {
        Set-Clipboard -Value $script:LogBox.Text
        Set-Status "Log kopiert."
    } catch {
        Add-Log "Log konnte nicht kopiert werden: $($_.Exception.Message)"
    }
})

$script:BtnDetails.Add_Click({
    $script:LogBox.Focus()
    $script:LogBox.ScrollToEnd()
    Set-Status "Details anzeigen."
})

# Catch UI thread exceptions to prevent window from closing silently.
$script:Window.Dispatcher.Add_UnhandledException({
    param($sender, $e)
    try {
        Add-Log "UI-Fehler (Dispatcher): $($e.Exception.Message)"
        if ($e.Exception.StackTrace) {
            Add-Log "StackTrace: $($e.Exception.StackTrace)"
        }
    } catch { }
    $e.Handled = $true
})

$script:Window.Add_Closing({
    param($sender, $e)
    if (-not $script:AllowClose) {
        Add-Log "Fenster-Schliessen blockiert (unerwartet). Bitte 'Schliessen' nutzen."
        $e.Cancel = $true
        return
    }
    Add-Log "Fenster wird geschlossen."
})

$script:Window.Add_Loaded({
    Add-Log "Fenster geladen."
    Bring-WindowToFront
})

[AppDomain]::CurrentDomain.add_UnhandledException({
    param($sender, $e)
    try {
        $ex = $e.ExceptionObject
        Add-Log "UI-Fehler (AppDomain): $($ex.Message)"
        if ($ex.StackTrace) {
            Add-Log "StackTrace: $($ex.StackTrace)"
        }
    } catch { }
})

$script:Window.Add_Closed({
    Add-Log "Fenster geschlossen."
})

$script:Window.Add_StateChanged({
    Add-Log "Fenster-StateChanged: $($script:Window.WindowState)"
})

$script:Window.Add_IsVisibleChanged({
    Add-Log "Fenster-IsVisibleChanged: $($script:Window.IsVisible)"
})

if ([System.Windows.Application]::Current) {
    [System.Windows.Application]::Current.add_Exit({
        Add-Log "Application Exit."
    })
}

# Initial UI state
Set-Status "Bereit. Klicken Sie auf Start."
Set-StepState "1" "Offen"
Set-StepState "2" "Offen"
Set-StepState "3" "Offen"
Set-StepState "4" "Offen"
Set-StepState "5" "Offen"
Init-ButtonDefaults
Set-NextActionButton -Button $script:BtnStart

# Run WPF app
$app = [System.Windows.Application]::Current
if ($null -eq $app) {
    $app = New-Object System.Windows.Application
    $app.ShutdownMode = [System.Windows.ShutdownMode]::OnMainWindowClose
    $app.MainWindow = $script:Window
    $null = $app.Run($script:Window)
} else {
    $null = $script:Window.ShowDialog()
}
