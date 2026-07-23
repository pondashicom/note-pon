[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repositoryRoot 'NotePon.csproj'
$executablePath = Join-Path $repositoryRoot 'bin\Release\net10.0-windows\NOTE-PON.exe'

function Get-WorkerProcess {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ParentProcessId,

        [int]$ExcludedProcessId = 0
    )

    Get-CimInstance Win32_Process -Filter "ParentProcessId = $ParentProcessId" |
        Where-Object {
            $_.ExecutablePath -eq $executablePath -and
            $_.ProcessId -ne $ExcludedProcessId
        } |
        Select-Object -First 1
}

function Wait-ForWorker {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ParentProcessId,

        [int]$ExcludedProcessId = 0,

        [int]$TimeoutMilliseconds = 5000
    )

    $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMilliseconds)
    do {
        $worker = Get-WorkerProcess `
            -ParentProcessId $ParentProcessId `
            -ExcludedProcessId $ExcludedProcessId
        if ($worker) {
            return $worker
        }

        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)

    return $null
}

function Stop-TestProcess {
    param(
        [System.Diagnostics.Process]$Process
    )

    if (-not $Process) {
        return
    }

    $Process.Refresh()
    if ($Process.HasExited) {
        return
    }

    [void]$Process.CloseMainWindow()
    if ($Process.WaitForExit(3000)) {
        return
    }

    $current = Get-Process -Id $Process.Id -ErrorAction SilentlyContinue
    if ($current -and $current.Path -eq $executablePath) {
        Stop-Process -Id $Process.Id -Force
    }
}

$previousLib = $env:LIB
try {
    $env:LIB = ''
    & dotnet build $projectPath --configuration Release --no-incremental -warnaserror
    if ($LASTEXITCODE -ne 0) {
        throw "NOTE-PON build failed with exit code $LASTEXITCODE."
    }

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class NotePonFaultTestProcessControl
{
    private const int GwlExStyle = -20;

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint virtualKey);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool UnregisterHotKey(IntPtr window, int id);

    [DllImport("ntdll.dll")]
    public static extern int NtSuspendProcess(IntPtr processHandle);

    [DllImport("ntdll.dll")]
    public static extern int NtResumeProcess(IntPtr processHandle);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW", SetLastError = true)]
    private static extern IntPtr GetWindowLongPtr(IntPtr window, int index);

    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(
        IntPtr window,
        uint message,
        IntPtr wParam,
        IntPtr lParam);

    public static long GetExtendedWindowStyle(IntPtr window)
    {
        return GetWindowLongPtr(window, GwlExStyle).ToInt64();
    }
}
'@
}
finally {
    $env:LIB = $previousLib
}

$existing = Get-CimInstance Win32_Process |
    Where-Object { $_.ExecutablePath -eq $executablePath }
if ($existing) {
    throw 'Close the development build of NOTE-PON before running this test.'
}

$requestedHotKeys = @(
    @{ Id = 899; Modifiers = 0x0003; VirtualKey = 0x79 },
    @{ Id = 900; Modifiers = 0x0003; VirtualKey = 0x7A },
    @{ Id = 901; Modifiers = 0; VirtualKey = 0xAF },
    @{ Id = 902; Modifiers = 0; VirtualKey = 0xAE }
)
foreach ($hotKey in $requestedHotKeys) {
    if (-not [NotePonFaultTestProcessControl]::RegisterHotKey(
        [IntPtr]::Zero,
        $hotKey.Id,
        $hotKey.Modifiers,
        $hotKey.VirtualKey)) {
        throw 'A requested hotkey is already in use before NOTE-PON starts.'
    }

    [void][NotePonFaultTestProcessControl]::UnregisterHotKey(
        [IntPtr]::Zero,
        $hotKey.Id)
}

$primary = $null
$duplicate = $null
$suspendedWorkerId = 0
$workerWasSuspended = $false

try {
    $primary = Start-Process -FilePath $executablePath -WindowStyle Hidden -PassThru
    $worker = Wait-ForWorker -ParentProcessId $primary.Id
    if (-not $worker) {
        throw 'The PowerPoint worker did not start.'
    }

    $primary.Refresh()
    if ($primary.HasExited -or -not $primary.Responding) {
        throw 'The NOTE-PON UI process is not responding after startup.'
    }

    $windowDeadline = [DateTime]::UtcNow.AddSeconds(3)
    while ($primary.MainWindowHandle -eq [IntPtr]::Zero -and
        [DateTime]::UtcNow -lt $windowDeadline) {
        Start-Sleep -Milliseconds 100
        $primary.Refresh()
    }

    $extendedWindowStyle =
        [NotePonFaultTestProcessControl]::GetExtendedWindowStyle($primary.MainWindowHandle)
    if (($extendedWindowStyle -band 0x08000000) -eq 0) {
        throw 'The NOTE-PON window can still take keyboard focus.'
    }

    $mouseActivateResult = [NotePonFaultTestProcessControl]::SendMessage(
        $primary.MainWindowHandle,
        0x0021,
        [IntPtr]::Zero,
        [IntPtr]::Zero)
    if ($mouseActivateResult.ToInt32() -ne 3) {
        throw 'The NOTE-PON window did not reject mouse activation.'
    }

    foreach ($hotKey in $requestedHotKeys) {
        if ([NotePonFaultTestProcessControl]::RegisterHotKey(
            [IntPtr]::Zero,
            $hotKey.Id,
            $hotKey.Modifiers,
            $hotKey.VirtualKey)) {
            [void][NotePonFaultTestProcessControl]::UnregisterHotKey(
                [IntPtr]::Zero,
                $hotKey.Id)
            throw 'NOTE-PON did not register a requested volume hotkey.'
        }
    }

    $duplicate = Start-Process -FilePath $executablePath -WindowStyle Hidden -PassThru
    if (-not $duplicate.WaitForExit(3000)) {
        throw 'A duplicate NOTE-PON instance was not rejected.'
    }

    $suspendedWorkerId = [int]$worker.ProcessId
    $workerProcess = Get-Process -Id $suspendedWorkerId
    $suspendResult =
        [NotePonFaultTestProcessControl]::NtSuspendProcess($workerProcess.Handle)
    if ($suspendResult -ne 0) {
        throw "NtSuspendProcess failed with status $suspendResult."
    }
    $workerWasSuspended = $true
    $recoveryStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $replacement = Wait-ForWorker `
        -ParentProcessId $primary.Id `
        -ExcludedProcessId $suspendedWorkerId `
        -TimeoutMilliseconds 7000
    if (-not $replacement) {
        throw 'The suspended PowerPoint worker was not replaced after the timeout.'
    }
    $recoveryStopwatch.Stop()

    if ($recoveryStopwatch.ElapsedMilliseconds -lt 3500) {
        throw 'The PowerPoint worker restarted without the expected backoff.'
    }

    $primary.Refresh()
    if ($primary.HasExited -or -not $primary.Responding) {
        throw 'The NOTE-PON UI process stopped responding during worker recovery.'
    }

    if (Get-Process -Id $suspendedWorkerId -ErrorAction SilentlyContinue) {
        throw 'The timed-out PowerPoint worker is still running.'
    }

    [pscustomobject]@{
        Result = 'PASS'
        PrimaryProcessId = $primary.Id
        TimedOutWorkerProcessId = $suspendedWorkerId
        ReplacementWorkerProcessId = $replacement.ProcessId
        DuplicateInstancePrevented = $true
        UiRemainedResponsive = $true
        VolumeHotKeysRegistered = $true
        ShortcutHotKeysRegistered = $true
        WindowDoesNotActivate = $true
        MouseClickDoesNotActivate = $true
        WorkerRecoveryMilliseconds = $recoveryStopwatch.ElapsedMilliseconds
    }
}
finally {
    if ($workerWasSuspended -and $suspendedWorkerId -gt 0) {
        $remainingWorker = Get-Process -Id $suspendedWorkerId -ErrorAction SilentlyContinue
        if ($remainingWorker) {
            [void][NotePonFaultTestProcessControl]::NtResumeProcess($remainingWorker.Handle)
        }
    }

    Stop-TestProcess -Process $duplicate
    Stop-TestProcess -Process $primary
}

Start-Sleep -Milliseconds 500
$survivors = Get-CimInstance Win32_Process |
    Where-Object { $_.ExecutablePath -eq $executablePath }
if ($survivors) {
    throw 'One or more NOTE-PON test processes remained after cleanup.'
}
