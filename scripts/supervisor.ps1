[CmdletBinding()]
param(
    [string]$Branch = "main",
    [ValidateRange(1, 1439)]
    [int]$CheckIntervalMinutes = 5
)

$ErrorActionPreference = "Stop"
$AppDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$PythonExe = Join-Path $AppDir ".venv\Scripts\python.exe"
$PythonScript = Join-Path $AppDir "excel_image_server.py"
$PidFile = Join-Path $AppDir "logs\excel_image_server.pid"
$LocalDiagnosticLog = Join-Path $AppDir "logs\supervisor-local.log"
$WorkerOutputLog = Join-Path $AppDir "logs\python-output.log"
$WorkerErrorLog = Join-Path $AppDir "logs\python-error.log"
$ConsoleLogFile = Join-Path $AppDir "logs\excel_image_server.log"
$MaxLogLines = 2000
$LogMutexName = "Global\ExcelImageServerSharedLog"
$Worker = $null

Set-Location $AppDir

try {
    $Config = Get-Content (Join-Path $AppDir "config.json") -Raw -Encoding UTF8 |
        ConvertFrom-Json
    if ($Config.public_log_dir) {
        $LogFilename = if ($Config.public_log_filename) {
            [string]$Config.public_log_filename
        }
        else {
            "excel_image_server.log"
        }
        $ConsoleLogFile = Join-Path ([string]$Config.public_log_dir) $LogFilename
    }
    if ($Config.max_log_lines) {
        $MaxLogLines = [int]$Config.max_log_lines
    }
}
catch {
    Write-Warning "Could not read public log settings: $($_.Exception.Message)"
}

function Get-NowText {
    return (Get-Date).ToString(
        "yyyy-MM-ddTHH:mm:sszzz",
        [Globalization.CultureInfo]::InvariantCulture
    )
}

function Write-SupervisorLog {
    param(
        [string]$Message,
        [string]$LocalDetails = ""
    )

    $PublicLine = "[$(Get-NowText)] [supervisor] $Message"
    $LocalLine = $PublicLine
    if ($LocalDetails) {
        $LocalLine = "$PublicLine Details: $LocalDetails"
    }
    Write-Output $LocalLine

    try {
        New-Item -ItemType Directory -Path (Split-Path $LocalDiagnosticLog) `
            -Force | Out-Null
        Add-Content -LiteralPath $LocalDiagnosticLog -Value $LocalLine -Encoding UTF8

        $LocalLines = Get-Content -LiteralPath $LocalDiagnosticLog
        if ($LocalLines.Count -gt $MaxLogLines) {
            $LocalLines |
                Select-Object -Last $MaxLogLines |
                Set-Content -LiteralPath $LocalDiagnosticLog -Encoding UTF8
        }
    }
    catch {
        Write-Warning "Could not write local diagnostic log: $($_.Exception.Message)"
    }

    $LogMutex = $null
    $LockAcquired = $false
    try {
        $LogMutex = [System.Threading.Mutex]::new($false, $LogMutexName)
        try {
            $LockAcquired = $LogMutex.WaitOne(30000)
        }
        catch [System.Threading.AbandonedMutexException] {
            $LockAcquired = $true
        }
        if (!$LockAcquired) {
            throw "Could not acquire shared log lock in 30 seconds."
        }

        $LogDirectory = Split-Path $ConsoleLogFile
        if ($LogDirectory) {
            New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
        }

        $Utf8Bom = New-Object System.Text.UTF8Encoding($true)
        $Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        if (!(Test-Path $ConsoleLogFile)) {
            [System.IO.File]::WriteAllBytes($ConsoleLogFile, $Utf8Bom.GetPreamble())
        }
        else {
            $ExistingBytes = [System.IO.File]::ReadAllBytes($ConsoleLogFile)
            $Preamble = $Utf8Bom.GetPreamble()
            $HasBom = (
                $ExistingBytes.Length -ge $Preamble.Length -and
                $ExistingBytes[0] -eq $Preamble[0] -and
                $ExistingBytes[1] -eq $Preamble[1] -and
                $ExistingBytes[2] -eq $Preamble[2]
            )
            if (!$HasBom) {
                $BytesWithBom = [byte[]]::new($Preamble.Length + $ExistingBytes.Length)
                [Array]::Copy($Preamble, 0, $BytesWithBom, 0, $Preamble.Length)
                [Array]::Copy($ExistingBytes, 0, $BytesWithBom, $Preamble.Length, $ExistingBytes.Length)
                [System.IO.File]::WriteAllBytes($ConsoleLogFile, $BytesWithBom)
            }
        }

        [System.IO.File]::AppendAllText(
            $ConsoleLogFile,
            "$PublicLine$([Environment]::NewLine)",
            $Utf8NoBom
        )

        $Lines = [System.IO.File]::ReadAllLines($ConsoleLogFile, [System.Text.Encoding]::UTF8)
        if ($Lines.Count -gt $MaxLogLines) {
            $NewestLines = $Lines | Select-Object -Last $MaxLogLines
            [System.IO.File]::WriteAllLines($ConsoleLogFile, $NewestLines, $Utf8Bom)
        }
    }
    catch {
        Write-Warning "Could not write supervisor log: $($_.Exception.Message)"
    }
    finally {
        if ($LockAcquired) {
            $LogMutex.ReleaseMutex()
        }
        if ($null -ne $LogMutex) {
            $LogMutex.Dispose()
        }
    }
}

function Test-WorkerRunning {
    return ($null -ne $script:Worker -and !$script:Worker.HasExited)
}

function Start-Worker {
    if (Test-WorkerRunning) {
        return
    }

    if (!(Test-Path $PythonExe)) {
        throw "Python virtual environment not found: $PythonExe"
    }

    New-Item -ItemType Directory -Path (Split-Path $PidFile) -Force | Out-Null
    Remove-Item -LiteralPath $WorkerOutputLog -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $WorkerErrorLog -Force -ErrorAction SilentlyContinue

    $script:Worker = Start-Process `
        -FilePath $PythonExe `
        -ArgumentList "`"$PythonScript`"" `
        -WorkingDirectory $AppDir `
        -WindowStyle Hidden `
        -RedirectStandardOutput $WorkerOutputLog `
        -RedirectStandardError $WorkerErrorLog `
        -PassThru

    Set-Content -LiteralPath $PidFile -Value $script:Worker.Id -Encoding ASCII
    Write-SupervisorLog "Excel image server started." "PID: $($script:Worker.Id)"

    Start-Sleep -Seconds 8
    $script:Worker.Refresh()
    if ($script:Worker.HasExited) {
        $ExitCode = $script:Worker.ExitCode
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue

        $Details = @()
        if (Test-Path $WorkerOutputLog) {
            $Details += Get-Content -LiteralPath $WorkerOutputLog -Tail 20
        }
        if (Test-Path $WorkerErrorLog) {
            $Details += Get-Content -LiteralPath $WorkerErrorLog -Tail 20
        }
        $DetailsText = ($Details | Where-Object { $_ }) -join " | "
        if (!$DetailsText) {
            $DetailsText = "Python produced no details."
        }

        $script:Worker = $null
        throw "Python exited immediately with code $ExitCode. Details: $DetailsText"
    }

    Write-SupervisorLog "Excel image server is running." "PID: $($script:Worker.Id)"
}

function Stop-Worker {
    if (!(Test-WorkerRunning)) {
        return
    }

    $WorkerId = $script:Worker.Id
    Stop-Process -Id $WorkerId -Force
    $script:Worker.WaitForExit()
    Write-SupervisorLog "Excel image server stopped." "PID: $WorkerId"
    $script:Worker = $null
    Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
}

function Restore-WorkerFromPidFile {
    if (!(Test-Path $PidFile)) {
        return
    }

    $SavedPid = 0
    if (![int]::TryParse((Get-Content $PidFile -Raw).Trim(), [ref]$SavedPid)) {
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
        return
    }

    $ProcessInfo = Get-CimInstance Win32_Process `
        -Filter "ProcessId = $SavedPid" `
        -ErrorAction SilentlyContinue

    if ($null -eq $ProcessInfo -or $ProcessInfo.CommandLine -notlike "*excel_image_server.py*") {
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
        return
    }

    $script:Worker = Get-Process -Id $SavedPid -ErrorAction SilentlyContinue
    if (Test-WorkerRunning) {
        Write-SupervisorLog "Excel image server is already running." "PID: $SavedPid"
    }
}

function Test-GitAvailableForProject {
    return ((Get-Command git -ErrorAction SilentlyContinue) -and (Test-Path (Join-Path $AppDir ".git")))
}

function Test-UpdateRequired {
    if (!(Test-GitAvailableForProject)) {
        return $false
    }

    git fetch origin $Branch --quiet
    if ($LASTEXITCODE -ne 0) {
        throw "Could not check origin/$Branch."
    }

    $Local = (git rev-parse HEAD).Trim()
    $Remote = (git rev-parse "origin/$Branch").Trim()
    $WorkingTreeChanges = git status --porcelain
    if ($LASTEXITCODE -ne 0) {
        throw "Could not inspect Git working tree."
    }

    return (($Local -ne $Remote) -or [bool]$WorkingTreeChanges)
}

function Get-ConfiguredPort {
    try {
        $Config = Get-Content (Join-Path $AppDir "config.json") -Raw -Encoding UTF8 |
            ConvertFrom-Json
        if ($Config.port) {
            return [int]$Config.port
        }
    }
    catch {
        Write-SupervisorLog "Could not read configured port." $_.Exception.Message
    }
    return 8091
}

function Test-ServerBusy {
    if (!(Test-WorkerRunning)) {
        return $false
    }
    $Port = Get-ConfiguredPort
    try {
        $Response = Invoke-WebRequest `
            -Uri "http://127.0.0.1:$Port/status" `
            -UseBasicParsing `
            -TimeoutSec 5
        $Payload = $Response.Content | ConvertFrom-Json
        return [bool]$Payload.busy
    }
    catch {
        Write-SupervisorLog "Could not query server busy state." $_.Exception.Message
        return $true
    }
}

function Wait-ServerIdleForUpdate {
    param(
        [int]$MaxWaitSeconds = 3600
    )

    $Deadline = (Get-Date).AddSeconds($MaxWaitSeconds)
    while (Test-ServerBusy) {
        if ((Get-Date) -ge $Deadline) {
            Write-SupervisorLog "Update postponed because processing is still active."
            return $false
        }
        Write-SupervisorLog "Update waiting for current Excel processing to finish."
        Start-Sleep -Seconds 30
    }
    return $true
}

function Update-Project {
    $Remote = (git rev-parse "origin/$Branch").Trim()
    Write-SupervisorLog "Installing service update." "Branch: $Branch"

    $ConfigBackup = Join-Path $env:TEMP "ExcelImageServer-config.json"
    if (Test-Path (Join-Path $AppDir "config.json")) {
        Copy-Item (Join-Path $AppDir "config.json") $ConfigBackup -Force
    }

    git reset --hard $Remote
    if ($LASTEXITCODE -ne 0) {
        throw "Could not sync code with origin/$Branch."
    }

    git clean -fd
    if ($LASTEXITCODE -ne 0) {
        throw "Could not remove untracked files."
    }

    if (Test-Path $ConfigBackup) {
        Copy-Item $ConfigBackup (Join-Path $AppDir "config.json") -Force
        Remove-Item $ConfigBackup -Force -ErrorAction SilentlyContinue
    }

    & $PythonExe -m pip install -r (Join-Path $AppDir "requirements.txt")
    if ($LASTEXITCODE -ne 0) {
        throw "Could not install dependencies after update."
    }
}

Write-SupervisorLog "Supervisor started." (
    "Branch: $Branch; Git check every $CheckIntervalMinutes minutes."
)
Restore-WorkerFromPidFile
$SupervisorTimer = [Diagnostics.Stopwatch]::StartNew()
$NextHeartbeatSeconds = 300.0
$NextGitCheckSeconds = 0.0

try {
    while ($true) {
        try {
            if ($SupervisorTimer.Elapsed.TotalSeconds -ge $NextGitCheckSeconds) {
                try {
                    if (Test-UpdateRequired) {
                        Write-SupervisorLog "Update or local divergence found."
                        if (!(Wait-ServerIdleForUpdate)) {
                            continue
                        }
                        Stop-Worker
                        Update-Project
                        Write-SupervisorLog "Update installed. Restarting supervisor."
                        exit 75
                    }
                }
                finally {
                    $NextGitCheckSeconds = (
                        $SupervisorTimer.Elapsed.TotalSeconds +
                        $CheckIntervalMinutes * 60
                    )
                }
            }

            Start-Worker

            if ($SupervisorTimer.Elapsed.TotalSeconds -ge $NextHeartbeatSeconds) {
                Write-SupervisorLog "Supervisor heartbeat: server control is active."
                $NextHeartbeatSeconds = $SupervisorTimer.Elapsed.TotalSeconds + 300
            }
        }
        catch {
            Write-SupervisorLog "Service control error." $_.Exception.Message
            try {
                Start-Worker
            }
            catch {
                Write-SupervisorLog "Excel image server could not start." $_.Exception.Message
            }
        }

        Start-Sleep -Seconds 60
    }
}
finally {
    Stop-Worker
}
