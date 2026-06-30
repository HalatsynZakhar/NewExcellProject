[CmdletBinding()]
param(
    [switch]$NoPause
)

$ErrorActionPreference = "Stop"
$AppDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$PythonScript = Join-Path $AppDir "excel_image_server.py"
$SupervisorScript = Join-Path $AppDir "scripts\supervisor.ps1"
$AutoUpdateBat = Join-Path $AppDir "scripts\auto_update.bat"
$PidFile = Join-Path $AppDir "logs\excel_image_server.pid"
$StoppedProcesses = 0

trap {
    Write-Host ""
    Write-Host "STOP FAILED" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.InvocationInfo.PositionMessage) {
        Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor DarkRed
    }
    Write-Host ""

    if (!$NoPause) {
        Read-Host "Press Enter to close"
    }
    exit 1
}

function Invoke-SchtasksIgnoringMissingTask {
    param([string[]]$Arguments)

    $Process = Start-Process `
        -FilePath "schtasks.exe" `
        -ArgumentList $Arguments `
        -WindowStyle Hidden `
        -Wait `
        -PassThru

    if ($Process.ExitCode -notin @(0, 1)) {
        throw "schtasks exited with code $($Process.ExitCode): $($Arguments -join ' ')"
    }
}

function Stop-ProjectProcesses {
    $Processes = Get-CimInstance Win32_Process |
        Where-Object {
            $CommandLine = [string]$_.CommandLine
            $CommandLine -and (
                $CommandLine.IndexOf($PythonScript, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $CommandLine.IndexOf($SupervisorScript, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $CommandLine.IndexOf($AutoUpdateBat, [StringComparison]::OrdinalIgnoreCase) -ge 0
            )
        } |
        Sort-Object {
            if ($_.Name -in @("python.exe", "pythonw.exe")) { 0 } else { 1 }
        }

    foreach ($Process in $Processes) {
        try {
            Stop-Process -Id $Process.ProcessId -Force -ErrorAction Stop
            Write-Output "Stopped process: $($Process.Name), PID $($Process.ProcessId)"
            $script:StoppedProcesses++
        }
        catch {
            Write-Warning "Could not stop PID $($Process.ProcessId): $($_.Exception.Message)"
        }
    }
}

Write-Output "Stopping ExcelImageServer task..."
Invoke-SchtasksIgnoringMissingTask -Arguments @("/End", "/TN", "ExcelImageServer")

Stop-ProjectProcesses

Write-Output "Deleting scheduled task..."
Invoke-SchtasksIgnoringMissingTask -Arguments @("/Delete", "/TN", "ExcelImageServer", "/F")

Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue

Start-Sleep -Seconds 1
Stop-ProjectProcesses

Write-Output "Stop complete:"
Write-Output "- Python excel_image_server.py stopped;"
Write-Output "- BAT and supervisor stopped;"
Write-Output "- scheduled task ExcelImageServer deleted;"
Write-Output "- PID file deleted;"
Write-Output "- stopped project processes: $StoppedProcesses."

Write-Host ""
Write-Host "ALL EXCEL IMAGE SERVER PROCESSES WERE STOPPED" -ForegroundColor Green
Write-Host ""

if (!$NoPause) {
    Read-Host "Press Enter to close"
}
