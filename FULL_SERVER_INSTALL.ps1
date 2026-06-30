[CmdletBinding()]
param(
    [string]$InstallDir = (Join-Path $env:ProgramData "ExcelImageServer"),
    [string]$Repository = "https://github.com/HalatsynZakhar/NewExcellProject.git",
    [string]$Branch = "main"
)

$ErrorActionPreference = "Stop"
$InstallLog = Join-Path $env:ProgramData "ExcelImageServer-install.log"
$TranscriptStarted = $false
$TaskRunAs = "SYSTEM"
$TaskPassword = $null

[Net.ServicePointManager]::SecurityProtocol = `
    [Net.ServicePointManager]::SecurityProtocol -bor `
    [Net.SecurityProtocolType]::Tls12

try {
    Start-Transcript -Path $InstallLog -Append -Force | Out-Null
    $TranscriptStarted = $true
}
catch {
    Write-Warning "Could not start installer transcript: $($_.Exception.Message)"
}

trap {
    $ErrorMessage = $_.Exception.Message
    $ErrorPosition = $_.InvocationInfo.PositionMessage

    Write-Host ""
    Write-Host "INSTALL FAILED" -ForegroundColor Red
    Write-Host $ErrorMessage -ForegroundColor Red
    if ($ErrorPosition) {
        Write-Host $ErrorPosition -ForegroundColor DarkRed
    }
    Write-Host ""
    Write-Host "Installer log: $InstallLog" -ForegroundColor Yellow

    if ($TranscriptStarted) {
        Stop-Transcript | Out-Null
        $TranscriptStarted = $false
    }

    Read-Host "Press Enter to close"
    exit 1
}

function Update-CurrentPath {
    $MachinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $UserPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$MachinePath;$UserPath"
}

function Install-WingetPackage {
    param(
        [string]$PackageId,
        [string]$DisplayName
    )

    Write-Output "Installing $DisplayName..."
    winget install `
        --id $PackageId `
        --exact `
        --silent `
        --accept-package-agreements `
        --accept-source-agreements

    if ($LASTEXITCODE -ne 0) {
        throw "Could not install $DisplayName through winget."
    }
}

function Assert-ValidSignature {
    param(
        [string]$Path,
        [string]$DisplayName
    )

    $Signature = Get-AuthenticodeSignature -FilePath $Path
    if ($Signature.Status -ne "Valid") {
        throw "Invalid digital signature for $DisplayName installer."
    }
}

function Install-PythonDirect {
    $Version = "3.13.14"
    $Installer = Join-Path $env:TEMP "python-$Version-amd64.exe"
    $Url = "https://www.python.org/ftp/python/$Version/python-$Version-amd64.exe"

    Write-Output "Downloading Python $Version from python.org..."
    Invoke-WebRequest -Uri $Url -OutFile $Installer -UseBasicParsing
    Assert-ValidSignature -Path $Installer -DisplayName "Python"

    $Process = Start-Process `
        -FilePath $Installer `
        -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" `
        -Wait `
        -PassThru
    if ($Process.ExitCode -ne 0) {
        throw "Python installer exited with code $($Process.ExitCode)."
    }
}

function Install-GitDirect {
    Write-Output "Reading latest Git for Windows release..."
    $Release = Invoke-RestMethod `
        -Uri "https://api.github.com/repos/git-for-windows/git/releases/latest" `
        -Headers @{ "User-Agent" = "ExcelImageServer-Installer" }

    $Asset = $Release.assets |
        Where-Object { $_.name -match "^Git-.+-64-bit\.exe$" } |
        Select-Object -First 1
    if ($null -eq $Asset) {
        throw "Could not find 64-bit Git installer."
    }

    $Installer = Join-Path $env:TEMP $Asset.name
    Write-Output "Downloading $($Asset.name)..."
    Invoke-WebRequest -Uri $Asset.browser_download_url -OutFile $Installer -UseBasicParsing
    Assert-ValidSignature -Path $Installer -DisplayName "Git"

    $Process = Start-Process `
        -FilePath $Installer `
        -ArgumentList "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES /SP-" `
        -Wait `
        -PassThru
    if ($Process.ExitCode -ne 0) {
        throw "Git installer exited with code $($Process.ExitCode)."
    }
}

function Test-PythonInstalled {
    foreach ($Command in @("py", "python")) {
        if (!(Get-Command $Command -ErrorAction SilentlyContinue)) {
            continue
        }
        & $Command --version 2>$null | Out-Null
        if ($LASTEXITCODE -eq 0) {
            return $true
        }
    }
    return $false
}

function Get-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        return @("py", "-3")
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        return @("python")
    }
    throw "Python 3 was not found after installation."
}

function Test-UncPath {
    param([string]$Path)
    return $Path.StartsWith("\\")
}

function Convert-SecureStringToPlainText {
    param([Security.SecureString]$SecureString)

    $Pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($Pointer)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Pointer)
    }
}

function Grant-LocalFolderAccess {
    param(
        [string]$Path,
        [string]$Identity,
        [ValidateSet("Read", "Modify")]
        [string]$Access
    )

    New-Item -ItemType Directory -Path $Path -Force | Out-Null

    $AclIdentity = if ($Identity -eq "SYSTEM") {
        "*S-1-5-18"
    }
    else {
        $Identity
    }
    $Rights = if ($Access -eq "Read") { "RX" } else { "M" }

    Write-Output "Granting $Access access for ${Identity}: $Path"
    & icacls.exe $Path /grant "${AclIdentity}:(OI)(CI)$Rights" /T /C | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Could not grant $Identity access to $Path."
    }
}

function Stop-ExistingRuntime {
    Write-Output "Stopping previous ExcelImageServer runtime..."

    Start-Process `
        -FilePath "schtasks.exe" `
        -ArgumentList @("/End", "/TN", "ExcelImageServer") `
        -WindowStyle Hidden `
        -Wait | Out-Null

    $PythonScript = Join-Path $InstallDir "excel_image_server.py"
    $SupervisorScript = Join-Path $InstallDir "scripts\supervisor.ps1"
    $AutoUpdateBat = Join-Path $InstallDir "scripts\auto_update.bat"

    $ProjectProcesses = Get-CimInstance Win32_Process |
        Where-Object {
            $CommandLine = [string]$_.CommandLine
            $CommandLine -and (
                $CommandLine.IndexOf($PythonScript, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $CommandLine.IndexOf($SupervisorScript, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $CommandLine.IndexOf($AutoUpdateBat, [StringComparison]::OrdinalIgnoreCase) -ge 0
            )
        }

    foreach ($Process in $ProjectProcesses) {
        try {
            Stop-Process -Id $Process.ProcessId -Force -ErrorAction Stop
            Write-Output "Stopped old process: $($Process.Name), PID $($Process.ProcessId)"
        }
        catch {
            Write-Warning "Could not stop PID $($Process.ProcessId): $($_.Exception.Message)"
        }
    }

    Remove-Item `
        -LiteralPath (Join-Path $InstallDir "logs\excel_image_server.pid") `
        -Force `
        -ErrorAction SilentlyContinue
}

function Install-ProjectFiles {
    $SourceDir = Split-Path -Parent $PSCommandPath
    $SourceHasProject = Test-Path (Join-Path $SourceDir "excel_image_server.py")

    if ($Repository) {
        if (!(Get-Command git -ErrorAction SilentlyContinue)) {
            throw "Git is required when Repository is provided."
        }
        Set-Location "C:\"
        if (Test-Path (Join-Path $InstallDir ".git")) {
            Write-Output "Project already exists: $InstallDir"
            Stop-ExistingRuntime
            Set-Location $InstallDir
            $ConfigBackup = Join-Path $env:TEMP "ExcelImageServer-config.json"
            if (Test-Path "config.json") {
                Copy-Item "config.json" $ConfigBackup -Force
            }
            git fetch origin $Branch
            if ($LASTEXITCODE -ne 0) {
                throw "Could not fetch project updates."
            }
            git reset --hard "origin/$Branch"
            if ($LASTEXITCODE -ne 0) {
                throw "Could not update existing project."
            }
            if (Test-Path $ConfigBackup) {
                Copy-Item $ConfigBackup "config.json" -Force
                Remove-Item $ConfigBackup -Force -ErrorAction SilentlyContinue
            }
        }
        elseif (Test-Path $InstallDir) {
            throw "Folder $InstallDir exists but is not a Git repository."
        }
        else {
            git clone --branch $Branch $Repository $InstallDir
            if ($LASTEXITCODE -ne 0) {
                throw "Could not clone project."
            }
        }
        return
    }

    if (!$SourceHasProject) {
        throw (
            "Repository parameter is empty and installer was not started from " +
            "a folder containing excel_image_server.py."
        )
    }

    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
    Stop-ExistingRuntime
    Write-Output "Copying project files from $SourceDir to $InstallDir..."
    robocopy `
        $SourceDir `
        $InstallDir `
        /E `
        /XD ".venv" "logs" "work" "__pycache__" `
        /XF "config.json" "*.pyc" "*.pyo" "*.pid" "*.log" |
        Out-Host
    if ($LASTEXITCODE -gt 7) {
        throw "robocopy failed with code $LASTEXITCODE."
    }
}

function Configure-StorageAccess {
    $ConfigPath = Join-Path $InstallDir "config.json"
    try {
        $Config = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw (
            "config.json contains invalid JSON: $($_.Exception.Message). " +
            "Use double backslashes in Windows paths."
        )
    }

    $ImagesPath = [string]$Config.images_dir
    $PublicLogPath = if ($Config.public_log_dir) {
        [string]$Config.public_log_dir
    }
    else {
        Join-Path $InstallDir "logs"
    }
    $WorkPath = if ($Config.work_dir) {
        [string]$Config.work_dir
    }
    else {
        Join-Path $InstallDir "work"
    }

    $UsesNetworkPath = (
        (Test-UncPath $ImagesPath) -or
        (Test-UncPath $PublicLogPath)
    )

    if ($UsesNetworkPath) {
        Write-Host ""
        Write-Host "NETWORK UNC PATH DETECTED" -ForegroundColor Yellow
        Write-Host "The scheduled task must run as a Windows user with SMB access." -ForegroundColor Yellow
        $DefaultUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        $EnteredUser = Read-Host "Task user [$DefaultUser]"
        $script:TaskRunAs = if ([string]::IsNullOrWhiteSpace($EnteredUser)) {
            $DefaultUser
        }
        else {
            $EnteredUser.Trim()
        }

        $SecurePassword = Read-Host "Password for $script:TaskRunAs" -AsSecureString
        if ($SecurePassword.Length -eq 0) {
            throw "Password cannot be empty for a network folder task user."
        }
        $script:TaskPassword = Convert-SecureStringToPlainText -SecureString $SecurePassword
    }
    else {
        $ConfiguredUser = [string]$Config.task_user
        $script:TaskRunAs = if ([string]::IsNullOrWhiteSpace($ConfiguredUser)) {
            "SYSTEM"
        }
        else {
            $ConfiguredUser.Trim()
        }

        if ($script:TaskRunAs -ne "SYSTEM") {
            $SecurePassword = Read-Host "Password for $script:TaskRunAs" -AsSecureString
            $script:TaskPassword = Convert-SecureStringToPlainText -SecureString $SecurePassword
        }
    }

    if (!(Test-UncPath $WorkPath)) {
        Grant-LocalFolderAccess -Path $WorkPath -Identity $script:TaskRunAs -Access Modify
    }
    if (!(Test-UncPath $PublicLogPath)) {
        Grant-LocalFolderAccess -Path $PublicLogPath -Identity $script:TaskRunAs -Access Modify
    }
    if (!(Test-UncPath $ImagesPath)) {
        Grant-LocalFolderAccess -Path $ImagesPath -Identity $script:TaskRunAs -Access Read
    }
    Grant-LocalFolderAccess -Path $InstallDir -Identity $script:TaskRunAs -Access Modify

    if (Get-Command git -ErrorAction SilentlyContinue) {
        $SafeDirectory = $InstallDir.Replace("\", "/")
        $ExistingSafeDirectories = @(
            git config --system --get-all safe.directory 2>$null
        )
        if ($ExistingSafeDirectories -notcontains $SafeDirectory) {
            git config --system --add safe.directory $SafeDirectory
            if ($LASTEXITCODE -ne 0) {
                throw "Could not add Git safe.directory: $InstallDir"
            }
        }
    }

    Write-Output "Scheduled task will run as: $script:TaskRunAs"
}

function Initialize-PythonEnvironment {
    Set-Location $InstallDir

    if (!(Test-Path ".venv\Scripts\python.exe")) {
        $Python = Get-PythonCommand
        Write-Output "Creating virtual environment..."
        if ($Python.Count -eq 2) {
            & $Python[0] $Python[1] -m venv .venv
        }
        else {
            & $Python[0] -m venv .venv
        }
        if ($LASTEXITCODE -ne 0) {
            throw "Could not create virtual environment."
        }
    }

    $VenvPython = Join-Path $InstallDir ".venv\Scripts\python.exe"
    & $VenvPython -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) {
        throw "Could not upgrade pip."
    }

    Write-Output "Installing Python dependencies..."
    & $VenvPython -m pip install -r "$InstallDir\requirements.txt"
    if ($LASTEXITCODE -ne 0) {
        throw "Could not install dependencies from requirements.txt."
    }

    New-Item -ItemType Directory -Path "$InstallDir\logs" -Force | Out-Null
    & $VenvPython -m compileall -q "$InstallDir\excel_image_server.py"
    if ($LASTEXITCODE -ne 0) {
        throw "Python compile check failed."
    }

    Write-Output "Resolved public log:"
    & $VenvPython "$InstallDir\excel_image_server.py" --print-config-log
    if ($LASTEXITCODE -ne 0) {
        throw "Could not load config.json."
    }
}

function Configure-Firewall {
    $Config = Get-Content (Join-Path $InstallDir "config.json") -Raw -Encoding UTF8 |
        ConvertFrom-Json
    $Port = if ($Config.port) { [int]$Config.port } else { 8091 }
    $RuleName = "ExcelImageServer-$Port"

    Write-Output "Opening Windows Firewall TCP port $Port..."
    $Existing = Get-NetFirewallRule -DisplayName $RuleName -ErrorAction SilentlyContinue
    if ($Existing) {
        Remove-NetFirewallRule -DisplayName $RuleName
    }
    New-NetFirewallRule `
        -DisplayName $RuleName `
        -Direction Inbound `
        -Action Allow `
        -Protocol TCP `
        -LocalPort $Port | Out-Null
}

function Create-ServerTask {
    Stop-ExistingRuntime

    $AutoUpdateBat = Join-Path $InstallDir "scripts\auto_update.bat"
    $TaskAction = "`"$AutoUpdateBat`" `"$Branch`" `"5`""

    $TaskArguments = @(
        "/Create",
        "/TN", "ExcelImageServer",
        "/SC", "ONSTART",
        "/TR", $TaskAction,
        "/RU", $script:TaskRunAs,
        "/RL", "HIGHEST",
        "/F"
    )
    if ($script:TaskRunAs -ne "SYSTEM") {
        $TaskArguments += @("/RP", $script:TaskPassword)
    }

    & schtasks.exe @TaskArguments
    if ($LASTEXITCODE -ne 0) {
        throw "Could not create scheduled task ExcelImageServer for $script:TaskRunAs."
    }
    $script:TaskPassword = $null

    $CreatedTask = Get-ScheduledTask -TaskName "ExcelImageServer"
    $CreatedTask.Settings.ExecutionTimeLimit = "PT0S"
    Set-ScheduledTask -InputObject $CreatedTask | Out-Null

    $SavedLimit = (
        Get-ScheduledTask -TaskName "ExcelImageServer"
    ).Settings.ExecutionTimeLimit
    if ($SavedLimit -ne "PT0S") {
        throw "Could not disable task execution limit. Current: $SavedLimit"
    }
}

function Start-AndVerifyService {
    $PidFile = Join-Path $InstallDir "logs\excel_image_server.pid"
    $DiagnosticLog = Join-Path $InstallDir "logs\supervisor-local.log"
    Remove-Item $PidFile, $DiagnosticLog -Force -ErrorAction SilentlyContinue

    schtasks /Run /TN "ExcelImageServer" | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Scheduled task was created but could not be started."
    }

    Write-Output "Verifying server process..."
    $Deadline = (Get-Date).AddSeconds(90)
    $WorkerIsRunning = $false
    $SupervisorConfirmedStart = $false

    while (
        (Get-Date) -lt $Deadline -and
        !($WorkerIsRunning -and $SupervisorConfirmedStart)
    ) {
        Start-Sleep -Seconds 2

        if (Test-Path $PidFile) {
            $SavedPid = 0
            if ([int]::TryParse((Get-Content $PidFile -Raw).Trim(), [ref]$SavedPid)) {
                $ProcessInfo = Get-CimInstance Win32_Process `
                    -Filter "ProcessId = $SavedPid" `
                    -ErrorAction SilentlyContinue
                $WorkerIsRunning = (
                    $null -ne $ProcessInfo -and
                    $ProcessInfo.CommandLine -like "*excel_image_server.py*"
                )
            }
        }

        if (Test-Path $DiagnosticLog) {
            $RecentLines = Get-Content $DiagnosticLog -Tail 20
            $SupervisorConfirmedStart = [bool](
                $RecentLines |
                    Where-Object { $_ -match "Excel image server is running" }
            )
            if (
                $RecentLines |
                    Where-Object { $_ -match "Python exited immediately" }
            ) {
                break
            }
        }
    }

    if (!$WorkerIsRunning -or !$SupervisorConfirmedStart) {
        $Details = "Local diagnostic log is empty."
        if (Test-Path $DiagnosticLog) {
            $Details = (Get-Content $DiagnosticLog -Tail 20) -join "`n"
        }
        throw "Server did not stay running. Diagnostic log: $DiagnosticLog`n$Details"
    }
}

function Show-Result {
    $Config = Get-Content (Join-Path $InstallDir "config.json") -Raw -Encoding UTF8 |
        ConvertFrom-Json
    $Port = if ($Config.port) { [int]$Config.port } else { 8091 }
    Write-Host ""
    Write-Host "INSTALL COMPLETE" -ForegroundColor Green
    Write-Host "Project: $InstallDir" -ForegroundColor Green
    Write-Host "Local URL: http://localhost:$Port" -ForegroundColor Green
    Write-Host "Public URL: http://<server-ip>:$Port" -ForegroundColor Green
    Write-Host "Installer log: $InstallLog" -ForegroundColor Green
    Write-Host ""
}

$Identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$Principal = New-Object Security.Principal.WindowsPrincipal($Identity)
$IsAdministrator = $Principal.IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)
if (!$IsAdministrator) {
    throw "Run PowerShell as Administrator."
}

if ($Repository -and !(Get-Command git -ErrorAction SilentlyContinue)) {
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Install-WingetPackage -PackageId "Git.Git" -DisplayName "Git"
    }
    else {
        Install-GitDirect
    }
}
elseif ($Repository) {
    Write-Output "Git is already installed."
}

if (!(Test-PythonInstalled)) {
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Install-WingetPackage -PackageId "Python.Python.3.13" -DisplayName "Python 3.13"
    }
    else {
        Install-PythonDirect
    }
}
else {
    Write-Output "Python is already installed."
}

Update-CurrentPath
Install-ProjectFiles
Set-Location $InstallDir

if (!(Test-Path "config.json")) {
    Copy-Item "config.example.json" "config.json"
}

Write-Output ""
Write-Output "Edit config.json and close Notepad to continue installation."
Start-Process notepad.exe -ArgumentList "`"$InstallDir\config.json`"" -Wait

Write-Output "Continuing installation..."
Configure-StorageAccess
Initialize-PythonEnvironment
Configure-Firewall
Create-ServerTask
Start-AndVerifyService
Show-Result

Write-Output "Full installation finished."

if ($TranscriptStarted) {
    Stop-Transcript | Out-Null
    $TranscriptStarted = $false
}

Read-Host "Press Enter to finish"
