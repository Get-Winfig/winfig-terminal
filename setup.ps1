# ====================================================================== #
# UTF-8 with BOM Encoding for output
# ====================================================================== #

if ($PSVersionTable.PSVersion.Major -eq 5) {
    $OutputEncoding = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8
} else {
    $utf8WithBom = New-Object System.Text.UTF8Encoding $true
    $OutputEncoding = $utf8WithBom
    [Console]::OutputEncoding = $utf8WithBom
}

# ====================================================================== #
#  Script Metadata
# ====================================================================== #

$Script:WinfigMeta = @{
    Author       = "Armoghan-ul-Mohmin"
    CompanyName  = "Get-Winfig"
    Description  = "Windows configuration and automation framework"
    Version     = "1.0.0"
    License     = "MIT"
    Platform    = "Windows"
    PowerShell  = $PSVersionTable.PSVersion.ToString()
}

# ====================================================================== #
#  Color Palette
# ====================================================================== #

$Script:WinfigColors = @{
    Primary   = "Blue"
    Success   = "Green"
    Info      = "Cyan"
    Warning   = "Yellow"
    Error     = "Red"
    Accent    = "Magenta"
    Light     = "White"
    Dark      = "DarkGray"
}

# ====================================================================== #
# User Prompts
# ====================================================================== #

$Script:WinfigPrompts = @{
    Confirm    = "[?] Do you want to proceed? (Y/N): "
    Retry      = "[?] Do you want to retry? (Y/N): "
    Abort      = "[!] Operation aborted by user."
    Continue   = "[*] Press any key to continue..."
}

# ====================================================================== #
#  Paths
# ====================================================================== #

$Global:WinfigPaths = @{
    Desktop         = [Environment]::GetFolderPath("Desktop")
    Documents       = [Environment]::GetFolderPath("MyDocuments")
    UserProfile     = [Environment]::GetFolderPath("UserProfile")
    Temp            = [Environment]::GetEnvironmentVariable("TEMP")
    AppDataRoaming  = [Environment]::GetFolderPath("ApplicationData")
    AppDataLocal    = [Environment]::GetFolderPath("LocalApplicationData")
    Downloads       = [System.IO.Path]::Combine([Environment]::GetFolderPath("UserProfile"), "Downloads")
    Logs            = [System.IO.Path]::Combine([Environment]::GetEnvironmentVariable("TEMP"), "Winfig-Logs")
}
$Global:WinfigPaths.DotFiles = [System.IO.Path]::Combine($Global:WinfigPaths.UserProfile, ".Dotfiles")
$Global:WinfigPaths.Templates = [System.IO.Path]::Combine($Global:WinfigPaths.DotFiles, "winfig-terminal")

# ====================================================================== #
# Start Time, Resets, Counters
# ====================================================================== #
$Global:WinfigLogStart = Get-Date
$Global:WinfigLogFilePath = $null
Remove-Variable -Name WinfigLogFilePath -Scope Global -ErrorAction SilentlyContinue
Remove-Variable -Name LogCount -Scope Global -ErrorAction SilentlyContinue
Remove-Variable -Name ErrorCount -Scope Global -ErrorAction SilentlyContinue
Remove-Variable -Name WarnCount -Scope Global -ErrorAction SilentlyContinue
$Script:RemovedCount = 0
$Script:NotFoundCount = 0
$Script:ErrorCount = 0

# ====================================================================== #
# Utility Functions
# ====================================================================== #

# ---------------------------------------------------------------------------- #
# Function to display a Success message
function Show-SuccessMessage {
    param (
        [string]$Message
    )
    Write-Host "[OK] $Message" -ForegroundColor $Script:WinfigColors.Success
}

# ---------------------------------------------------------------------------- #
# Function to display an Error message
function Show-ErrorMessage {
    param (
        [string]$Message
    )
    Write-Host "[ERROR] $Message" -ForegroundColor $Script:WinfigColors.Error
}

# ---------------------------------------------------------------------------- #
# Function to display an Info message
function Show-InfoMessage {
    param (
        [string]$Message
    )
    Write-Host "[INFO] $Message" -ForegroundColor $Script:WinfigColors.Info
}

# ---------------------------------------------------------------------------- #
# Function to display a Warning message
function Show-WarningMessage {
    param (
        [string]$Message
    )
    Write-Host "[WARN] $Message" -ForegroundColor $Script:WinfigColors.Warning
}

# ---------------------------------------------------------------------------- #
# Function to prompt user for input with a specific color
function Prompt-UserInput {
    param (
        [string]$PromptMessage = $Script:WinfigPrompts.Confirm,
        [string]$PromptColor   = $Script:WinfigColors.Primary
    )
    # Write prompt in the requested color, keep cursor on same line, then read input
    Write-Host -NoNewline $PromptMessage -ForegroundColor $PromptColor
    $response = Read-Host

    return $response
}

# ---------------------------------------------------------------------------- #
# Function to Prompt user for confirmation (Y/N)
function Prompt-UserConfirmation {
    while ($true) {
        $response = Prompt-UserInput -PromptMessage $Script:WinfigPrompts.Confirm -PromptColor $Script:WinfigColors.Primary
        switch ($response.ToUpper()) {
            "Y" { return $true }
            "N" { return $false }
            default {
                Show-WarningMessage "Invalid input. Please enter Y or N."
            }
        }
    }
}

# ---------------------------------------------------------------------------- #
# Function to Prompt user to Retry (Y/N)
function Prompt-UserRetry {
    while ($true) {
        $response = Prompt-UserInput -PromptMessage $Script:WinfigPrompts.Retry -PromptColor $Script:WinfigColors.Primary
        switch ($response.ToUpper()) {
            "Y" { return $true }
            "N" { return $false }
            default {
                Show-WarningMessage "Invalid input. Please enter Y or N."
            }
        }
    }
}

# ---------------------------------------------------------------------------- #
# Function to Prompt user to continue
function Prompt-UserContinue {
    Write-Host $Script:WinfigPrompts.Continue -ForegroundColor $Script:WinfigColors.Primary
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ---------------------------------------------------------------------------- #
# Function to Abort operation
function Abort-Operation {
    Show-ErrorMessage $Script:WinfigPrompts.Abort
    # Write log footer before exiting
    if ($Global:WinfigLogFilePath) {
        Log-Message -Message "Script terminated." -EndRun
    }
    exit 1
}

# ---------------------------------------------------------------------------- #
# Function to Write a Section Header
function Write-SectionHeader {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Title,

        [Parameter(Mandatory=$false)]
        [string]$Description = ""
    )
    $separator = "=" * 70
    Write-Host $separator -ForegroundColor $Script:WinfigColors.Accent
    Write-Host "$Title" -ForegroundColor $Script:WinfigColors.Primary
    if ($Description) {
        Write-Host "$Description" -ForegroundColor $Script:WinfigColors.Accent
    }
    Write-Host $separator -ForegroundColor $Script:WinfigColors.Accent
}

# ---------------------------------------------------------------------------- #
# Function to Write a Subsection Header
function Write-SubsectionHeader {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Title
    )
    $separator = "-" * 50
    Write-Host $separator -ForegroundColor $Script:WinfigColors.Accent
    Write-Host "$Title" -ForegroundColor $Script:WinfigColors.Primary
    Write-Host $separator -ForegroundColor $Script:WinfigColors.Accent
}

# ---------------------------------------------------------------------------- #
#  Function to Write a Log Message
function Log-Message {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet("DEBUG", "INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO",

        [Parameter(Mandatory=$false)]
        [switch]$EndRun
    )

    if (-not $Global:LogCount) { $Global:LogCount = 0 }
    if (-not $Global:ErrorCount) { $Global:ErrorCount = 0 }
    if (-not $Global:WarnCount) { $Global:WarnCount = 0 }


    if (-not (Test-Path -Path $Global:WinfigPaths.Logs)) {
        New-Item -ItemType Directory -Path $Global:WinfigPaths.Logs -Force | Out-Null
    }

    $enc = New-Object System.Text.UTF8Encoding $true

    $identity = try { [System.Security.Principal.WindowsIdentity]::GetCurrent().Name } catch { $env:USERNAME }
    $isElevated = try {
        (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        $false
    }
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } else { $null }
    $psVersion = $PSVersionTable.PSVersion.ToString()
    $dotNetVersion = [System.Environment]::Version.ToString()
    $workingDir = (Get-Location).Path
    $osInfo = try {
        (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).Caption
    } catch {
        [Environment]::OSVersion.VersionString
    }
    # ---------------------------------------------------------------------------------------

    if (-not $Global:WinfigLogFilePath) {
        # $Global:WinfigLogStart is set in the main script execution block for each run
        $fileStamp = $Global:WinfigLogStart.ToString('yyyy-MM-dd_HH-mm-ss')
        $Global:WinfigLogFilePath = [System.IO.Path]::Combine($Global:WinfigPaths.Logs, "winfig-terminal-$fileStamp.log")

        $header = @()
        $header += "==================== Winfig Terminal Log ===================="
        $header += "Start Time  : $($Global:WinfigLogStart.ToString('yyyy-MM-dd HH:mm:ss'))"
        $header += "Host Name   : $env:COMPUTERNAME"
        $header += "User        : $identity"
        $header += "IsElevated  : $isElevated"
        if ($scriptPath) { $header += "Script Path : $scriptPath" }
        $header += "Working Dir : $workingDir"
        $header += "PowerShell  : $psVersion"
        $header += "NET Version : $dotNetVersion"
        $header += "OS          : $osInfo"
        $header += "=============================================================="
        $header += ""

        try {
            [System.IO.File]::WriteAllLines($Global:WinfigLogFilePath, $header, $enc)
        } catch {
            $header | Out-File -FilePath $Global:WinfigLogFilePath -Encoding UTF8 -Force
        }
    } else {
        if (-not $Global:WinfigLogStart) {
            $Global:WinfigLogStart = Get-Date
        }

        try {
            if (Test-Path -Path $Global:WinfigLogFilePath) {
                $firstLine = Get-Content -Path $Global:WinfigLogFilePath -TotalCount 1 -ErrorAction SilentlyContinue
                if ($firstLine -and ($firstLine -notmatch 'Winfig Terminal Log')) {

                    $header = @()
                    $header += "==================== Winfig Terminal Log  ===================="
                    $header += "Start Time  : $($Global:WinfigLogStart.ToString('yyyy-MM-dd HH:mm:ss'))"
                    $header += "Host Name   : $env:COMPUTERNAME"
                    $header += "User        : $identity"
                    $header += "IsElevated  : $isElevated"
                    if ($scriptPath) { $header += "Script Path : $scriptPath" }
                    $header += "Working Dir : $workingDir"
                    $header += "PowerShell  : $psVersion"
                    $header += "NET Version : $dotNetVersion"
                    $header += "OS          : $osInfo"
                    $header += "======================================================================="
                    $header += ""

                    # Prepend header safely: write header to temp file then append original content
                    $temp = [System.IO.Path]::GetTempFileName()
                    try {
                        [System.IO.File]::WriteAllLines($temp, $header, $enc)
                        [System.IO.File]::AppendAllLines($temp, (Get-Content -Path $Global:WinfigLogFilePath -Raw).Split([Environment]::NewLine), $enc)
                        Move-Item -Force -Path $temp -Destination $Global:WinfigLogFilePath
                    } finally {
                        if (Test-Path $temp) { Remove-Item $temp -ErrorAction SilentlyContinue }
                    }
                }
            }
        } catch {
            # ignore header-fix failures; continue logging
        }
    }

    if ($EndRun) {
        $endTime = Get-Date
        # $Global:WinfigLogStart is guaranteed to be set now
        $duration = $endTime - $Global:WinfigLogStart
        $footer = @()
        $footer += ""
        $footer += "--------------------------------------------------------------"
        $footer += "End Time    : $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))"
        $footer += "Duration    : $($duration.ToString('dd\.hh\:mm\:ss') -replace '^00\.', '')"
        $footer += "Log Count   : $Global:LogCount"
        $footer += "Errors/Warn : $Global:ErrorCount / $Global:WarnCount"
        $footer += "===================== End of Winfig Log ======================"
        try {
            [System.IO.File]::AppendAllLines($Global:WinfigLogFilePath, $footer, $enc)
        } catch {
            $footer | Out-File -FilePath $Global:WinfigLogFilePath -Append -Encoding UTF8
        }
        return
    }

    $now = Get-Date
    $timestamp = $now.ToString("yyyy-MM-dd HH:mm:ss.fff")
    $logEntry = "[$timestamp] [$Level] $Message"

    $Global:LogCount++
    if ($Level -eq 'ERROR') { $Global:ErrorCount++ }
    if ($Level -eq 'WARN') { $Global:WarnCount++ }

    try {
        [System.IO.File]::AppendAllText($Global:WinfigLogFilePath, $logEntry + [Environment]::NewLine, $enc)
    } catch {
        Write-Host "Failed to write log to file: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host $logEntry
    }
}

# ---------------------------------------------------------------------------- #
# Helper function to show category menu
function Show-CategoryMenu {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Category,

        [Parameter(Mandatory=$true)]
        [array]$Profiles
    )

    $selected = @()

    Write-Host ""
    Write-Host "┤ $Category ├" -ForegroundColor $Script:WinfigColors.Accent
    Write-Host ""

    # Display profiles with numbers
    for ($i = 0; $i -lt $Profiles.Count; $i++) {
        Write-Host "  $($i + 1). $($Profiles[$i].name)" -ForegroundColor $Script:WinfigColors.Light
    }

    Write-Host ""
    Show-InfoMessage "Category: $Category ($($Profiles.Count) profiles)"
    Show-InfoMessage "Enter 'all', 'none', or comma-separated numbers (e.g., 1,3,5)"

    $validInput = $false
    while (-not $validInput) {
        $selection = Prompt-UserInput -PromptMessage "[?] Select profiles from $($Category): " -PromptColor $Script:WinfigColors.Primary

        if ($selection.Trim() -eq 'all') {
            $selected = $Profiles
            Show-SuccessMessage "Selected all $($Profiles.Count) $Category profiles"
            $validInput = $true

        } elseif ($selection.Trim() -eq 'none') {
            Show-InfoMessage "Skipped $Category"
            $validInput = $true

        } elseif ($selection.Trim()) {
            # Parse comma-separated numbers
            $selectedIndices = $selection.Split(',') | ForEach-Object { $_.Trim() }
            $hasErrors = $false

            foreach ($indexStr in $selectedIndices) {
                if ([int]::TryParse($indexStr, [ref]$null)) {
                    $index = [int]$indexStr - 1
                    if ($index -ge 0 -and $index -lt $Profiles.Count) {
                        $selected += $Profiles[$index]
                        Write-Host "    [+] Selected: $($Profiles[$index].name)" -ForegroundColor $Script:WinfigColors.Success
                    } else {
                        Show-WarningMessage "    Invalid index: $indexStr (must be between 1 and $($Profiles.Count))"
                        $hasErrors = $true
                    }
                } else {
                    Show-WarningMessage "    Invalid input: '$indexStr' (not a number)"
                    $hasErrors = $true
                }
            }

            if (-not $hasErrors) {
                Show-SuccessMessage "Selected $($selected.Count) profiles from $Category"
                $validInput = $true
            } else {
                Show-WarningMessage "Please try again with valid numbers"
                $selected = @()  # Reset for retry
            }

        } else {
            Show-WarningMessage "Please enter 'all', 'none', or numbers"
        }
    }

    return $selected
}

# ---------------------------------------------------------------------------- #
# Helper function to process profile ordering
function Process-ProfileOrder {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Profiles,

        [Parameter(Mandatory=$true)]
        [string]$OrderInput
    )

    # Clean and split input
    $orderNumbers = $OrderInput.Split(',') | ForEach-Object { $_.Trim() }

    # Validate count matches
    if ($orderNumbers.Count -ne $Profiles.Count) {
        throw "You must specify exactly $($Profiles.Count) numbers (got $($orderNumbers.Count))"
    }

    # Validate all are numbers
    foreach ($num in $orderNumbers) {
        if (-not [int]::TryParse($num, [ref]$null)) {
            throw "'$num' is not a valid number"
        }
        $intNum = [int]$num
        if ($intNum -lt 1 -or $intNum -gt $Profiles.Count) {
            throw "Number $intNum is out of range (1-$($Profiles.Count))"
        }
    }

    # Check for duplicates
    $uniqueCount = ($orderNumbers | Select-Object -Unique).Count
    if ($uniqueCount -ne $Profiles.Count) {
        throw "Duplicate numbers found. Each position must be unique"
    }

    # Reorder profiles
    $reordered = @()
    foreach ($position in $orderNumbers) {
        $index = [int]$position - 1
        $reordered += $Profiles[$index]
    }

    return $reordered
}

# ====================================================================== #
#  Main Functions
# ====================================================================== #

# ---------------------------------------------------------------------------- #
# Function to check if running as Administrator
function IsAdmin{
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($currentUser)
    if ($principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Log-Message -Message "Script is running with Administrator privileges." -Level "SUCCESS"
    } else {
        Show-ErrorMessage "Script is NOT running with Administrator privileges."
        Log-Message -Message "Script is NOT running with Administrator privileges." -Level "ERROR"
        Log-Message "Forced exit." -EndRun
        $LogPathMessage = "Check the Log file for details: $($Global:WinfigLogFilePath)"
        Show-InfoMessage -Message $LogPathMessage
        exit 1
    }
}

# ---------------------------------------------------------------------------- #
# Function to check Working Internet Connection
function Test-InternetConnection {
    try {
        $request = [System.Net.WebRequest]::Create("http://www.google.com")
        $request.Timeout = 5000
        $response = $request.GetResponse()
        $response.Close()
        Log-Message -Message "Internet connection is available." -Level "SUCCESS"
        return $true
    } catch {
        Show-ErrorMessage "No internet connection available: $($_.Exception.Message)"
        Log-Message -Message "No internet connection available: $($_.Exception.Message)" -Level "ERROR"
        Log-Message "Forced exit." -EndRun
        $LogPathMessage = "Check the Log file for details: $($Global:WinfigLogFilePath)"
        Show-InfoMessage -Message $LogPathMessage
        exit 1

    }
}

# ---------------------------------------------------------------------------- #
# Function to check if PowerShell version is 7 or higher
function Test-PSVersion {
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -ge 7) {
        Log-Message -Message "PowerShell version is sufficient: $($psVersion.ToString())." -Level "SUCCESS"
    } else {
        Show-ErrorMessage "PowerShell version is insufficient: $($psVersion.ToString()). Version 7 or higher is required."
        Log-Message -Message "PowerShell version is insufficient: $($psVersion.ToString()). Version 7 or higher is required." -Level "ERROR"
        Log-Message "Forced exit." -EndRun
        $LogPathMessage = "Check the Log file for details: $($Global:WinfigLogFilePath)"
        Show-InfoMessage -Message $LogPathMessage
        exit 1
    }
}

# ---------------------------------------------------------------------------- #
# Function to Display Banner
function Winfig-Banner {
    Clear-Host
    Write-Host ""
    Write-Host ("  ██╗    ██╗██╗███╗   ██╗███████╗██╗ ██████╗  ".PadRight(70)) -ForegroundColor $Script:WinfigColors.Light
    Write-Host ("  ██║    ██║██║████╗  ██║██╔════╝██║██╔════╝  ".PadRight(70)) -ForegroundColor $Script:WinfigColors.Light
    Write-Host ("  ██║ █╗ ██║██║██╔██╗ ██║█████╗  ██║██║  ███╗ ".PadRight(70)) -ForegroundColor $Script:WinfigColors.Accent
    Write-Host ("  ██║███╗██║██║██║╚██╗██║██╔══╝  ██║██║   ██║ ".PadRight(70)) -ForegroundColor $Script:WinfigColors.Accent
    Write-Host ("  ╚███╔███╔╝██║██║ ╚████║██║     ██║╚██████╔╝ ".PadRight(70)) -ForegroundColor $Script:WinfigColors.Success
    Write-Host ("   ╚══╝╚══╝ ╚═╝╚═╝  ╚═══╝╚═╝     ╚═╝ ╚═════╝  ".PadRight(70)) -ForegroundColor $Script:WinfigColors.Success
    Write-Host ((" " * 70)) -ForegroundColor $Script:WinfigColors.Primary
    Write-Host ("" + $Script:WinfigMeta.CompanyName).PadLeft(40).PadRight(70) -ForegroundColor $Script:WinfigColors.Primary
    Write-Host ((" " * 70)) -ForegroundColor $Script:WinfigColors.Primary
    Write-Host ("  " + $Script:WinfigMeta.Description).PadRight(70) -ForegroundColor $Script:WinfigColors.Accent
    Write-Host ((" " * 70)) -ForegroundColor $Script:WinfigColors.Primary
    Write-Host (("  Version: " + $Script:WinfigMeta.Version + "    PowerShell: " + $Script:WinfigMeta.PowerShell).PadRight(70)) -ForegroundColor $Script:WinfigColors.Warning
    Write-Host (("  Author:  " + $Script:WinfigMeta.Author + "    Platform: " + $Script:WinfigMeta.Platform).PadRight(70)) -ForegroundColor $Script:WinfigColors.Warning
    Write-Host ""
}

# ---------------------------------------------------------------------------- #
# CTRL+C Signal Handler
trap {
    # Check if the error is due to a user interrupt (CTRL+C)
    if ($_.Exception.GetType().Name -eq "HostException" -and $_.Exception.Message -match "stopped by user") {

        # 1. Print the desired message
        Write-Host ""
        Write-Host ">>> [!] User interruption (CTRL+C) detected. Exiting gracefully..." -ForegroundColor $Script:WinfigColors.Accent

        # 2. Log the event before exit
        Log-Message -Message "Script interrupted by user (CTRL+C)." -Level "WARN"

        # 3. Write log footer before exiting
        if ($Global:WinfigLogFilePath) {
            Log-Message -Message "Script terminated by user (CTRL+C)." -EndRun
        }

        # 4. Terminate the script cleanly (exit code 1 is standard for non-zero exit)
        exit 1
    }
    # If it's a different kind of error, let the default behavior (or next trap) handle it
    continue
}

# ---------------------------------------------------------------------------- #
#  Check if windows terminal is installed or not
function Test-WindowsTerminalInstalled {
    $wtPath = "$($Global:WinfigPaths.AppDataLocal)\Microsoft\WindowsApps\wt.exe"
    if (Test-Path -Path $wtPath) {
        Log-Message -Message "Windows Terminal is installed." -Level "SUCCESS"
        return $true
    } else {
        Show-WarningMessage "Windows Terminal is not installed."
        Log-Message -Message "Windows Terminal is not installed." -Level "WARN"
        exit 1
    }
}

# ---------------------------------------------------------------------------- #
#  Create Dotfiles Directory if not exists
function Create-DotfilesDirectory {
    if (-not (Test-Path -Path $Global:WinfigPaths.DotFiles)) {
        try {
            New-Item -ItemType Directory -Path $Global:WinfigPaths.DotFiles -Force | Out-Null
            Show-SuccessMessage "Created Dotfiles directory at $($Global:WinfigPaths.DotFiles)."
            Log-Message -Message "Created Dotfiles directory at $($Global:WinfigPaths.DotFiles)." -Level "SUCCESS"
        } catch {
            Show-ErrorMessage "Failed to create Dotfiles directory: $($_.Exception.Message)"
            Log-Message -Message "Failed to create Dotfiles directory: $($_.Exception.Message)" -Level "ERROR"
            Abort-Operation
        }
    } else {
        Log-Message -Message "Dotfiles directory already exists at $($Global:WinfigPaths.DotFiles)." -Level "INFO"
    }
}

# ---------------------------------------------------------------------------- #
#  Check if git is installed
function Test-GitInstalled {
    try {
        git --version *> $null
        Log-Message -Message "Git is installed." -Level "SUCCESS"
        return $true
    } catch {
        Show-ErrorMessage "Git is not installed or not found in PATH."
        Log-Message -Message "Git is not installed or not found in PATH." -Level "ERROR"
        exit 1
    }
}

# ---------------------------------------------------------------------------- #
# Create Config File for Windows terminal.
function Create-ConfigJsonFile {
    $jsonFilePath = Join-Path -Path $Global:WinfigPaths.Templates -ChildPath "settings.json"
    if (Test-Path -Path $jsonFilePath) {
        Move-Item -Path $jsonFilePath -Destination "$jsonFilePath.bak" -Force
        Log-Message -Message "Backed up existing JSON file at $jsonFilePath.bak." -Level "INFO"
    }else {
        New-Item -ItemType File -Path $jsonFilePath -Force | Out-Null
        Log-Message -Message "Created new JSON file at $jsonFilePath." -Level "INFO"
    }
}

# ---------------------------------------------------------------------------- #
#  Read the content of "settings-base.json" and copy it to newly created "settings.json"
function Copy-BaseSettings {
    $baseJsonPath     = Join-Path $Global:WinfigPaths.Templates "settings-base.json"
    $modifiedJsonPath = Join-Path $Global:WinfigPaths.Templates "settings.json"

    if (-not (Test-Path $baseJsonPath)) {
        Show-ErrorMessage "Base settings file not found at $baseJsonPath."
        Log-Message -Message "Base settings file not found at $baseJsonPath." -Level "ERROR"
        exit 1
    }

    Copy-Item -Path $baseJsonPath -Destination $modifiedJsonPath -Force
}

# ---------------------------------------------------------------------------- #
# Read the content of "actions.json" file and then copy them into its appropriate place in "settings.json"
function Copy-ActionsSettings {
    try {
        $actionsPath = Join-Path $Global:WinfigPaths.Templates "actions.json"
        $settingsPath = Join-Path $Global:WinfigPaths.Templates "settings.json"
        $tempActionsPath = Join-Path $Global:WinfigPaths.Temp "actions-clean.json"

        # Step 1: Remove comment lines and write to temp file
        Get-Content $actionsPath | Where-Object { -not ($_ -match '^\s*//') } | Set-Content $tempActionsPath -Encoding UTF8

        # Step 2: Read temp file, remove trailing commas, and parse JSON
        $json = (Get-Content $tempActionsPath -Raw) -replace ',(\s*[\]\}])', '$1'
        $actions = $json | ConvertFrom-Json | Select-Object -ExpandProperty actions

        # Step 3: Read settings, update actions, and save
        $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json
        $settings.actions = $actions
        $settings | ConvertTo-Json -Depth 10 | Set-Content $settingsPath -Encoding UTF8

        # Cleanup temp file
        Remove-Item $tempActionsPath -ErrorAction SilentlyContinue

        Show-SuccessMessage "Copied $($actions.Count) actions to settings.json"
        Log-Message -Message "Copied $($actions.Count) actions to settings.json" -Level "SUCCESS"

    } catch {
        Show-ErrorMessage "Failed to copy actions: $($_.Exception.Message)"
        Log-Message -Message "Failed to copy actions: $($_.Exception.Message)" -Level "ERROR"
        exit 1
    }
}

# ---------------------------------------------------------------------------- #
# Read the content of "schemes.json" file and then ask user what schemes they wanna add in config then copy only user required schemes in "settings.json" file in appropriate location
function Copy-SchemesSettings {
    try {
        $schemesPath = Join-Path $Global:WinfigPaths.Templates "schemes.json"
        $settingsPath = Join-Path $Global:WinfigPaths.Templates "settings.json"

        # Read schemes from JSON
        $schemesJson = Get-Content $schemesPath -Raw | ConvertFrom-Json
        $allSchemes = $schemesJson.schemes

        # Display available schemes
        Write-SubsectionHeader -Title "Available Color Schemes"
        Write-Host ""

        # Create numbered list of schemes
        $schemeList = @()
        for ($i = 0; $i -lt $allSchemes.Count; $i++) {
            $schemeName = $allSchemes[$i].name
            $schemeList += @{Index = $i + 1; Name = $schemeName}
            Write-Host "$($i + 1). $schemeName" -ForegroundColor $Script:WinfigColors.Info
        }

        Write-Host ""
        Write-Host "You can select multiple schemes by entering numbers separated by commas (e.g., 1,3,5)" -ForegroundColor $Script:WinfigColors.Warning
        Write-Host "Enter 'all' to select all schemes" -ForegroundColor $Script:WinfigColors.Warning
        Write-Host "Enter 'none' to skip adding schemes" -ForegroundColor $Script:WinfigColors.Warning
        Write-Host ""

        # Prompt user for selection
        $selection = Prompt-UserInput -PromptMessage "[?] Select schemes to add: " -PromptColor $Script:WinfigColors.Primary

        # Process selection
        $selectedSchemes = @()

        if ($selection.Trim() -eq 'all') {
            # Select all schemes
            $selectedSchemes = $allSchemes
            Show-InfoMessage "Selected all $($allSchemes.Count) schemes"
        }
        elseif ($selection.Trim() -eq 'none') {
            # No schemes selected
            Show-InfoMessage "No schemes selected"
            $selectedSchemes = @()
        }
        else {
            # Parse comma-separated numbers
            $selectedIndices = $selection.Split(',') | ForEach-Object { $_.Trim() }

            foreach ($indexStr in $selectedIndices) {
                if ([int]::TryParse($indexStr, [ref]$null)) {
                    $index = [int]$indexStr - 1
                    if ($index -ge 0 -and $index -lt $allSchemes.Count) {
                        $selectedSchemes += $allSchemes[$index]
                        Write-Host "  [+] Selected: $($allSchemes[$index].name)" -ForegroundColor $Script:WinfigColors.Success
                    } else {
                        Write-Host "  [!] Invalid index: $indexStr (out of range)" -ForegroundColor $Script:WinfigColors.Error
                    }
                } else {
                    Write-Host "  [!] Invalid input: $indexStr (not a number)" -ForegroundColor $Script:WinfigColors.Error
                }
            }
        }

        if ($selectedSchemes.Count -gt 0) {
            # Read current settings
            $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json

            # Add or update schemes in settings
            $settings.schemes = $selectedSchemes

            # Save updated settings
            $settings | ConvertTo-Json -Depth 10 | Set-Content $settingsPath -Encoding UTF8

            Show-SuccessMessage "Added $($selectedSchemes.Count) color schemes to settings.json"
            Log-Message -Message "Added $($selectedSchemes.Count) color schemes to settings.json" -Level "SUCCESS"

            # Show selected scheme names
            Write-Host ""
            Write-Host "Selected schemes:" -ForegroundColor $Script:WinfigColors.Info
            foreach ($scheme in $selectedSchemes) {
                Write-Host "  • $($scheme.name)" -ForegroundColor $Script:WinfigColors.Light
            }
        } else {
            Show-InfoMessage "No schemes were added to settings.json"
            Log-Message -Message "No schemes were added to settings.json" -Level "INFO"
        }

        Write-Host ""

    } catch {
        $errorMsg = "Failed to copy schemes: $($_.Exception.Message)"
        Show-ErrorMessage $errorMsg
        Log-Message -Message $errorMsg -Level "ERROR"
    }
}

# ---------------------------------------------------------------------------- #
#  Read the content of "profiles.json" file and then ask user what profile they wants and in what order they wants their profiles then configure them in "settings.json" file in appropriate location
function Copy-ProfileSettings {
    try {
        $profilesPath = Join-Path $Global:WinfigPaths.Templates "profiles.json"
        $settingsPath = Join-Path $Global:WinfigPaths.Templates "settings.json"

        # Read profiles from JSON
        $profilesJson = Get-Content $profilesPath -Raw | ConvertFrom-Json
        $defaults = $profilesJson.profiles.defaults
        $allProfiles = $profilesJson.profiles.list

        # Display available profiles count
        Show-InfoMessage "Found $($allProfiles.Count) available profiles"
        Log-Message -Message "Found $($allProfiles.Count) available profiles" -Level "INFO"

        # --------------------------------------------------------------------
        # Selection Process
        # --------------------------------------------------------------------
        Write-SubsectionHeader -Title "PROFILE SELECTION PROCESS"

        # Show selection instructions
        Show-InfoMessage "You will now select which profiles to include in your terminal."
        Show-InfoMessage "For each category, you can select 'all', 'none', or specific numbers."
        Write-Host ""

        # --------------------------------------------------------------------
        # Create selection menus
        # --------------------------------------------------------------------
        $selectedProfiles = @()

        # Menu 1: Basic Profiles (PowerShell, CMD, Git Bash)
        $basicProfiles = $allProfiles | Where-Object {
            $_.name -match "PowerShell|Command Prompt|Git Bash"
        }
        if ($basicProfiles.Count -gt 0) {
            $basicSelected = Show-CategoryMenu -Category "Basic Profiles" -Profiles $basicProfiles
            $selectedProfiles += $basicSelected
        }

        # Menu 2: WSL Profiles
        $wslProfiles = $allProfiles | Where-Object {
            $_.commandline -and $_.commandline -match '^wsl -d '
        }
        if ($wslProfiles.Count -gt 0) {
            $wslSelected = Show-CategoryMenu -Category "WSL Distributions" -Profiles $wslProfiles
            $selectedProfiles += $wslSelected
        }

        # Menu 3: Distrobox Profiles
        $distroboxProfiles = $allProfiles | Where-Object {
            $_.commandline -and $_.commandline -match '^distrobox-enter '
        }
        if ($distroboxProfiles.Count -gt 0) {
            $distroboxSelected = Show-CategoryMenu -Category "Distrobox Templates" -Profiles $distroboxProfiles
            $selectedProfiles += $distroboxSelected
        }

        # Menu 4: Other Profiles
        $otherProfiles = $allProfiles | Where-Object {
            $selectedProfiles -notcontains $_ -and
            $basicProfiles -notcontains $_ -and
            $wslProfiles -notcontains $_ -and
            $distroboxProfiles -notcontains $_
        }
        if ($otherProfiles.Count -gt 0) {
            $otherSelected = Show-CategoryMenu -Category "Other Profiles" -Profiles $otherProfiles
            $selectedProfiles += $otherSelected
        }

        # Check if any profiles were selected
        if ($selectedProfiles.Count -eq 0) {
            Show-WarningMessage "No profiles were selected. Using default profiles."
            Log-Message -Message "No profiles selected by user" -Level "WARN"
            $selectedProfiles = @($allProfiles[0], $allProfiles[1], $allProfiles[3])  # PowerShell Core, Windows PowerShell, Command Prompt
        }

        # --------------------------------------------------------------------
        # Ordering Process
        # --------------------------------------------------------------------
        Write-SubsectionHeader -Title "PROFILE ORDERING"

        # Show current order
        Show-InfoMessage "Current order of selected profiles:"
        for ($i = 0; $i -lt $selectedProfiles.Count; $i++) {
            Write-Host "  $($i + 1). $($selectedProfiles[$i].name)" -ForegroundColor $Script:WinfigColors.Light
        }

        # Ask if user wants to reorder
        Write-Host ""
        $reorderResponse = Prompt-UserInput -PromptMessage "[?] Would you like to reorder these profiles? (Y/N): " -PromptColor $Script:WinfigColors.Primary

        if ($reorderResponse.Trim().ToUpper() -eq 'Y') {
            Show-InfoMessage "Enter the new order using the numbers above, separated by commas."
            Show-InfoMessage "Example: To swap first and second profiles, enter: 2,1,3,4,..."

            $orderInput = Prompt-UserInput -PromptMessage "[?] Enter new order: " -PromptColor $Script:WinfigColors.Primary

            if ($orderInput.Trim()) {
                try {
                    $newOrder = Process-ProfileOrder -Profiles $selectedProfiles -OrderInput $orderInput
                    if ($newOrder) {
                        $selectedProfiles = $newOrder
                        Show-SuccessMessage "Profiles reordered successfully"
                        Log-Message -Message "User reordered $($selectedProfiles.Count) profiles" -Level "SUCCESS"
                    }
                } catch {
                    Show-WarningMessage "Could not reorder profiles: $($_.Exception.Message)"
                    Show-InfoMessage "Keeping original order."
                    Log-Message -Message "Failed to reorder profiles: $($_.Exception.Message)" -Level "WARN"
                }
            }
        }

        # --------------------------------------------------------------------
        # Final Confirmation
        # --------------------------------------------------------------------
        Write-SubsectionHeader -Title "FINAL CONFIGURATION"

        Show-InfoMessage "Final profile configuration:"
        for ($i = 0; $i -lt $selectedProfiles.Count; $i++) {
            Write-Host "  $($i + 1). $($selectedProfiles[$i].name)" -ForegroundColor $Script:WinfigColors.Success
        }
        Write-Host ""

        # Ask for final confirmation
        if (Prompt-UserConfirmation) {
            # Read current settings
            $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json

            # Update profiles in settings
            $settings.profiles = @{
                defaults = $defaults
                list = $selectedProfiles
            }

            # Save updated settings
            $settings | ConvertTo-Json -Depth 10 | Set-Content $settingsPath -Encoding UTF8

            Show-SuccessMessage "Successfully configured $($selectedProfiles.Count) profiles in settings.json"
            Log-Message -Message "Configured $($selectedProfiles.Count) profiles in settings.json" -Level "SUCCESS"

        } else {
            Show-InfoMessage "Profile configuration cancelled by user"
            Log-Message -Message "Profile configuration cancelled by user" -Level "INFO"
        }

    } catch {
        $errorMsg = "Failed to configure profiles: $($_.Exception.Message)"
        Show-ErrorMessage $errorMsg
        Log-Message -Message $errorMsg -Level "ERROR"
    }
}

# ====================================================================== #
#  Main Script Execution
# ====================================================================== #

Winfig-Banner
Write-SectionHeader -Title "Checking Requirements"
Write-Host ""

IsAdmin | Out-Null
Show-SuccessMessage "Administrator privileges confirmed."

Test-InternetConnection | Out-Null
Show-SuccessMessage "Internet connection is available."

Test-PSVersion | Out-Null
Show-SuccessMessage "PowerShell version is sufficient."

Test-WindowsTerminalInstalled | Out-Null
Show-SuccessMessage "Windows Terminal installation check completed."

Test-GitInstalled | Out-Null
Show-SuccessMessage "Git installation check completed."

Create-DotfilesDirectory | Out-Null
Show-SuccessMessage "Dotfiles directory setup completed."
Write-Host ""
Prompt-UserContinue

Winfig-Banner
Write-SectionHeader -Title "Cloning Winfig Terminal Repository"
Write-Host ""
$repoPath = Join-Path $Global:WinfigPaths.DotFiles "winfig-terminal"
if (-not (Test-Path -Path $repoPath)) {
    try {
        Show-InfoMessage "Cloning Winfig Terminal repository..."
        Log-Message -Message "Cloning Winfig Terminal repository..." -Level "INFO"
        git clone https://github.com/Get-Winfig/winfig-terminal.git $repoPath *> $null
    } catch {
        Show-ErrorMessage "Failed to clone Winfig Terminal repository: $($_.Exception.Message)"
        Log-Message -Message "Failed to clone Winfig Terminal repository: $($_.Exception.Message)" -Level "ERROR"
        exit 1
    }
    if (Test-Path -Path $repoPath) {
        Show-SuccessMessage "Cloned Winfig Terminal repository to $repoPath."
        Log-Message -Message "Cloned Winfig Terminal repository to $repoPath." -Level "SUCCESS"
    } else {
        Show-ErrorMessage "Winfig Terminal repository was not cloned. Please check your internet connection or repository URL."
        Log-Message -Message "Winfig Terminal repository was not cloned. Please check your internet connection or repository URL." -Level "ERROR"
        exit 1
    }
} else {
    try {
        Show-InfoMessage "Updating Winfig Terminal repository..."
        Log-Message -Message "Updating Winfig Terminal repository..." -Level "INFO"
        Push-Location $repoPath
        git pull *> $null
        Pop-Location
        Show-SuccessMessage "Updated Winfig Terminal repository at $repoPath."
        Log-Message -Message "Updated Winfig Terminal repository at $repoPath." -Level "SUCCESS"
    } catch {
        Show-ErrorMessage "Failed to update Winfig Terminal repository: $($_.Exception.Message)"
        Log-Message -Message "Failed to update Winfig Terminal repository: $($_.Exception.Message)" -Level "ERROR"
        exit 1
    }
}
Write-Host ""
Prompt-UserContinue

Winfig-Banner
Write-SectionHeader -Title "Download Wallpaper for Terminal Background"
Write-Host ""

$wallpaperUrl = "https://raw.githubusercontent.com/Get-Winfig/winfig-wallpaper/refs/heads/main/images/003-3904x2240.png"
$wallpaperPath = "C:\Windows\Web\Wallpaper\Terminal.png"
try {
    Invoke-WebRequest -Uri $wallpaperUrl -OutFile $wallpaperPath -ErrorAction Stop
    Show-SuccessMessage "Downloaded wallpaper to $wallpaperPath."
    Log-Message -Message "Downloaded wallpaper to $wallpaperPath." -Level "SUCCESS"
} catch {
    Show-ErrorMessage "Failed to download wallpaper: $($_.Exception.Message)"
    Log-Message -Message "Failed to download wallpaper: $($_.Exception.Message)" -Level "ERROR"
}
Write-Host ""
Prompt-UserContinue

Winfig-Banner
Write-SectionHeader -Title "Creating Windows Terminal Configuration"
Write-Host ""
Create-ConfigJsonFile | Out-Null
Show-SuccessMessage "Created Settings.json"
Write-Host ""
Copy-BaseSettings | Out-Null
Show-SuccessMessage "Copied base settings to Settings.json."
Write-Host ""
Copy-ActionsSettings | Out-Null
Show-SuccessMessage "Copied actions to Settings.json."
Prompt-UserContinue

Winfig-Banner
Write-SectionHeader -Title "Configuring Color Schemes"
Write-Host ""
Copy-SchemesSettings | Out-Null
Show-SuccessMessage "Schemes configuration completed."
Prompt-UserContinue

Winfig-Banner
Write-SectionHeader -Title "Configuring Terminal Profiles"
Write-Host ""
Copy-ProfileSettings | Out-Null
Show-SuccessMessage "Profiles configuration completed."
Prompt-UserContinue

Winfig-Banner
Write-SectionHeader -Title "Symlink the config file"
Write-Host ""
try {
    $source = Join-Path $Global:WinfigPaths.Templates "settings.json"
    $target = Join-Path $Global:WinfigPaths.AppDataLocal "Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"

    if (Test-Path $target) { Remove-Item $target -Force }
    New-Item -ItemType SymbolicLink -Path $target -Target $source -Force

    Show-SuccessMessage "Symlink created: $target -> $source"
    Log-Message -Message "Symlink created: $target -> $source" -Level "SUCCESS"

} catch {
    Show-ErrorMessage "Failed to create symlink: $($_.Exception.Message)"
    Log-Message -Message "Failed to create symlink: $($_.Exception.Message)" -Level "ERROR"
}
Write-Host ""
Write-SectionHeader -Title "Thank You For Using Winfig Terminal" -Description "https://github.com/Get-Winfig/"
Show-WarningMessage -Message "Restart Windows to apply changes"
Write-Host ""
Log-Message -Message "Logging Completed." -EndRun
