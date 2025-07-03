# This PowerShell script creates an Excel file with two worksheets.
# It also handles the installation of its required 'Import-Excel' module
# and calculates its own execution duration.

param(
    [Parameter(Mandatory=$true)] # The environment name is a mandatory input.
    [string]$Environment
)

# Record the start time of the script execution for duration calculation.
$scriptStartTime = Get-Date

# --- Dependency Installation ---
Write-Host "Checking for and installing required PowerShell modules..."

# Check if PowerShellGet is installed and up to date.
if (-not (Get-Module -ListAvailable -Name PowerShellGet)) {
    Write-Host "PowerShellGet module not found. Installing..."
    try {
        Install-Module -Name PowerShellGet -Force -Scope CurrentUser -AllowClobber -Repository PSGallery -ErrorAction Stop
        Write-Host "PowerShellGet installed successfully."
    }
    catch {
        Write-Error "Failed to install PowerShellGet: $($_.Exception.Message)"
        exit 1 # Exit with an error code if installation fails
    }
} else {
    Write-Host "PowerShellGet module found."
}

# Check if Import-Excel is installed.
if (-not (Get-Module -ListAvailable -Name Import-Excel)) {
    Write-Host "Import-Excel module not found. Installing..."
    try {
        Install-Module -Name Import-Excel -Force -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
        Write-Host "Import-Excel installed successfully."
    }
    catch {
        Write-Error "Failed to install Import-Excel: $($_.Exception.Message)"
        exit 1 # Exit with an error code if installation fails
    }
} else {
    Write-Host "Import-Excel module found."
}

# Import the module for use in the current session.
Import-Module -Name Import-Excel -ErrorAction Stop
Write-Host "Import-Excel module loaded."

# Define the name and path for the output Excel file.
$excelFilePath = "environment_report.xlsx"

Write-Host "Starting Excel file creation at $excelFilePath..."

# --- Worksheet 1: Environment Info ---
# Get the current timestamp when this script is executed.
$scriptExecutionTimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Prepare the data for the first worksheet.
$envData = @(
    @{
        "Name of the environment"        = $Environment;
        "TimeStamp of the execution" = $scriptExecutionTimeStamp
    }
)

Write-Host "Creating 'Environment Info' worksheet..."
$envData | Export-Excel -Path $excelFilePath `
                       -WorksheetName "Environment Info" `
                       -TableName "EnvironmentDetails" `
                       -TableStyle Light9 `
                       -ClearSheet `
                       -AutoFit

# --- Worksheet 2: Script Execution Duration ---
# Calculate the duration from the script's start time to now.
$scriptEndTime = Get-Date
$durationSeconds = ($scriptEndTime - $scriptStartTime).TotalSeconds
$durationString = "$([math]::Round($durationSeconds, 2)) seconds"

# Prepare the data for the second worksheet.
$durationData = @(
    @{ "Duration of the PowerShell script execution (seconds)" = $durationString }
)

Write-Host "Creating 'Script Execution Duration' worksheet..."
$durationData | Export-Excel -Path $excelFilePath `
                           -WorksheetName "Script Execution Duration" `
                           -TableName "ScriptDurationDetails" `
                           -TableStyle Light9 `
                           -AutoFit `
                           -Append

Write-Host "Excel file '$excelFilePath' created successfully and ready for upload."
