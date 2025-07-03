# This PowerShell script creates an Excel file with two worksheets.
# It also handles the installation of its required 'Import-Excel' module.
# The second worksheet will contain the GitHub branch name and the short Git commit SHA.

param(
    [Parameter(Mandatory=$true)] # The environment name is a mandatory input.
    [string]$Environment,

    [Parameter(Mandatory=$true)] # The GitHub branch name is a mandatory input.
    [string]$BranchName,

    [Parameter(Mandatory=$true)] # The full Git commit SHA is a mandatory input.
    [string]$CommitSha
)

# --- Dependency Installation ---
Write-Host "Checking for and installing required PowerShell modules..."

# Ensure PSGallery is registered as a trusted repository.
# This often resolves issues where modules cannot be found.
Write-Host "Checking/Registering PSGallery repository..."
try {
    # Check if PSGallery is already registered.
    if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
        # If not registered, register it.
        Register-PSRepository -Default -InstallationPolicy Trusted -ErrorAction Stop
        Write-Host "PSGallery repository registered and trusted."
    } else {
        # If registered, ensure it's trusted.
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
        Write-Host "PSGallery repository found and set to trusted."
    }
}
catch {
    Write-Error "Failed to register/trust PSGallery repository: $($_.Exception.Message)"
    exit 1 # Exit with an error code if repository setup fails
}

# List repositories to confirm (for debugging purposes in logs)
Write-Host "Available PowerShell Repositories:"
Get-PSRepository | Format-Table -AutoSize

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

# --- Worksheet 2: Git Branch and Commit Info ---
# Get the short version of the Git commit SHA (first 7 characters).
$shortCommitSha = $CommitSha.Substring(0, [System.Math]::Min(7, $CommitSha.Length))

# Prepare the data for the second worksheet.
$gitInfoData = @(
    @{
        "GitHub Branch Name" = $BranchName;
        "Short Git Commit"   = $shortCommitSha
    }
)

Write-Host "Creating 'Git Info' worksheet..."
$gitInfoData | Export-Excel -Path $excelFilePath `
                           -WorksheetName "Git Info" `
                           -TableName "GitDetails" `
                           -TableStyle Light9 `
                           -AutoFit `
                           -Append

Write-Host "Excel file '$excelFilePath' created successfully and ready for upload."
