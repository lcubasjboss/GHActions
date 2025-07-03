# This PowerShell script creates an Excel file with two worksheets.
# It also handles the installation of its required 'ImportExcel' module.
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
Write-Host "Checking/Registering PSGallery repository..."
try {
    # Check if PSGallery is already registered.
    if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
        Register-PSRepository -Default -InstallationPolicy Trusted -ErrorAction Stop
        Write-Host "PSGallery repository registered and trusted."
    } else {
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

# Force-reinstall PowerShellGet to ensure it's the latest and not corrupted.
# This is often the root cause of "No match was found" errors.
Write-Host "Attempting to ensure PowerShellGet is up-to-date and correctly installed..."
try {
    # Try to remove it first to ensure a clean install, ignore errors if it's not found or in use.
    Remove-Module -Name PowerShellGet -ErrorAction SilentlyContinue -Force
    Uninstall-Module -Name PowerShellGet -ErrorAction SilentlyContinue -Force

    # Install PowerShellGet with verbose logging
    Install-Module -Name PowerShellGet -Force -Scope CurrentUser -AllowClobber -Repository PSGallery -ErrorAction Stop -Verbose
    Write-Host "PowerShellGet installed/updated successfully."
}
catch {
    Write-Error "Failed to install/update PowerShellGet: $($_.Exception.Message)"
    exit 1 # Exit with an error code if installation fails
}

# Now, try to find and install ImportExcel.
Write-Host "Attempting to find and install ImportExcel module..."
try {
    # First, try to find the module to get diagnostic info.
    Write-Host "Searching for ImportExcel in PSGallery..."
    $foundModule = Find-Module -Name ImportExcel -Repository PSGallery -ErrorAction SilentlyContinue -Verbose

    if ($foundModule) {
        Write-Host "ImportExcel found in PSGallery. Version: $($foundModule.Version)"
        # If found, proceed with installation.
        Install-Module -Name ImportExcel -Force -Scope CurrentUser -Repository PSGallery -ErrorAction Stop -Verbose
        Write-Host "ImportExcel installed successfully."
    } else {
        Write-Error "ImportExcel module was NOT found in PSGallery. This indicates a deeper issue with the repository or network."
        exit 1 # Exit if the module cannot be found
    }
}
catch {
    Write-Error "Failed to install ImportExcel: $($_.Exception.Message)"
    exit 1 # Exit with an error code if installation fails
}

# Import the module for use in the current session.
Import-Module -Name ImportExcel -ErrorAction Stop
Write-Host "ImportExcel module loaded."

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
