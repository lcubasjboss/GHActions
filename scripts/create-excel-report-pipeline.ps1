# This PowerShell script creates an Excel file with two worksheets.
# It handles the installation of its required 'ImportExcel' module (leveraging caching for speed).
# The first worksheet includes environment and branch info.
# The second worksheet includes pipeline run number and short Git commit.

param(
    [Parameter(Mandatory=$true)] # The environment name is a mandatory input.
    [string]$Environment,

    [Parameter(Mandatory=$true)] # The GitHub branch name is a mandatory input.
    [string]$BranchName,

    [Parameter(Mandatory=$true)] # The full Git commit SHA is a mandatory input.
    [string]$CommitSha,

    [Parameter(Mandatory=$true)] # The GitHub Actions run number is a mandatory input.
    [int]$RunNumber
)

# --- Dependency Installation ---
Write-Host "Checking for and installing required PowerShell modules..."

# Ensure PSGallery is registered as a trusted repository.
Write-Host "Checking/Registering PSGallery repository..."
try {
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

# Function to safely install/update a PowerShell module
function Install-ModuleIfMissing {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModuleName
    )

    Write-Host "Checking module: $ModuleName"
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "$ModuleName module not found. Installing..."
        try {
            # Try to remove it first in case of partial/corrupted install
            Remove-Module -Name $ModuleName -ErrorAction SilentlyContinue -Force
            Uninstall-Module -Name $ModuleName -ErrorAction SilentlyContinue -Force

            Install-Module -Name $ModuleName -Force -Scope CurrentUser -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-Host "$ModuleName installed successfully."
        }
        catch {
            Write-Error "Failed to install/update $ModuleName : $($_.Exception.Message)"
            exit 1 # Exit with an error code if installation fails
        }
    } else {
        Write-Host "$ModuleName module found."
    }
}

# Install/update required modules
Install-ModuleIfMissing -ModuleName PackageManagement
Install-ModuleIfMissing -ModuleName PowerShellGet
Install-ModuleIfMissing -ModuleName ImportExcel

# Import the module for use in the current session.
Import-Module -Name ImportExcel -ErrorAction Stop
Write-Host "ImportExcel module loaded."

# Define the name and path for the output Excel file.
$excelFilePath = "environment_report.xlsx"

Write-Host "Starting Excel file creation at $excelFilePath..."

# --- Worksheet 1: Environment and Branch Info ---
# Get the current timestamp when this script is executed.
$scriptExecutionTimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Prepare the data for the first worksheet.
$envBranchData = @(
    @{
        "Environment Name"   = $Environment;
        "Branch Used"        = $BranchName;
        "Timestamp"          = $scriptExecutionTimeStamp
    }
)

Write-Host "Creating 'Environment & Branch Info' worksheet..."
$envBranchData | Export-Excel -Path $excelFilePath `
                             -WorksheetName "Environment & Branch Info" `
                             -TableName "EnvBranchDetails" `
                             -TableStyle Light9 `
                             -ClearSheet `
                             -AutoSize

# --- Worksheet 2: Pipeline Run and Git Commit Info ---
# Get the short version of the Git commit SHA (first 7 characters).
$shortCommitSha = $CommitSha.Substring(0, [System.Math]::Min(7, $CommitSha.Length))

# Prepare the data for the second worksheet.
$pipelineGitData = @(
    @{
        "Pipeline Run Number" = $RunNumber;
        "Short Git Commit"    = $shortCommitSha
    }
)

Write-Host "Creating 'Pipeline & Git Info' worksheet..."
$pipelineGitData | Export-Excel -Path $excelFilePath `
                               -WorksheetName "Pipeline & Git Info" `
                               -TableName "PipelineGitDetails" `
                               -TableStyle Light9 `
                               -AutoSize `
                               -Append

Write-Host "Excel file '$excelFilePath' created successfully and ready for upload."
